import zipfile
import tempfile
import os
import pyogrio
import geopandas as gpd
import pandas as pd
from openpyxl.styles import PatternFill, Font, Alignment

# ==========================================
# FUNÇÕES AUXILIARES
# ==========================================
def abrir_kml_ou_kmz(caminho):
    def ler_todas_camadas(arquivo_kml):
        camadas = pyogrio.list_layers(arquivo_kml)
        lista_gdfs = []
        for info_camada in camadas:
            nome_camada = info_camada[0]
            try:
                gdf = gpd.read_file(arquivo_kml, engine="pyogrio", layer=nome_camada)
                if not gdf.empty:
                    lista_gdfs.append(gdf)
            except Exception:
                continue
        if not lista_gdfs:
            raise Exception(f"Nenhum dado geográfico válido encontrado em: {caminho}")
        return pd.concat(lista_gdfs, ignore_index=True)

    if caminho.lower().endswith(".kml"):
        return ler_todas_camadas(caminho)
    elif caminho.lower().endswith(".kmz"):
        with tempfile.TemporaryDirectory() as tmp:
            with zipfile.ZipFile(caminho, 'r') as z:
                z.extractall(tmp)
            kml_extraido = None
            for raiz, _, arquivos in os.walk(tmp):
                for arq in arquivos:
                    if arq.lower().endswith(".kml"):
                        kml_extraido = os.path.join(raiz, arq)
                        break
            if kml_extraido is None:
                raise Exception("Nenhum arquivo KML encontrado dentro do KMZ.")
            return ler_todas_camadas(kml_extraido)
    else:
        raise ValueError("O formato do arquivo precisa ser .kml ou .kmz")

def ler_clientes_unico_csv(caminho):
    if caminho.lower().endswith('.xlsx') or caminho.lower().endswith('.xls'):
        df = pd.read_excel(caminho, dtype=str)
    else:
        configuracoes = [
            {'encoding': 'utf-16', 'sep': '\t'},
            {'encoding': 'utf-16', 'sep': ';'},
            {'encoding': 'utf-16', 'sep': ','},
            {'encoding': 'utf-8',  'sep': ';'},
            {'encoding': 'utf-8',  'sep': ','},
            {'encoding': 'latin1', 'sep': ';'},
            {'encoding': 'latin1', 'sep': ','}
        ]
        df = None
        for config in configuracoes:
            try:
                df_temp = pd.read_csv(caminho, sep=config['sep'], encoding=config['encoding'], dtype=str)
                if 'latitude' in df_temp.columns and 'longitude' in df_temp.columns:
                    df = df_temp
                    break
            except Exception:
                continue
        if df is None:
            df = pd.read_csv(caminho, encoding='utf-16', sep=None, engine='python', dtype=str)

    for col in ['latitude', 'longitude']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace('"', '').str.strip().str.replace(',', '.')
            df[col] = pd.to_numeric(df[col], errors='coerce')
    
    df = df.dropna(subset=['latitude', 'longitude'])
    gdf = gpd.GeoDataFrame(
        df, 
        geometry=gpd.points_from_xy(df['longitude'], df['latitude']),
        crs="EPSG:4326"
    )
    return gdf

def ler_multiplos_poligonos(lista_caminhos):
    gdfs = []
    for cam in lista_caminhos:
        gdfs.append(abrir_kml_ou_kmz(cam))
    if not gdfs:
        raise Exception("Nenhum dado de polígono válido pôde ser extraído.")
    return pd.concat(gdfs, ignore_index=True)

def ler_multiplas_bases_clientes(lista_caminhos):
    gdfs = []
    for cam in lista_caminhos:
        gdfs.append(ler_clientes_unico_csv(cam))
    if not gdfs:
        raise Exception("Nenhuma base de clientes válida pôde ser lida.")
    return pd.concat(gdfs, ignore_index=True)

def salvar_e_formatar_excel(df, caminho, nome_planilha="Planilha1"):
    with pd.ExcelWriter(caminho, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=nome_planilha)
        worksheet = writer.sheets[nome_planilha]
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="left", vertical="center")
        worksheet.auto_filter.ref = worksheet.dimensions
        for col in worksheet.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except: pass
            worksheet.column_dimensions[col_letter].width = max_length + 3

# ==========================================
# MOTOR PRINCIPAL
# ==========================================
def executar_automacao_ccp_backend(lista_arquivos_poligonos, lista_arquivos_clientes, callback_progresso):
    
    callback_progresso(10, f"Lendo {len(lista_arquivos_poligonos)} arquivo(s) de polígonos...")
    poligonos = ler_multiplos_poligonos(lista_arquivos_poligonos)
    
    callback_progresso(30, f"Lendo {len(lista_arquivos_clientes)} base(s) de sucursais escolhidas...")
    clientes = ler_multiplas_bases_clientes(lista_arquivos_clientes)
    
    callback_progresso(50, "Padronizando sistemas e limpando geometrias...")
    poligonos = poligonos.to_crs(4326)
    if hasattr(poligonos.geometry, 'force_2d'): poligonos.geometry = poligonos.geometry.force_2d()
    poligonos.geometry = poligonos.geometry.make_valid()

    def encontrar_melhor_coluna(df, is_cliente=False):
        if is_cliente and 'numero_cliente' in df.columns:
            return 'numero_cliente'
        prioridades = ["name", "nome", "id", "description", "descrição"]
        for p in prioridades:
            for col in df.columns:
                if col.lower() == p:
                    dados_validos = df[col].dropna().astype(str).str.strip()
                    dados_validos = dados_validos[~dados_validos.isin(["", "nan", "None", "<Null>"])]
                    if not dados_validos.empty: return col
        for col in df.columns:
            if col.lower() != 'geometry':
                dados_validos = df[col].dropna().astype(str).str.strip()
                dados_validos = dados_validos[~dados_validos.isin(["", "nan", "None", "<Null>"])]
                if not dados_validos.empty: return col
        return df.columns[0]

    col_poli_orig = encontrar_melhor_coluna(poligonos)
    col_cli_orig = encontrar_melhor_coluna(clientes, is_cliente=True)
    
    callback_progresso(70, "Realizando junção espacial (bem mais rápido agora)...")
    resultado = gpd.sjoin(clientes, poligonos, how="left", predicate="intersects")
    col_cli_nome = f"{col_cli_orig}_left" if f"{col_cli_orig}_left" in resultado.columns else col_cli_orig
    col_poli_nome = f"{col_poli_orig}_right" if f"{col_poli_orig}_right" in resultado.columns else col_poli_orig

    def extrair_coordenadas(geom):
        if not geom or geom.is_empty: return ""
        try: return f"{geom.y:.6f}, {geom.x:.6f}"
        except Exception: return f"{geom.centroid.y:.6f}, {geom.centroid.x:.6f}"
            
    resultado["Coord_Lat_Lon"] = resultado.geometry.apply(extrair_coordenadas)
    
    callback_progresso(85, "Preparando dados e resultados extras...")
    
    # ---------------------------------------------------------
    # IDENTIFICAÇÃO E INCLUSÃO DAS NOVAS COLUNAS
    # ---------------------------------------------------------
    colunas_finais = [col_poli_nome, col_cli_nome, "Coord_Lat_Lon"]
    
    # Busca dinamicamente ignorando letras maiúsculas/minúsculas e nomes cortados
    col_estado_cliente = next((c for c in resultado.columns if 'estado_cliente' in c.lower()), None)
    col_estado_forn = next((c for c in resultado.columns if 'fornec' in c.lower()), None) # Pega ESTADO_FORNECIMENTO ou EST_FORNECI...
    
    # Adiciona as colunas na lista de extração, se elas existirem na base
    if col_estado_cliente: colunas_finais.append(col_estado_cliente)
    if col_estado_forn: colunas_finais.append(col_estado_forn)

    # Extrai o resumo com as novas colunas
    resumo = resultado[colunas_finais].copy()
    resumo[col_poli_nome] = resumo[col_poli_nome].fillna("Fora dos Polígonos")
    
    # Dicionário de renomeação para as colunas do Excel ficarem padronizadas e bonitas
    renomeacoes = {
        col_poli_nome: "Poligono", 
        col_cli_nome: "Cliente", 
        "Coord_Lat_Lon": "Coordenadas (Lat, Lon)"
    }
    if col_estado_cliente: renomeacoes[col_estado_cliente] = "Estado do Cliente"
    if col_estado_forn: renomeacoes[col_estado_forn] = "Estado de Fornecimento"
    
    resumo = resumo.rename(columns=renomeacoes)
    resumo = resumo.sort_values(by=["Poligono", "Cliente"])
    # ---------------------------------------------------------

    callback_progresso(95, "Gerando resultados em memória...")
    contagem_interface = resumo.groupby("Poligono").size().reset_index(name="Quantidade")
    
    dados_interface = []
    for _, row in contagem_interface.iterrows():
        dados_interface.append({
            "localidade": str(row["Poligono"]),
            "quantidade": str(row["Quantidade"])
        })
        
    callback_progresso(100, "Análise concluída com sucesso!")
    mensagem = "A contagem foi finalizada! Verifique os resultados na tela."
    
    return mensagem, dados_interface, resumo
