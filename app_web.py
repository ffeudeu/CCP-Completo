import streamlit as st
import os
import tempfile
import urllib3
import pandas as pd

# Configuração global da página (Sem barra lateral)
st.set_page_config(page_title="C.C.P - Contagem Rápida", page_icon="", layout="wide", initial_sidebar_state="collapsed")
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Importação do motor Back-end
from app_ccp import executar_automacao_ccp_backend, salvar_e_formatar_excel

# =========================================================
# ESTILO VISUAL CORPORATIVO
# =========================================================
ESTILO_CUSTOMIZADO = '''
<style>
    [data-testid="collapsedControl"] { display: none !important; }
    .stApp { background-color: #F4F7F6 !important; }
    .block-container .element-container div.stAlert, 
    .block-container .stMarkdown, 
    div.stContainer {
        background-color: #FFFFFF !important;
        border: 1px solid #EAEAEA !important;
        border-radius: 12px !important;
    }
    button[kind="primary"] {
        background-color: #2A7B76 !important;
        color: #FFFFFF !important;
        border-radius: 8px !important;
        font-weight: bold !important;
        border: none !important;
        padding: 12px !important;
    }
    button[kind="primary"]:hover {
        background-color: #226460 !important;
        color: #FFFFFF !important;
    }
    .titulo-decorativo::after {
        content: "";
        display: block;
        width: 40px;
        height: 3px;
        background-color: #2A7B76;
        border-radius: 1px;
        margin-top: 8px;
        margin-bottom: 20px;
    }
</style>
'''
st.markdown(ESTILO_CUSTOMIZADO, unsafe_allow_html=True)

# =========================================================
# FUNÇÃO AUXILIAR
# =========================================================
def salvar_arquivo_temporario(uploaded_file):
    if uploaded_file is not None:
        caminho_temp = os.path.join(tempfile.gettempdir(), uploaded_file.name)
        with open(caminho_temp, "wb") as f:
            f.write(uploaded_file.getbuffer())
        return caminho_temp
    return None

# =========================================================
# TELA ÚNICA DE PROCESSAMENTO
# =========================================================
CAMINHO_PASTA_SUCURSAIS = r"C:\Users\BR0177332417\OneDrive - Enel Spa\Documents\Nova pasta\OneDrive - Enel Spa\tabelas_sucursais"

st.write("") 
st.markdown("<h1 class='titulo-decorativo'>CONTAGEM DE CLIENTES POR POLÍGONOS (C.C.P)</h1>", unsafe_allow_html=True)
st.markdown("Cruze múltiplas bases de clientes com vários polígonos das áreas operacionais de forma super rápida.")
st.write("")

col_form, col_resultado = st.columns([3, 2])

# LADO ESQUERDO: UPLOAD DOS ARQUIVOS E EXECUÇÃO
with col_form:
    with st.container(border=True):
        st.markdown("##### Configuração de Análise")
        st.info("Escolha uma ou mais sucursais.")
        
        pasta_sucursais = st.text_input(" Caminho da Pasta de Sucursais:", value=CAMINHO_PASTA_SUCURSAIS)
        
        arquivos_disponiveis = []
        if os.path.exists(pasta_sucursais):
            arquivos_disponiveis = [f for f in os.listdir(pasta_sucursais) if f.lower().endswith(('.csv', '.xlsx', '.xls'))]
            
        if not arquivos_disponiveis:
            st.warning("⚠️ Nenhum arquivo encontrado nesta pasta. Verifique se o caminho acima está correto.")
            sucursais_selecionadas = []
        else:
            opcoes_formatadas = {}
            for arq in arquivos_disponiveis:
                nome_bonito = arq.replace("tabela_", "").replace(".csv", "").replace(".xlsx", "").replace("_", " ")
                opcoes_formatadas[nome_bonito] = arq
                
            # MUDANÇA: multiselect permite escolher 1 ou todas as sucursais da pasta
            escolhas_bonitas = st.multiselect(" Selecione a(s) Sucursal(is):", options=sorted(list(opcoes_formatadas.keys())))
            
            # Recupera o nome real do arquivo de cada escolha feita
            sucursais_selecionadas = [opcoes_formatadas[escolha] for escolha in escolhas_bonitas]
        
        # MUDANÇA: accept_multiple_files=True permite enviar vários KMLs/KMZs ao mesmo tempo
        arquivos_poligonos = st.file_uploader(" Selecione o(s) arquivo(s) de Polígonos (KML/KMZ)", type=["kml", "kmz"], accept_multiple_files=True)
        
        if st.button(" INICIAR PROCESSAMENTO", type="primary", use_container_width=True):
            if not arquivos_poligonos:
                st.warning("Por favor, faça o upload de pelo menos um arquivo de Polígonos antes de iniciar.")
            elif not sucursais_selecionadas:
                st.error("Selecione pelo menos uma base de sucursal antes de continuar.")
            else:
                # Salva todos os KMLs/KMZs numa lista de caminhos temporários
                caminhos_pol = [salvar_arquivo_temporario(arq) for arq in arquivos_poligonos]
                
                # Monta os caminhos completos das bases escolhidas
                caminhos_bases = [os.path.join(pasta_sucursais, suc) for suc in sucursais_selecionadas]
                
                barra_progresso = st.progress(0)
                status_texto = st.empty()
                
                def atualizar_progresso_web(valor, texto):
                    barra_progresso.progress(valor)
                    status_texto.text(texto)

                try:
                    mensagem, dados_tabela, df_resumo_completo = executar_automacao_ccp_backend(caminhos_pol, caminhos_bases, atualizar_progresso_web)
                    st.success(f"✔ Análise processada com sucesso!")
                    st.session_state["ccp_dados"] = dados_tabela
                    st.session_state["ccp_df_completo"] = df_resumo_completo
                except Exception as e:
                    st.error(f"❌ Ocorreu um erro no processamento: {e}")

# LADO DIREITO: RESULTADO E DOWNLOADS
with col_resultado:
    with st.container(border=True):
        st.markdown("##### Resultados da Análise")
        
        if "ccp_dados" in st.session_state:
            df_exibicao = pd.DataFrame(st.session_state["ccp_dados"])
            df_exibicao.columns = ["Localidade", "Quantidade de Clientes"]
            
            st.dataframe(df_exibicao, use_container_width=True, hide_index=True)
            
            df_completo = st.session_state["ccp_df_completo"]
            df_excel_filtrado = df_completo[~df_completo['Poligono'].str.contains('Fora', case=False, na=False)].copy()
            pode_gerar_excel = not df_excel_filtrado.empty and len(df_excel_filtrado) <= 1048500
            
            st.write("---")
            st.markdown("###### Opções de Exportação:")
            
            col_btn1, col_btn2 = st.columns(2)
            
            with col_btn1:
                if pode_gerar_excel:
                    caminho_excel_temp = os.path.join(tempfile.gettempdir(), "Clientes_Filtrados_Poligono.xlsx")
                    salvar_e_formatar_excel(df_excel_filtrado, caminho_excel_temp, "Resumo")
                    with open(caminho_excel_temp, "rb") as f: bytes_excel = f.read()
                        
                    st.download_button(
                        label="EXCEL (Apenas Dentro)",
                        data=bytes_excel,
                        file_name="Clientes_Dentro_Poligono.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True
                    )
                else:
                    st.info("⚠️ Nenhum cliente encontrado dentro ou limite excedido.")
                        
            with col_btn2:
                csv_export = df_completo.to_csv(index=False, sep=";", encoding="utf-8-sig")
                st.download_button(
                    label=" CSV (Base Completa)",
                    data=csv_export,
                    file_name="Resumo_Completo_Clientes.csv",
                    mime="text/csv",
                    type="secondary",
                    use_container_width=True
                )
        else:
            st.markdown("<p style='color: #7A8B8B; font-size: 14px; text-align: center; margin-top: 50px;'>Aguardando o processamento...</p>", unsafe_allow_html=True)
