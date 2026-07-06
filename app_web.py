import streamlit as st
import requests
import urllib3
import os
import tempfile
import json
import urllib.parse
from datetime import datetime

# Configuração global da página
st.set_page_config(page_title="Sistema SRC's", page_icon="☁️", layout="wide")
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Configurações do Supabase
SUPABASE_URL = "https://guqwynqjgqtfnqhcxbgn.supabase.co"
SUPABASE_KEY = "sb_publishable_TevT4SZ8bJX76xZLVYyY6Q_7CS_yDgr"

# Importações dos motores Back-end
from app_ccp import executar_automacao_ccp_backend, salvar_e_formatar_excel

# =========================================================
# INJEÇÃO DE ESTILO
# =========================================================
ESTILO_EXE_CUSTOMIZADO = """
<style>
    /* Altera o fundo geral do aplicativo para o cinza claro corporativo */
    .stApp {
        background-color: #F4F7F6 !important;
    }
    
    /* Customização do Menu Lateral (Sidebar) para ficar idêntico ao do .exe */
    [data-testid="stSidebar"] {
        background-color: #FFFFFF !important;
        border-right: 1px solid #EAEAEA !important;
    }
    
    /* Ajustes nos textos e títulos da Sidebar */
    [data-testid="stSidebar"] .stMarkdown h3 {
        color: #2A7B76 !important;
        font-weight: bold !important;
    }
    
    /* Customização dos botões de rádio (Navegação) para imitar o menu PyQt6 */
    div[data-testid="stRadio"] > label {
        display: none !important; /* Oculta o título "NAVEGAÇÃO" para limpar o visual */
    }
    
    /* Deixa os itens de menu com espaçamento e cores do .exe */
    div[data-testid="stRadio"] div[role="radiogroup"] label {
        background-color: transparent !important;
        border-radius: 6px !important;
        padding: 10px 15px !important;
        color: #7A8B8B !important;
        font-weight: bold !important;
        font-size: 14px !important;
        transition: all 0.2s ease;
        margin-bottom: 5px !important;
        border-left: 4px solid transparent !important;
    }
    
    /* Efeito de Hover (Passar o mouse) no menu */
    div[data-testid="stRadio"] div[role="radiogroup"] label:hover {
        background-color: #F4F7F6 !important;
        color: #2A7B76 !important;
    }
    
    /* Elemento Ativo/Selecionado no Menu (A mágica do estilo checked do PyQt6) */
    div[data-testid="stRadio"] div[role="radiogroup"] label[data-checked="true"] {
        background-color: #E8F1F1 !important;
        color: #2A7B76 !important;
        border-left: 4px solid #2A7B76 !important;
    }
    
    /* Oculta os círculos de seleção padrão do rádio para parecer botões reais */
    div[data-testid="stRadio"] div[role="radiogroup"] label div[data-testid="stMarkdownContainer"] p::before {
        display: none !important;
    }
    
    /* Customização dos Cartões de Conteúdo (Cards Brancos) */
    .block-container .element-container div.stAlert, 
    .block-container .stMarkdown, 
    div[data-testid="stExpander"] {
        background-color: #FFFFFF !important;
        border: 1px solid #EAEAEA !important;
        border-radius: 12px !important;
    }
    
    /* Botões Principais de Ação (A cor Verde-Azulada da Enel) */
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
    
    /* Botões Secundários (Procurar / Configurar) */
    button[kind="secondary"] {
        background-color: #FDFDFD !important;
        color: #2A7B76 !important;
        border: 1px solid #DCDCDC !important;
        border-radius: 8px !important;
        font-weight: bold !important;
    }
    button[kind="secondary"]:hover {
        background-color: #F4F7F6 !important;
        color: #2A7B76 !important;
    }
    
    /* Customização das Caixas de Input (Campos de texto e formulário) */
    div[data-testid="stTextInput"] input, div[data-testid="stTextArea"] textarea {
        background-color: #FFFFFF !important;
        border-radius: 8px !important;
        border: 1px solid #DCDCDC !important;
        color: #333333 !important;
    }
    div[data-testid="stTextInput"] input:focus {
        border-color: #2A7B76 !important;
    }
    
    /* Linha decorativa abaixo dos títulos (Igual ao detalhe de 3px do .exe) */
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
"""
st.markdown(ESTILO_EXE_CUSTOMIZADO, unsafe_allow_html=True)

# =========================================================
# GERENCIAMENTO DE CONFIGURAÇÃO DE E-MAILS
# =========================================================
ARQUIVO_CONFIG_EMAILS = "config_emails.json"
EMAILS_PADRAO = {
    "script_to": ["raquel.fernandes@enel.com", "raquel.barros@applus.com", "carla.silva@applus.com", "danilo.mendes@applus.com", "loana.silva@applus.com", "gilmar.santiago@applus.com", "lroccacr@emeal.nttdata.com"],
    "script_cc": ["leandro.vasconcellos@enel.com", "fabio.casanova@enel.com", "raquel.fernandes@enel.com"],
    "relatorio_to": ["raquel.fernandes@enel.com", "anderson.santana@enel.com", "antonio.souza@enel.com", "leandro.vasconcellos@enel.com", "wellington.nascimento@enel.com", "glauco.chaves@enel.com", "marcelo.araujo@enel.com", "luizhenrique.araujo@enel.com", "edesio.teixeira@enel.com", "alberto.sartori@enel.com", "fabio.damasceno@enel.com", "felipe.pinto@enel.com", "ajose.barbosa@enel.com", "brender.rangel@enel.com", "andre.caldas@enel.com", "laura.mothe@enel.com", "douglas.leal@enel.com", "carolina.teixeira@enel.com", "fernando.brito@enel.com", "fabio.casanova@enel.com", "lroccacr@emeal.nttdata.com"],
    "relatorio_cc": ["romolo.moreira@enel.com", "erika.luca@enel.com", "silvan.souza@enel.com", "fernandofalcao.silva@enel.com", "wilson.figueiredo@enel.com", "eduardomota.santos@enel.com", "carla.silva@applus.com", "daniel.azevedo@enel.com", "thales.rezende@enel.com", "thais.mendonca@enel.com", "rai.amorim@enel.com", "renan.martinez@enel.com", "marcio.rsantos@enel.com", "frederico.peixoto@enel.com", "guilherme.frauches@enel.com", "rogerio.soares@enel.com", "everton.paiva@enel.com", "bruno.garcia@enel.com", "leandro.pinto@enel.com", "carlos.manhaes@enel.com", "rubens.ferraz@enel.com", "raquel.barros@applus.com", "josecicero.silva@enel.com", "danilo.mendes@applus.com", "gabrielle.silva@enel.com", "douglas.asilva@enel.com", "guilherme.frauches@enel.com", "alessander.costaoliveira@nttdata.com", "igor.gama@enel.com", "fellipe.aubynferrando@nttdata.com", "estela.pereira@enel.com", "ericson.dasilva@enel.com", "thiago.fcurvao@enel.com", "celine.nascimentoferreira@nttdata.com", "adriana.mottamarques@nttdata.com", "thainar.silveriosilva@nttdata.com", "robson.pereira@enel.com"]
}

def carregar_config_emails():
    if os.path.exists(ARQUIVO_CONFIG_EMAILS):
        try:
            with open(ARQUIVO_CONFIG_EMAILS, "r", encoding="utf-8") as f:
                return json.load(f)
        except: pass
    return EMAILS_PADRAO.copy()

def salvar_config_emails(config):
    with open(ARQUIVO_CONFIG_EMAILS, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=4)

# =========================================================
# AUXILIARES DE MEMÓRIA
# =========================================================
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False
    st.session_state["usuario"] = ""

def salvar_arquivo_temporario(uploaded_file):
    if uploaded_file is not None:
        caminho_temp = os.path.join(tempfile.gettempdir(), uploaded_file.name)
        with open(caminho_temp, "wb") as f:
            f.write(uploaded_file.getbuffer())
        return caminho_temp
    return None

# =========================================================
# VALIDAÇÃO DO LOGIN (SUPABASE)
# =========================================================
def validar_login(br0, senha):
    if not br0 or not senha:
        st.warning("Por favor, preencha o ID e a Senha.")
        return
    try:
        endpoint = f"{SUPABASE_URL}/rest/v1/usuarios"
        headers = {"apikey": SUPABASE_KEY, "Authorization": f"Bearer {SUPABASE_KEY}"}
        params = {"br0": f"eq.{br0}", "senha": f"eq.{senha}"}
        
        resposta = requests.get(endpoint, headers=headers, params=params, verify=False)
        if resposta.status_code == 200 and len(resposta.json()) > 0:
            st.session_state["autenticado"] = True
            st.session_state["usuario"] = br0
            st.rerun()
        else:
            st.error("❌ Senha ou BR0 incorreto.")
    except Exception as e:
        st.error(f"❌ Erro de Conexão: {e}")

# =========================================================
# FUNÇÕES DE RENDERIZAÇÃO DAS TELAS
# =========================================================

def renderizar_tela_home():
    st.markdown("<h1 class='titulo-decorativo'>BEM-VINDO!</h1>", unsafe_allow_html=True)
    st.markdown(f"### Olá, {st.session_state['usuario']}! O que vamos fazer hoje?")
    st.info(" Selecione uma ferramenta no menu lateral esquerdo para iniciar os trabalhos.")

def renderizar_tela_ccp():
    st.markdown("<h1 class='titulo-decorativo'>CONTAGEM DE CLIENTES POR POLÍGONOS</h1>", unsafe_allow_html=True)
    st.markdown("Cruze a base de clientes georreferenciados com os polígonos das áreas operacionais.")
    
    col_form, col_resultado = st.columns([3, 2])
    
    with col_form:
        st.markdown("##### Envie os arquivos Geográficos")
        arquivo_clientes = st.file_uploader("Selecione o arquivo de Clientes (KML ou KMZ)", type=["kml", "kmz"])
        arquivo_poligonos = st.file_uploader("Selecione o arquivo de Polígonos (KML ou KMZ)", type=["kml", "kmz"])
        
        if st.button("▶ INICIAR PROCESSAMENTO", type="primary", use_container_width=True):
            if not arquivo_clientes or not arquivo_poligonos:
                st.warning("Por favor, faça o upload dos dois arquivos antes de iniciar.")
            else:
                cam_cli = salvar_arquivo_temporario(arquivo_clientes)
                cam_pol = salvar_arquivo_temporario(arquivo_poligonos)
                
                barra_progresso = st.progress(0)
                status_texto = st.empty()
                
                def atualizar_progresso_web(valor, texto):
                    barra_progresso.progress(valor)
                    status_texto.text(texto)

                try:
                    mensagem, dados_tabela, df_resumo = executar_automacao_ccp_backend(cam_pol, cam_cli, atualizar_progresso_web)
                    st.success("Análise concluída com sucesso!")
                    st.session_state["ccp_dados"] = dados_tabela
                    st.session_state["ccp_df"] = df_resumo
                except Exception as e:
                    st.error(f"❌ Ocorreu um erro no processamento: {e}")

    with col_resultado:
        st.markdown("##### Resultados da Análise")
        if "ccp_dados" in st.session_state:
            import pandas as pd
            df_exibicao = pd.DataFrame(st.session_state["ccp_dados"])
            df_exibicao.columns = ["Localidade", "Quantidade de Clientes"]
            st.dataframe(df_exibicao, use_container_width=True, hide_index=True)
            
            caminho_excel_temp = os.path.join(tempfile.gettempdir(), "Resumo_Clientes_por_Poligono.xlsx")
            salvar_e_formatar_excel(st.session_state["ccp_df"], caminho_excel_temp, "Resumo")
            
            with open(caminho_excel_temp, "rb") as f:
                bytes_excel = f.read()
                
            st.download_button(
                label="EXPORTAR RELATÓRIO EXCEL",
                data=bytes_excel,
                file_name="Resumo_Clientes_por_Poligono.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )

def renderizar_tela_script():
    st.markdown("<h1 class='titulo-decorativo'>UNIFICADOR DE ORDENS</h1>", unsafe_allow_html=True)
    st.markdown("Unifique o arquivo da Ordem Gerada com a Ordem VIC e prepare a planilha final.")
    
    with st.expander("Configurar E-mails de Destino (Para e CC)", expanded=False):
        config = carregar_config_emails()
        col_to, col_cc = st.columns(2)
        with col_to:
            txt_to = st.text_area("Para (TO) - Um por linha:", value="\n".join(config.get("script_to", [])), height=120)
        with col_cc:
            txt_cc = st.text_area("Cópia (CC) - Um por linha:", value="\n".join(config.get("script_cc", [])), height=120)
            
        if st.button("Salvar E-mails Configurados", type="secondary"):
            config["script_to"] = [e.strip() for e in txt_to.split('\n') if e.strip()]
            config["script_cc"] = [e.strip() for e in txt_cc.split('\n') if e.strip()]
            salvar_config_emails(config)
            st.success("E-mails atualizados no sistema!")

    st.write("")
    col_form, col_resultado = st.columns([3, 2])
    
    with col_form:
        st.markdown("##### Selecione os Arquivos TXT")
        arq_gerada = st.file_uploader("Selecione o arquivo da Ordem Gerada (.txt)", type=["txt"])
        arq_vic = st.file_uploader("Selecione o arquivo da Ordem VIC (.txt)", type=["txt"])
        
        if st.button("▶ INICIAR PROCESSAMENTO SCRIPT", type="primary", use_container_width=True):
            if not arq_gerada or not arq_vic:
                st.warning("Por favor, faça o upload dos dois arquivos TXT.")
            else:
                cam_gerada = salvar_arquivo_temporario(arq_gerada)
                cam_vic = salvar_arquivo_temporario(arq_vic)
                pasta_saida = tempfile.gettempdir()
                
                barra_progresso = st.progress(0)
                status_texto = st.empty()
                
                def atualizar_progresso_web(valor, texto):
                    barra_progresso.progress(valor)
                    status_texto.text(texto)

                try:
                    from app_script import executar_automacao_backend
                    caminho_final = executar_automacao_backend(cam_gerada, cam_vic, pasta_saida, atualizar_progresso_web, [], [])
                    st.success("Script executado com sucesso!")
                    st.session_state["script_caminho"] = caminho_final
                except Exception as e:
                    st.error(f"❌ Erro no processamento: {e}")

    with col_resultado:
        st.markdown("##### Entrega do Relatório")
        if "script_caminho" in st.session_state and os.path.exists(st.session_state["script_caminho"]):
            caminho_excel = st.session_state["script_caminho"]
            nome_arquivo = os.path.basename(caminho_excel)
            
            st.info("A planilha unificada está pronta para download.")
            with open(caminho_excel, "rb") as f:
                bytes_excel = f.read()
                
            st.download_button(
                label="1. BAIXAR PLANILHA CONSOLIDADA",
                data=bytes_excel,
                file_name=nome_arquivo,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
            
            config_atual = carregar_config_emails()
            to_emails = ";".join(config_atual.get("script_to", []))
            cc_emails = ";".join(config_atual.get("script_cc", []))
            assunto = urllib.parse.quote("Relatório Diário - Unificador de Ordens")
            corpo = urllib.parse.quote("Prezados,\n\nSegue em anexo o relatório unificado de ordens processado pelo sistema SRC's.\n\nAtenciosamente,")
            
            mailto_link = f"mailto:{to_emails}?cc={cc_emails}&subject={assunto}&body={corpo}"
            btn_outlook = f'<a href="{mailto_link}" target="_blank" style="text-decoration: none; background-color: #2A7B76; color: white; padding: 12px 20px; border-radius: 8px; display: inline-block; font-weight: bold; text-align: center; width: 100%;">📧 2. PREPARAR E-MAIL OUTLOOK</a>'
            st.markdown(btn_outlook, unsafe_allow_html=True)

def renderizar_tela_relatorio():
    st.markdown("<h1 class='titulo-decorativo'>RELATÓRIO DIÁRIO</h1>", unsafe_allow_html=True)
    st.markdown("Gere a consolidação diária das ordens para envio da gestão executiva.")
    
    with st.expander("Configurar E-mails do Relatório (Para e CC)", expanded=False):
        config = carregar_config_emails()
        col_to, col_cc = st.columns(2)
        with col_to:
            txt_to = st.text_area("Para (TO) - Um por linha (Relatório):", value="\n".join(config.get("relatorio_to", [])), height=120)
        with col_cc:
            txt_cc = st.text_area("Cópia (CC) - Um por linha (Relatório):", value="\n".join(config.get("relatorio_cc", [])), height=120)
            
        if st.button("Salvar E-mails do Relatório", type="secondary"):
            config["relatorio_to"] = [e.strip() for e in txt_to.split('\n') if e.strip()]
            config["relatorio_cc"] = [e.strip() for e in txt_cc.split('\n') if e.strip()]
            salvar_config_emails(config)
            st.success("E-mails do relatório atualizados!")

    st.write("")
    col_form, col_resultado = st.columns([3, 2])
    
    with col_form:
        st.markdown("##### Base do Relatório")
        arq_relatorio = st.file_uploader("Selecione a planilha de Ordens Consolidadas (.xlsx)", type=["xlsx", "xls"])
        
        if st.button("▶ INICIAR CONSOLIDAÇÃO DIÁRIA", type="primary", use_container_width=True):
            if not arq_relatorio:
                st.warning("Por favor, faça o upload do arquivo consolidado.")
            else:
                cam_relatorio = salvar_arquivo_temporario(arq_relatorio)
                pasta_saida = tempfile.gettempdir()
                
                barra_progresso = st.progress(0)
                status_texto = st.empty()
                
                def atualizar_progresso_web(valor, texto):
                    barra_progresso.progress(valor)
                    status_texto.text(texto)

                try:
                    from app_relatorio import executar_automacao_relatorio_backend
                    caminho_final = executar_automacao_relatorio_backend(cam_relatorio, pasta_saida, atualizar_progresso_web, [], [])
                    st.success("Relatório diário compilado!")
                    st.session_state["relatorio_caminho"] = caminho_final
                except Exception as e:
                    st.error(f"❌ Falha no processamento: {e}")

    with col_resultado:
        st.markdown("##### Entrega do Relatório")
        if "relatorio_caminho" in st.session_state and os.path.exists(st.session_state["relatorio_caminho"]):
            caminho_excel = st.session_state["relatorio_caminho"]
            nome_arquivo = os.path.basename(caminho_excel)
            
            st.info("O arquivo do relatório diário foi montado.")
            with open(caminho_excel, "rb") as f:
                bytes_excel = f.read()
                
            st.download_button(
                label="1. BAIXAR RELATÓRIO OPERAÇÕES",
                data=bytes_excel,
                file_name=nome_arquivo,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
            
            config_atual = carregar_config_emails()
            to_emails = ";".join(config_atual.get("relatorio_to", []))
            cc_emails = ";".join(config_atual.get("relatorio_cc", []))
            assunto = urllib.parse.quote("Relatório Diário de Operações")
            corpo = urllib.parse.quote("Prezados,\n\nSegue em anexo o relatório diário consolidado das operações.\n\nAtenciosamente,")
            
            mailto_link = f"mailto:{to_emails}?cc={cc_emails}&subject={assunto}&body={corpo}"
            btn_outlook = f'<a href="{mailto_link}" target="_blank" style="text-decoration: none; background-color: #2A7B76; color: white; padding: 12px 20px; border-radius: 8px; display: inline-block; font-weight: bold; text-align: center; width: 100%;">📧 2. PREPARAR E-MAIL OUTLOOK</a>'
            st.markdown(btn_outlook, unsafe_allow_html=True)

def renderizar_tela_acac():
    st.markdown("<h1 class='titulo-decorativo'>GERAÇÃO DE OFÍCIO ACAC</h1>", unsafe_allow_html=True)
    st.markdown("Crie documentos oficiais de contingência mapeando restrições logísticas e de segurança.")
    
    col_form, col_resultado = st.columns([3, 2])
    
    with col_form:
        st.markdown("##### Arquivo da Área")
        arquivo_poligono = st.file_uploader("Selecione o arquivo KML/KMZ do Polígono Alvo", type=["kml", "kmz"])
        
        if st.button("▶ GERAR OFÍCIO WORD (ACAC)", type="primary", use_container_width=True):
            if not arquivo_poligono:
                st.warning("Por favor, insira o polígono antes de prosseguir.")
            else:
                cam_poligono = salvar_arquivo_temporario(arquivo_poligono)
                pasta_saida = tempfile.gettempdir()
                
                barra_progresso = st.progress(0)
                status_texto = st.empty()
                
                def atualizar_progresso_web(valor, texto):
                    barra_progresso.progress(valor)
                    status_texto.text(texto)

                try:
                    from app_breve import executar_automacao_breve_backend
                    caminho_final = executar_automacao_breve_backend(cam_poligono, pasta_saida, atualizar_progresso_web)
                    st.success("Ofício ACAC gerado!")
                    st.session_state["acac_caminho"] = caminho_final
                except Exception as e:
                    st.error(f"❌ Erro na compilação do Word: {e}")

    with col_resultado:
        st.markdown("##### Arquivo Gerado")
        if "acac_caminho" in st.session_state and os.path.exists(st.session_state["acac_caminho"]):
            caminho_docx = st.session_state["acac_caminho"]
            nome_arquivo = os.path.basename(caminho_docx)
            
            st.info("O Ofício Word foi estruturado com os dados criminais e o mapa.")
            with open(caminho_docx, "rb") as f:
                bytes_docx = f.read()
                
            st.download_button(
                label="BAIXAR OFÍCIO WORD (DOCX)",
                data=bytes_docx,
                file_name=nome_arquivo,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                use_container_width=True
            )

# =========================================================
# CONTROLADOR DE FLUXO (LOGIN VS PAINEL)
# =========================================================
if not st.session_state["autenticado"]:
    # TELA DE LOGIN ESTILIZADA
    st.write("")
    st.write("")
    col1, col2, col3 = st.columns([1, 1.8, 1])
    with col2:
        st.markdown("<h1 style='text-align: center; color: #2A7B76; font-weight: bold;'>☁️ Sistema SRC's</h1>", unsafe_allow_html=True)
        st.markdown("<h5 style='text-align: center; color: #7A8B8B;'>Ambiente de Nuvem Corporativa</h5>", unsafe_allow_html=True)
        st.write("")
        
        with st.form("form_login"):
            br0_input = st.text_input("ID do Usuário (BR0):", placeholder="Ex: BR012345")
            senha_input = st.text_input("Senha Corporativa:", type="password", placeholder="Digite sua senha")
            submit = st.form_submit_button("CONECTAR AO SISTEMA", use_container_width=True)
            if submit: validar_login(br0_input, senha_input)

else:
    # SISTEMA OPERACIONAL LOGADO (SIDEBAR + CONTEÚDO)
    with st.sidebar:
        st.markdown(f"### 👤 {st.session_state['usuario']}")
        st.markdown("<p style='color: #7A8B8B; margin-top:-15px; font-size:13px;'>Perfil: Administrador</p>", unsafe_allow_html=True)
        st.divider()
        
        # Menu de rádio com os nomes limpos e elegantes (sem os emojis poluindo)
        menu_selecionado = st.radio(
            "MÓDULOS:", 
            ["Home", "C.C.P.", "Script", "Relatório", "Gerador ACAC"]
        )
        
        st.divider()
        if st.button("Desconectar", use_container_width=True):
            st.session_state.clear()
            st.rerun()

    # Roteamento nativo das funções baseado na seleção do usuário
    if menu_selecionado == "Home":
        renderizar_tela_home()
    elif menu_selecionado == "C.C.P.":
        renderizar_tela_ccp()
    elif menu_selecionado == "Script":
        renderizar_tela_script()
    elif menu_selecionado == "Relatório":
        renderizar_tela_relatorio()
    elif menu_selecionado == "Gerador ACAC":
        renderizar_tela_acac()
