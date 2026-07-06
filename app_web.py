import streamlit as st
import requests
import urllib3
import os
import tempfile
import json
import urllib.parse

# Configuração global da página
st.set_page_config(page_title="Sistema SRC's", page_icon="☁️", layout="wide")
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Configurações do Supabase
SUPABASE_URL = "https://guqwynqjgqtfnqhcxbgn.supabase.co"
SUPABASE_KEY = "sb_publishable_TevT4SZ8bJX76xZLVYyY6Q_7CS_yDgr"

# Importações dos motores Back-end
from app_ccp import executar_automacao_ccp_backend, salvar_e_formatar_excel

# =========================================================
# GERENCIAMENTO DE E-MAILS (Aqui estavam as funções faltando!)
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
# GERENCIAMENTO DE SESSÃO E AUXILIARES
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
# LÓGICA DE LOGIN
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
# RENDERIZAÇÃO DAS TELAS (FRONT-END)
# =========================================================

def renderizar_tela_home():
    st.title("BEM-VINDO!")
    st.markdown(f"### Olá, {st.session_state['usuario']}! O que vamos fazer hoje?")
    st.info("Selecione uma ferramenta no menu lateral para começar.")

def renderizar_tela_ccp():
    st.title("Contagem de Clientes por Polígonos")
    st.markdown("Cruze a base de clientes com os polígonos das áreas.")
    
    col_form, col_resultado = st.columns([3, 2])
    
    with col_form:
        st.markdown("##### 1. Envie os arquivos")
        arquivo_clientes = st.file_uploader("Base de Clientes (KML ou KMZ)", type=["kml", "kmz"])
        arquivo_poligonos = st.file_uploader("Base de Polígonos (KML ou KMZ)", type=["kml", "kmz"])
        
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
        st.markdown("##### 2. Resultados")
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
                label="BAIXAR EXCEL",
                data=bytes_excel,
                file_name="Resumo_Clientes_por_Poligono.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )


def renderizar_tela_script():
    st.title("Unificador de Ordens (Script)")
    st.markdown("Unifique a Ordem Gerada com a Ordem VIC e baixe o relatório consolidado.")
    
    # ========================================
    # PAINEL DE CONFIGURAÇÃO DE E-MAILS
    # ========================================
    with st.expander("Configurar E-mails de Destino", expanded=False):
        config = carregar_config_emails()
        
        col_to, col_cc = st.columns(2)
        with col_to:
            txt_to = st.text_area("Para (TO) - Um por linha:", value="\n".join(config.get("script_to", [])), height=150)
        with col_cc:
            txt_cc = st.text_area("Cópia (CC) - Um por linha:", value="\n".join(config.get("script_cc", [])), height=150)
            
        if st.button("Salvar E-mails", type="secondary"):
            config["script_to"] = [e.strip() for e in txt_to.split('\n') if e.strip()]
            config["script_cc"] = [e.strip() for e in txt_cc.split('\n') if e.strip()]
            salvar_config_emails(config)
            st.success("E-mails atualizados com sucesso!")

    st.divider()

    # ========================================
    # ÁREA DE PROCESSAMENTO
    # ========================================
    col_form, col_resultado = st.columns([3, 2])
    
    with col_form:
        st.markdown("##### 1. Envie os arquivos TXT")
        arq_gerada = st.file_uploader("Ordem Gerada (TXT)", type=["txt"])
        arq_vic = st.file_uploader("Ordem VIC (TXT)", type=["txt"])
        
        if st.button("▶ INICIAR PROCESSAMENTO", type="primary", use_container_width=True):
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
                    
                    caminho_final = executar_automacao_backend(
                        cam_gerada, cam_vic, pasta_saida, atualizar_progresso_web, [], []
                    )
                    
                    st.success("Arquivo processado com sucesso!")
                    st.session_state["script_caminho"] = caminho_final
                    
                except Exception as e:
                    st.error(f"❌ Ocorreu um erro no processamento: {e}")

    # ========================================
    # ÁREA DE RESULTADOS (DOWNLOAD E E-MAIL)
    # ========================================
    with col_resultado:
        st.markdown("##### 2. Resultados")
        
        if "script_caminho" in st.session_state and os.path.exists(st.session_state["script_caminho"]):
            caminho_excel = st.session_state["script_caminho"]
            nome_arquivo = os.path.basename(caminho_excel)
            
            st.info("Sua planilha consolidada está pronta!")
            
            with open(caminho_excel, "rb") as f:
                bytes_excel = f.read()
                
            st.download_button(
                label="1. BAIXAR RELATÓRIO EXCEL",
                data=bytes_excel,
                file_name=nome_arquivo,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
            
            st.markdown("<p style='text-align: center; font-size: 13px; color: #777;'>Após baixar, clique abaixo para preparar o e-mail e arraste o anexo.</p>", unsafe_allow_html=True)
            
            config_atual = carregar_config_emails()
            to_emails = ";".join(config_atual.get("script_to", []))
            cc_emails = ";".join(config_atual.get("script_cc", []))
            
            assunto = urllib.parse.quote("Relatório Diário - Unificador de Ordens")
            corpo = urllib.parse.quote("Prezados,\n\nSegue em anexo o relatório unificado de ordens processado pelo sistema SRC's.\n\nAtenciosamente,")
            
            mailto_link = f"mailto:{to_emails}?cc={cc_emails}&subject={assunto}&body={corpo}"
            
            btn_outlook = f"""
            <a href="{mailto_link}" target="_blank" style="
                text-decoration: none;
                background-color: #2A7B76;
                color: white;
                padding: 10px 20px;
                border-radius: 8px;
                display: inline-block;
                font-weight: bold;
                text-align: center;
                width: 100%;
                border: 1px solid #226460;
            ">2. ABRIR OUTLOOK</a>
            """
            st.markdown(btn_outlook, unsafe_allow_html=True)


def renderizar_tela_relatorio():
    st.title("Relatório Diário")
    st.markdown("Gere o relatório consolidado diário com base nas ordens processadas.")
    
    # ========================================
    # PAINEL DE CONFIGURAÇÃO DE E-MAILS
    # ========================================
    with st.expander("Configurar E-mails de Destino", expanded=False):
        config = carregar_config_emails()
        
        col_to, col_cc = st.columns(2)
        with col_to:
            txt_to = st.text_area("Para (TO) - Um por linha (Relatório):", value="\n".join(config.get("relatorio_to", [])), height=150)
        with col_cc:
            txt_cc = st.text_area("Cópia (CC) - Um por linha (Relatório):", value="\n".join(config.get("relatorio_cc", [])), height=150)
            
        if st.button("Salvar E-mails do Relatório", type="secondary"):
            config["relatorio_to"] = [e.strip() for e in txt_to.split('\n') if e.strip()]
            config["relatorio_cc"] = [e.strip() for e in txt_cc.split('\n') if e.strip()]
            salvar_config_emails(config)
            st.success("E-mails atualizados com sucesso!")

    st.divider()

    # ========================================
    # ÁREA DE PROCESSAMENTO
    # ========================================
    col_form, col_resultado = st.columns([3, 2])
    
    with col_form:
        st.markdown("##### 1. Envie o arquivo Excel")
        arq_relatorio = st.file_uploader("Ordens Consolidadas (XLSX)", type=["xlsx", "xls"])
        
        if st.button("▶ INICIAR PROCESSAMENTO", type="primary", use_container_width=True):
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
                    
                    # Passamos listas vazias ([], []) para o motor pular a abertura forçada do Outlook!
                    caminho_final = executar_automacao_relatorio_backend(
                        cam_relatorio, pasta_saida, atualizar_progresso_web, [], []
                    )
                    
                    st.success("Relatório gerado com sucesso!")
                    st.session_state["relatorio_caminho"] = caminho_final
                    
                except Exception as e:
                    st.error(f"❌ Ocorreu um erro no processamento: {e}")

    # ========================================
    # ÁREA DE RESULTADOS (DOWNLOAD E E-MAIL)
    # ========================================
    with col_resultado:
        st.markdown("##### 2. Resultados")
        
        if "relatorio_caminho" in st.session_state and os.path.exists(st.session_state["relatorio_caminho"]):
            caminho_excel = st.session_state["relatorio_caminho"]
            nome_arquivo = os.path.basename(caminho_excel)
            
            st.info("Seu relatório diário está pronto!")
            
            # --- BOTÃO 1: DOWNLOAD DO EXCEL ---
            with open(caminho_excel, "rb") as f:
                bytes_excel = f.read()
                
            st.download_button(
                label="1. BAIXAR RELATÓRIO EXCEL",
                data=bytes_excel,
                file_name=nome_arquivo,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
            
            # --- BOTÃO 2: ABRIR OUTLOOK (MAILTO) ---
            st.markdown("<p style='text-align: center; font-size: 13px; color: #777;'>Após baixar, clique abaixo para preparar o e-mail e arraste o anexo.</p>", unsafe_allow_html=True)
            
            config_atual = carregar_config_emails()
            to_emails = ";".join(config_atual.get("relatorio_to", []))
            cc_emails = ";".join(config_atual.get("relatorio_cc", []))
            
            assunto = urllib.parse.quote("Relatório Diário de Operações")
            corpo = urllib.parse.quote("Prezados,\n\nSegue em anexo o relatório diário consolidado das operações.\n\nAtenciosamente,")
            
            mailto_link = f"mailto:{to_emails}?cc={cc_emails}&subject={assunto}&body={corpo}"
            
            btn_outlook = f"""
            <a href="{mailto_link}" target="_blank" style="
                text-decoration: none;
                background-color: #2A7B76;
                color: white;
                padding: 10px 20px;
                border-radius: 8px;
                display: inline-block;
                font-weight: bold;
                text-align: center;
                width: 100%;
                border: 1px solid #226460;
            ">2. ABRIR OUTLOOK</a>
            """
            st.markdown(btn_outlook, unsafe_allow_html=True)


def renderizar_tela_acac():
    st.title("Gerador de Ofícios ACAC")
    st.markdown("Faça o upload do polígono da área para gerar o ofício com dados criminais, geográficos e restrições de entrega.")
    
    st.divider()

    col_form, col_resultado = st.columns([3, 2])
    
    with col_form:
        st.markdown("##### 1. Envie o Polígono da Área")
        arquivo_poligono = st.file_uploader("Polígono (KML ou KMZ)", type=["kml", "kmz"])
        
        if st.button("▶ GERAR OFÍCIO WORD", type="primary", use_container_width=True):
            if not arquivo_poligono:
                st.warning("Por favor, faça o upload do arquivo KML/KMZ.")
            else:
                cam_poligono = salvar_arquivo_temporario(arquivo_poligono)
                
                # A pasta de saída é temporária do servidor web
                pasta_saida = tempfile.gettempdir()
                
                barra_progresso = st.progress(0)
                status_texto = st.empty()
                
                def atualizar_progresso_web(valor, texto):
                    barra_progresso.progress(valor)
                    status_texto.text(texto)

                try:
                    from app_breve import executar_automacao_breve_backend
                    
                    # Roda o motor pesado conectando com o SharePoint e desenhando os mapas
                    caminho_final = executar_automacao_breve_backend(
                        cam_poligono, pasta_saida, atualizar_progresso_web
                    )
                    
                    st.success("Ofício gerado com sucesso!")
                    
                    # Salva o caminho do arquivo Word na memória da página
                    st.session_state["acac_caminho"] = caminho_final
                    
                except Exception as e:
                    st.error(f"❌ Ocorreu um erro no processamento: {e}")

    with col_resultado:
        st.markdown("##### 2. Resultados")
        
        # Verifica se o arquivo foi processado e existe na memória
        if "acac_caminho" in st.session_state and os.path.exists(st.session_state["acac_caminho"]):
            caminho_docx = st.session_state["acac_caminho"]
            nome_arquivo = os.path.basename(caminho_docx)
            
            st.info("Seu Ofício ACAC está pronto!")
            
            # Lê o arquivo Word gerado para liberar o download no navegador
            with open(caminho_docx, "rb") as f:
                bytes_docx = f.read()
                
            st.download_button(
                label="BAIXAR OFÍCIO WORD",
                data=bytes_docx,
                file_name=nome_arquivo,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                use_container_width=True
            )
# =========================================================
# ESTRUTURA PRINCIPAL DA PÁGINA
# =========================================================
if not st.session_state["autenticado"]:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("<h1 style='text-align: center; color: #2A7B76;'>Sistema SRC's</h1>", unsafe_allow_html=True)
        st.markdown("<h5 style='text-align: center; color: #7A8B8B;'>Faça login para continuar</h5>", unsafe_allow_html=True)
        st.write("")
        
        with st.form("form_login"):
            br0_input = st.text_input("ID (BRO):")
            senha_input = st.text_input("SENHA:", type="password")
            submit = st.form_submit_button("ENTRAR", use_container_width=True)
            if submit: validar_login(br0_input, senha_input)

else:
    with st.sidebar:
        st.markdown(f"### 👤 {st.session_state['usuario']}")
        st.markdown("Administrador")
        st.divider()
        
        menu_selecionado = st.radio(
            "NAVEGAÇÃO", 
            ["Home", "C.C.P.", "Script", "Relatório", "Gerador ACAC"]
        )
        
        st.divider()
        if st.button("Sair do Sistema", use_container_width=True):
            st.session_state.clear()
            st.rerun()

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