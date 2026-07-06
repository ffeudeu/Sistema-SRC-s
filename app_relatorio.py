import os
import time
import pandas as pd
from datetime import datetime, timedelta
import win32com.client as win32
from PIL import ImageGrab
import traceback
import calendar
import unicodedata
import tempfile
import requests
import urllib3
import pythoncom

# ==========================================
# CONFIGURAÇÕES E CONSTANTES
# ==========================================
FERIADOS_2026 = [
    datetime(2026, 1, 1).date(), datetime(2026, 2, 16).date(), datetime(2026, 2, 17).date(),
    datetime(2026, 4, 3).date(), datetime(2026, 4, 21).date(), datetime(2026, 5, 1).date(),
    datetime(2026, 6, 4).date(), datetime(2026, 9, 7).date(), datetime(2026, 10, 12).date(),
    datetime(2026, 11, 2).date(), datetime(2026, 11, 15).date(), datetime(2026, 12, 25).date(),
]

def get_last_business_day(date_obj):
    return date_obj - timedelta(days=1)

DATA_ATUAL_REAL = datetime.now()
DATA_REFERENCIA_D1 = get_last_business_day(DATA_ATUAL_REAL)
DATA_REFERENCIA_STR = DATA_REFERENCIA_D1.strftime('%d/%m/%Y')

meses_ano = {1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"}
NOME_MES_REFERENCIA = meses_ano[DATA_REFERENCIA_D1.month]

import tempfile
from supabase import create_client, Client

# Conexão com o Supabase
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
SUPABASE_URL = "https://guqwynqjgqtfnqhcxbgn.supabase.co"
SUPABASE_KEY = "sb_publishable_TevT4SZ8bJX76xZLVYyY6Q_7CS_yDgr"
supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
'''
EMAIL_DESTINO = [
    "raquel.fernandes@enel.com", "anderson.santana@enel.com", "antonio.souza@enel.com", 
    "leandro.vasconcellos@enel.com", "wellington.nascimento@enel.com", "glauco.chaves@enel.com", 
    "marcelo.araujo@enel.com", "luizhenrique.araujo@enel.com", "edesio.teixeira@enel.com", 
    "alberto.sartori@enel.com", "fabio.damasceno@enel.com", "felipe.pinto@enel.com", 
    "ajose.barbosa@enel.com", "brender.rangel@enel.com", "andre.caldas@enel.com", 
    "laura.mothe@enel.com", "douglas.leal@enel.com", "carolina.teixeira@enel.com", 
    "fernando.brito@enel.com", "fabio.casanova@enel.com", "lroccacr@emeal.nttdata.com"
]
EMAIL_CC = [
    "romolo.moreira@enel.com", "erika.luca@enel.com", "silvan.souza@enel.com", "fernandofalcao.silva@enel.com",
    "wilson.figueiredo@enel.com", "eduardomota.santos@enel.com", "carla.silva@applus.com", "daniel.azevedo@enel.com",
    "thales.rezende@enel.com", "thais.mendonca@enel.com", "rai.amorim@enel.com", "renan.martinez@enel.com",
    "marcio.rsantos@enel.com", "frederico.peixoto@enel.com", "guilherme.frauches@enel.com", "rogerio.soares@enel.com",
    "everton.paiva@enel.com", "bruno.garcia@enel.com", "leandro.pinto@enel.com", "carlos.manhaes@enel.com",
    "rubens.ferraz@enel.com", "raquel.barros@applus.com", "josecicero.silva@enel.com", "danilo.mendes@applus.com",
    "gabrielle.silva@enel.com", "douglas.asilva@enel.com", "guilherme.frauches@enel.com", "alessander.costaoliveira@nttdata.com",
    "igor.gama@enel.com", "fellipe.aubynferrando@nttdata.com", "estela.pereira@enel.com", "ericson.dasilva@enel.com",
    "thiago.fcurvao@enel.com", "celine.nascimentoferreira@nttdata.com", "adriana.mottamarques@nttdata.com",
    "thainar.silveriosilva@nttdata.com", "robson.pereira@enel.com"
]
'''
# ==========================================
# FUNÇÕES AUXILIARES
# ==========================================
def capturar_imagem_excel(caminho_excel, nome_aba, caminho_imagem_saida):
    excel = None
    wb = None
    try:
        excel = win32.DispatchEx('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        caminho_absoluto = os.path.abspath(caminho_excel)
        wb = excel.Workbooks.Open(caminho_absoluto)
        ws = wb.Sheets(nome_aba)
        ultima_linha = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
        ws.Range(f"A1:AB{ultima_linha}").CopyPicture(Format=2)
        time.sleep(1.5)
        img = ImageGrab.grabclipboard()
        if img is not None:
            img.save(caminho_imagem_saida, 'PNG')
            return True
        return False
    except Exception as e:
        print(f"Erro captura: {e}")
        return False
    finally:
        if wb: wb.Close(SaveChanges=False)
        if excel: excel.Quit()

def enviar_email(caminho_arquivo, data_referencia_str, emails_to, emails_cc, caminho_imagem=None):
    try:
        try: outlook = win32.GetActiveObject("Outlook.Application")
        except: outlook = win32.DispatchEx('Outlook.Application')
        mail = outlook.CreateItem(0)
        mail.To = "; ".join(emails_to)
        if emails_cc: mail.CC = "; ".join(emails_cc)
        mail.Subject = f"Relatório - Implantação de Clandestinos - {datetime.now().strftime('%d/%m/%Y')}"

        html_imagem = ""
        if caminho_imagem and os.path.exists(caminho_imagem):
            attachment_img = mail.Attachments.Add(os.path.abspath(caminho_imagem))
            attachment_img.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", "print_resumo")
            html_imagem = '<br><img src="cid:print_resumo" style="max-width: 100%; height: auto; border: 1px solid #ccc;"><br>'
        else:
            html_imagem = "<p><i>[Não foi possível carregar o print do resumo nesta mensagem]</i></p>"

        mail.HTMLBody = f"""
        <html>
            <body style="font-family: Arial, sans-serif; font-size: 14px;">    
                <p>Olá,</p>
                <p>Seguem os indicadores de produção de clientes implantados referente ao dia {data_referencia_str}.</p>
                {html_imagem}
                <br><p><b>Data/Hora:</b><br>{datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
                <br><p><b>João Lucas L. P. Antunes</b><br>Clandestine Connections Management<br>Customer Engagement Rio<br>I&N Rio<br></p>
            </body>
        </html>
        """
        mail.Attachments.Add(os.path.abspath(caminho_arquivo))
        mail.Send()
        return True, None
    except Exception as e:
        return False, traceback.format_exc()

def executar_automacao_relatorio_backend(arquivo_ordens, pasta_saida, callback_progresso, emails_to, emails_cc):
    import pythoncom # Importado aqui dentro para garantir a rota da thread web
    
    callback_progresso(10, "Iniciando processamento do relatório...")
    ordens_df = pd.read_excel(arquivo_ordens, sheet_name=0, dtype=str)
    
    colunas_utilizadas = ["numero_ordem", "numero_cliente", "municipio", "POLO", "data_finalizacao1", "numero_ordem_filha", "estado_ordem_filha"]
    base_importacao = ordens_df[[c for c in colunas_utilizadas if c in ordens_df.columns]].copy()
    base_importacao.drop_duplicates(inplace=True)
    base_importacao.dropna(how="all", inplace=True)

    base_importacao["data_finalizacao1_dt"] = pd.to_datetime(base_importacao["data_finalizacao1"], dayfirst=True, errors="coerce")
    mascara_mes_ano = (
        (base_importacao["data_finalizacao1_dt"].dt.month == DATA_REFERENCIA_D1.month) & 
        (base_importacao["data_finalizacao1_dt"].dt.year == DATA_REFERENCIA_D1.year) &
        (base_importacao["data_finalizacao1_dt"].dt.day <= DATA_REFERENCIA_D1.day)
    )
    base_importacao = base_importacao[mascara_mes_ano].copy()
    base_importacao["dia"] = base_importacao["data_finalizacao1_dt"].dt.day

    if "estado_ordem_filha" in base_importacao.columns:
        filtro = base_importacao["estado_ordem_filha"].fillna("").str.upper().str.strip()
        base_importacao = base_importacao[filtro.str.contains("IMPLANTADO", na=False)]

    base_importacao_d1 = base_importacao[base_importacao["dia"] == DATA_REFERENCIA_D1.day].copy()
    base_importacao["data_finalizacao1"] = base_importacao["data_finalizacao1_dt"].dt.strftime('%d/%m/%Y').fillna("")
    base_importacao.drop(columns=["data_finalizacao1_dt"], inplace=True)

    callback_progresso(30, "Lendo Ordens Consolidadas...")
    df_ordens_consolidadas = pd.read_excel(arquivo_ordens, sheet_name="Ordens Consolidadas")
    df_filtrado = df_ordens_consolidadas[df_ordens_consolidadas["estado_ordem"] == "ENVIADO E-ORDER"]
    contagem_polos = df_filtrado.groupby("POLO").size().to_dict()

    polos_mapa = {"CAMPOS": "CAMPOS", "LAGOS": "LAGOS", "MACAE":"MACAE", "NOROESTE": "NOROESTE", "MAGÉ": "MAGÉ", "NITEROI": "NITEROI", "SÃO GONÇALO": "SÃO GONÇALO", "SERRANA": "SERRANA", "SUL": "SUL"}

    callback_progresso(45, "Baixando template atualizado da nuvem...")
    
    # URL direta para o seu bucket público
    url_download = f"{SUPABASE_URL}/storage/v1/object/public/templates/Relatorio.xlsx"
    
    # Baixa o arquivo ignorando o firewall da empresa
    resposta_download = requests.get(url_download, verify=False)
    
    if resposta_download.status_code == 200:
        arquivo_bytes = resposta_download.content
    else:
        raise Exception("Erro ao baixar o template base da nuvem.")
    
    caminho_absoluto_base = os.path.join(tempfile.gettempdir(), "Relatorio_Template_Temp.xlsx")
    
    with open(caminho_absoluto_base, "wb") as f:
        f.write(arquivo_bytes)

    callback_progresso(50, "Abrindo relatório base...")

    nome_arquivo = f"Relatorio_Implantação_Clandestino_{datetime.now().strftime('%d%m%Y')}.xlsx"
    caminho_saida = os.path.join(pasta_saida, nome_arquivo)
    
    excel = None
    wb = None
    try:
        # A LINHA MÁGICA FOI MOVIDA PARA CÁ (Imediatamente antes de chamar o Excel)
        pythoncom.CoInitialize() 
        excel = win32.DispatchEx('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(caminho_absoluto_base)

        callback_progresso(55, "Atualizando aba 'Meta NL Livres' e 'Meta Polos'...")
        ws_meta = wb.Sheets("Meta NL Livres")
        ultimo_dia_mes = calendar.monthrange(DATA_REFERENCIA_D1.year, DATA_REFERENCIA_D1.month)[1]
        
        dias_uteis_ate_d1 = sum(1 for dia in range(1, DATA_REFERENCIA_D1.day + 1) if datetime(DATA_REFERENCIA_D1.year, DATA_REFERENCIA_D1.month, dia).weekday() < 5 and datetime(DATA_REFERENCIA_D1.year, DATA_REFERENCIA_D1.month, dia).date() not in FERIADOS_2026)

        # Atualização Meta NL Livres
        for i, linha in enumerate(ws_meta.UsedRange.Value or []):
            for j, celula in enumerate(linha):
                if celula:
                    v = str(celula).strip().lower()
                    if v == "data atualização": ws_meta.Cells(i + 2, j + 1).Value = f"'{DATA_REFERENCIA_STR}"
                    elif v == "dias úteis": ws_meta.Cells(i + 4, j + 1).Value = dias_uteis_ate_d1

        callback_progresso(60, "Atualizando aba 'BASE IMP MÊS'...")
        ws_base = wb.Sheets("BASE IMP MÊS")
        ws_base.Columns("E").NumberFormat = "@" 
        ultima_linha = ws_base.UsedRange.Rows.Count
        if ultima_linha > 1: ws_base.Range(ws_base.Cells(2, 1), ws_base.Cells(ultima_linha + 10, 50)).ClearContents()
        
        if not base_importacao.empty:
            dados_lista = base_importacao.fillna("").values.tolist()
            if dados_lista:
                ws_base.Range(ws_base.Cells(2, 1), ws_base.Cells(1 + len(dados_lista), len(dados_lista[0]))).Value = dados_lista

        callback_progresso(70, "Atualizando aba '01 - Resumo'...")
        ws_resumo = wb.Sheets("01 - Resumo")
        for i, linha in enumerate(ws_resumo.UsedRange.Value or []):
            for j, celula in enumerate(linha):
                if celula and "data" in str(celula).strip().lower():
                    ws_resumo.Cells(i + 1, j + 1).Value = f"Data: {DATA_REFERENCIA_STR}"
                    break
        
        callback_progresso(85, "Salvando arquivo consolidado...")
        wb.SaveAs(os.path.abspath(caminho_saida))
    finally:
        if wb: wb.Close(SaveChanges=False)
        if excel: excel.Quit()

    # ========================================
    # TRAVA DE SEGURANÇA DA WEB (Corrigida e sem duplicações!)
    # ========================================
    
    # Se as listas estiverem vazias, o código sabe que está na Web (Pula Imagens e Outlook)
    if not emails_to and not emails_cc:
        callback_progresso(100, "Relatório Excel gerado com sucesso!")
        return caminho_saida
        
    # Se não estiver vazio, o código sabe que está no .exe de computador
    else:
        callback_progresso(88, "Gerando imagem do resumo...")
        
        # Garante a permissão do Windows antes de capturar a imagem também!
        pythoncom.CoInitialize() 
        caminho_imagem = os.path.join(pasta_saida, "print_temp_resumo.png")
        sucesso_img = capturar_imagem_excel(caminho_saida, "01 - Resumo", caminho_imagem)
        
        callback_progresso(95, "Abrindo o Outlook para envio de e-mail...")
        caminho_imagem_anexo = caminho_imagem if sucesso_img else None
        
        sucesso_email, erro_email = enviar_email(caminho_saida, DATA_REFERENCIA_STR, emails_to, emails_cc, caminho_imagem_anexo)
        
        # Remove a imagem temporária do disco após anexar
        if caminho_imagem and os.path.exists(caminho_imagem):
            try: os.remove(caminho_imagem)
            except: pass

        callback_progresso(100, "Processo finalizado!")
        if sucesso_email:
            return "Relatório gerado e e-mail aberto no Outlook com sucesso!"
        else:
            return f"Relatório gerado, mas falhou ao abrir o e-mail.\nErro: {erro_email}"