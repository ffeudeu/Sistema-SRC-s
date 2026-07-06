import os
import zipfile
import re
import warnings
import time
import random
import tempfile
import urllib.request
import geopandas as gpd
import pandas as pd
from shapely.geometry import Point, Polygon
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import folium
from html2image import Html2Image
from bs4 import BeautifulSoup
import requests
from geopy.geocoders import Nominatim


# Ignora avisos
warnings.filterwarnings('ignore')
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

SUPABASE_URL = "https://guqwynqjgqtfnqhcxbgn.supabase.co"

# ==========================================
# FUNÇÃO PARA BAIXAR ARQUIVOS DA NUVEM
# ==========================================
def baixar_do_supabase(nome_arquivo):
    url = f"{SUPABASE_URL}/storage/v1/object/public/templates/{nome_arquivo}"
    # url com codificação para espaços no nome do arquivo
    url = url.replace(" ", "%20")
    
    resp = requests.get(url, verify=False)
    if resp.status_code == 200:
        caminho_temp = os.path.join(tempfile.gettempdir(), nome_arquivo)
        with open(caminho_temp, 'wb') as f:
            f.write(resp.content)
        return caminho_temp
    else:
        raise Exception(f"Falha ao baixar '{nome_arquivo}' do Supabase. Verifique se o nome está correto no bucket 'templates'.")

# ==========================================
# FUNÇÕES DE GEOPROCESSAMENTO
# ==========================================
def ler_kml_ou_kmz(caminho_arquivo):
    if zipfile.is_zipfile(caminho_arquivo):
        with zipfile.ZipFile(caminho_arquivo, 'r') as kmz:
            for filename in kmz.namelist():
                if filename.lower().endswith('.kml'):
                    with kmz.open(filename, 'r') as f:
                        return f.read()
    else:
        with open(caminho_arquivo, 'r', encoding='utf-8', errors='ignore') as f:
            return f.read()

def extrair_termo_busca(nome_poligono):
    nome = nome_poligono.replace('.kml', '').replace('.kmz', '').replace('.KML', '').replace('.KMZ', '')
    nome = nome.replace('-', '_')
    partes = nome.split('_')
    termo = partes[-1].strip()
    return re.sub(r'\s*\d+$', '', termo).strip()

def descobrir_regional(municipio):
    if not municipio: return 'Não identificada'
    mun_lower = municipio.lower()
    cidades_sul = ['angra', 'niterói', 'niteroi', 'maricá', 'marica', 'são gonçalo', 'sao goncalo', 'magé', 'mage', 'petrópolis', 'petropolis', 'teresópolis', 'teresopolis', 'resende']
    cidades_norte = ['araruama', 'cabo frio', 'campos', 'macaé', 'macae', 'itaperuna', 'pádua', 'padua', 'cantagalo']
    if any(x in mun_lower for x in cidades_sul): return 'Sul'
    elif any(x in mun_lower for x in cidades_norte): return 'Norte'
    return 'Não identificada'

def kml_poligono_para_gdf(caminho_arquivo):
    soup = BeautifulSoup(ler_kml_ou_kmz(caminho_arquivo), 'xml')
    poligonos = []
    for placemark in soup.find_all('Placemark'):
        nome_tag = placemark.find('name')
        nome = nome_tag.text.strip() if nome_tag else os.path.basename(caminho_arquivo)
        poly_tag = placemark.find('Polygon')
        if poly_tag:
            coords_tag = poly_tag.find('coordinates')
            if coords_tag:
                coords_lista = [(float(xyz.split(',')[0]), float(xyz.split(',')[1])) for xyz in coords_tag.text.strip().split() if len(xyz.split(',')) >= 2]
                if len(coords_lista) >= 3:
                    poligonos.append({'Name': nome, 'geometry': Polygon(coords_lista)})
    return gpd.GeoDataFrame(poligonos, crs="EPSG:4326")

def kml_pontos_para_gdf(caminho_arquivo):
    soup = BeautifulSoup(ler_kml_ou_kmz(caminho_arquivo), 'xml')
    pontos, nomes = [], []
    for placemark in soup.find_all('Placemark'):
        pt_tag = placemark.find('Point')
        desc_tag = placemark.find('description')
        nome_pt = desc_tag.text.strip() if desc_tag else "Sem Identificação"
        if pt_tag and pt_tag.find('coordinates'):
            xyz = pt_tag.find('coordinates').text.strip().split(',')
            if len(xyz) >= 2:
                pontos.append(Point(float(xyz[0]), float(xyz[1])))
                nomes.append(nome_pt)
    return gpd.GeoDataFrame({'Nome': nomes, 'geometry': pontos}, crs="EPSG:4326")

def processar_clientes_faturados(poligono_row, caminho_clientes):
    clientes_gdf = kml_pontos_para_gdf(caminho_clientes)
    if clientes_gdf.empty: return 0, []
    clientes_dentro = clientes_gdf[clientes_gdf.geometry.within(poligono_row.geometry)]
    lista_ordens = [{'nome': row['Nome'], 'lat': round(row.geometry.y, 6), 'lon': round(row.geometry.x, 6)} for _, row in clientes_dentro.iterrows()]
    return len(clientes_dentro), lista_ordens

def buscar_ocorrencias(poligono_row, caminho_crimes, caminho_policia):
    try: df_crimes = pd.read_excel(caminho_crimes)
    except: df_crimes = pd.DataFrame()
    try: df_policia = pd.read_excel(caminho_policia)
    except: df_policia = pd.DataFrame()
    
    todas_ocorrencias = pd.concat([df_crimes, df_policia], ignore_index=True)
    termo_busca = extrair_termo_busca(poligono_row['Name'])
    
    ocorrencias_finais = pd.DataFrame()
    if termo_busca and not todas_ocorrencias.empty:
        mask = pd.Series(False, index=todas_ocorrencias.index)
        for col in todas_ocorrencias.columns:
            mask = mask | todas_ocorrencias[col].fillna('').astype(str).str.contains(termo_busca, case=False, regex=False)
        ocorrencias_finais = todas_ocorrencias[mask]
    
    lista_noticias = []
    col_desc = next((c for c in ocorrencias_finais.columns if any(x in c.lower() for x in ['descri', 'título', 'address', 'titulo'])), None)
    col_data = next((c for c in ocorrencias_finais.columns if 'data' in c.lower()), None)
    col_fonte = next((c for c in ocorrencias_finais.columns if any(x in c.lower() for x in ['fonte', 'link', 'url'])), None)

    for _, row in ocorrencias_finais.head(5).iterrows():
        titulo = str(row[col_desc]).replace('_x000d_', '').replace('\r', ' ').replace('\n', ' ').strip() if col_desc and pd.notna(row[col_desc]) else "Ocorrência registrada"
        data = pd.to_datetime(row[col_data]).strftime('%d/%m/%Y') if col_data and pd.notna(row[col_data]) else "Data não informada"
        fonte = str(row[col_fonte]).strip() if col_fonte and pd.notna(row[col_fonte]) else "Base de Dados Interna"
        lista_noticias.append({'titulo': titulo, 'data': data, 'fonte': fonte})
    
    return lista_noticias, len(lista_noticias) > 0

def verificar_restricao_correios(poligono_geom, lista_ordens):
    TOKEN_MELHOR_ENVIO = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJhdWQiOiIxIiwianRpIjoiYWYwZDdhZGQyOGU2NTMyZTllMDRiNDgyMGQ1OTlkNTFiNDdiYTdlZjIxZWVlOGM5ZWM4ZjNkNzllZGI0YTA4NjI3MjkyMTc1ZDAzZGY2ODAiLCJpYXQiOjE3ODMwODU5NTIuNjcyMDY5LCJuYmYiOjE3ODMwODU5NTIuNjcyMDcxLCJleHAiOjE4MTQ2MjE5NTIuNjYxOTYyLCJzdWIiOiJhMjJiYWU2Yy1lN2M0LTQzMmQtYjFkZi05ZWRhZDc0OWJhMzYiLCJzY29wZXMiOlsiY2FydC1yZWFkIiwiY2FydC13cml0ZSIsImNvbXBhbmllcy1yZWFkIiwiY29tcGFuaWVzLXdyaXRlIiwiY291cG9ucy1yZWFkIiwiY291cG9ucy13cml0ZSIsIm5vdGlmaWNhdGlvbnMtcmVhZCIsIm9yZGVycy1yZWFkIiwicHJvZHVjdHMtcmVhZCIsInByb2R1Y3RzLWRlc3Ryb3kiLCJwcm9kdWN0cy13cml0ZSIsInB1cmNoYXNlcy1yZWFkIiwic2hpcHBpbmctY2FsY3VsYXRlIiwic2hpcHBpbmctY2FuY2VsIiwic2hpcHBpbmctY2hlY2tvdXQiLCJzaGlwcGluZy1jb21wYW5pZXMiLCJzaGlwcGluZy1nZW5lcmF0ZSIsInNoaXBwaW5nLXByZXZpZXciLCJzaGlwcGluZy1wcmludCIsInNoaXBwaW5nLXNoYXJlIiwic2hpcHBpbmctdHJhY2tpbmciLCJlY29tbWVyY2Utc2hpcHBpbmciLCJ0cmFuc2FjdGlvbnMtcmVhZCIsInVzZXJzLXJlYWQiLCJ1c2Vycy13cml0ZSIsIndlYmhvb2tzLXJlYWQiLCJ3ZWJob29rcy13cml0ZSIsIndlYmhvb2tzLWRlbGV0ZSIsInRkZWFsZXItd2ViaG9vayJdfQ.xLVfhDjbok9-Cjolqw_ZWrK65zrOO8OW4P40ai6ylUybKWI2tJJoRPthGa9hIeAsCQki22sSTApTcz9c2ijDHgB-F3-KaYu8jHoHg1GObP7pP_iJMnDq0-DvK432hrChiv7AXlt_CYAgAxId7ToU5SG6GTufspr2r77fHQIyH-51pRpLXXjlihjyg5Sis3zdsc-Dw3DZxm5ADkSgwswbfyLNq7OfiBfm9WncsOqcry8jzN9sDBAR3NXRMy1GBUK1IH6kxd0KqsP-xChU974fgXEd_iKZFE_jdsw_kdm-KmY9-eAvkhtipwuVRDJNXhB1uTWiTHbuKWFkOWZuHV0Ut34naikDXlRF1Hz7J9GcJTAGCs8Ksv8JC9gpr5a3UfaMXc-MzExfEbwIjQTtNN2WR2NLGGwpjMNFHIi8LakrzzwRrolriXvSnELgZOArfmGTwCL0Or-_TcTAuWyEBDzJQkvyiE3U_TB4zELtmDFZdw_GXwUTtRVIb4FUvFGCOcS3JO23vf-EefrFVISXZB3BV17ilUaaqfIQMpbIeYUuUmtxuWWY7n_S1JiWkmzF5OILKQlpA-ryM_Ix_h_4wvBZ2X1M1d9O7Rzg7JUdQi1ynS6Uq2AftfoIdlyoL3a46X0FLWd-ea9UKybhYZ39qZ8r9BScf-adStAR6teM1s2jzEY"
    geolocator = Nominatim(user_agent="acac_enel_automacao", timeout=15)
    pontos = [poligono_geom.centroid] + [Point(o['lon'], o['lat']) for o in lista_ordens[:15]]
    
    minx, miny, maxx, maxy = poligono_geom.bounds
    while len(pontos) < 20:
        pnt = Point(random.uniform(minx, maxx), random.uniform(miny, maxy))
        if poligono_geom.contains(pnt): pontos.append(pnt)

    ceps_ja_testados = set()
    for pnt in pontos:
        try:
            time.sleep(1)
            location = geolocator.reverse(f"{pnt.y}, {pnt.x}")
            cep = location.raw['address'].get('postcode', '').replace('-', '').strip()
            if not cep or len(cep) != 8 or not cep.isdigit() or cep in ceps_ja_testados: continue
            ceps_ja_testados.add(cep)
            
            bairro_cep, municipio_cep = "", ""
            try:
                resp_viacep = requests.get(f"https://viacep.com.br/ws/{cep}/json/", timeout=5, verify=False)
                if resp_viacep.status_code == 200 and not resp_viacep.json().get('erro'):
                    bairro_cep = resp_viacep.json().get('bairro', '')
                    municipio_cep = resp_viacep.json().get('localidade', '')
            except: pass
            
            if TOKEN_MELHOR_ENVIO.startswith("COLE_SEU_TOKEN"): continue
                
            resp = requests.post("https://www.melhorenvio.com.br/api/v2/me/shipment/calculate", 
                                 json={"from": {"postal_code": "24859488"}, "to": {"postal_code": cep}, "package": {"weight": 1, "width": 20, "height": 20, "length": 20}},
                                 headers={"Accept": "application/json", "Content-Type": "application/json", "Authorization": f"Bearer {TOKEN_MELHOR_ENVIO}"}, timeout=10, verify=False)
            
            if resp.status_code == 200:
                dados = resp.json()
                pac = next((s for s in dados if s.get('id') == 1), None)
                if pac and pac.get('error') and any(x in pac.get('error').lower() for x in ["restrição", "área de risco", "domiciliar"]):
                    return True, cep, "O CEP de destino está com entrega domiciliária restrita.", bairro_cep, municipio_cep
                return False, cep, "", bairro_cep, municipio_cep
        except: continue
    return False, None, "", "", ""

def gerar_print_mapa(poligono_row):
    centro = poligono_row.geometry.centroid
    m = folium.Map(location=[centro.y, centro.x], zoom_start=15, tiles='CartoDB Positron')
    folium.GeoJson(poligono_row.geometry, style_function=lambda x: {'fillColor': '#cccc00', 'color': '#cccc00', 'fillOpacity': 0.5}).add_to(m)
    
    temp_dir = tempfile.gettempdir()
    html_path = os.path.join(temp_dir, "mapa_temp.html")
    png_path = os.path.join(temp_dir, "mapa_temp.png")
    
    m.save(html_path)
    hti = Html2Image(output_path=temp_dir)
    hti.screenshot(html_file=html_path, save_as="mapa_temp.png", size=(800, 500))
    return png_path

# ==========================================
# MOTOR PRINCIPAL CHAMADO PELO MAIN.PY
# ==========================================
def executar_automacao_breve_backend(caminho_poligono, pasta_saida, callback_progresso):
    
    # Faz o download das bases (silenciosamente) da nuvem
    callback_progresso(10, "Baixando base de clientes (Z94) da Nuvem...")
    kml_clientes = baixar_do_supabase("Todos Z94 2025.kml")
    
    callback_progresso(25, "Baixando bases de Ocorrências da Nuvem...")
    xlsx_crimes = baixar_do_supabase("Oficio Aneel_RJ_Crimes.xlsx")
    xlsx_policia = baixar_do_supabase("Oficio Aneel_RJ_Atividade Policial.xlsx")
    
    callback_progresso(35, "Baixando Template do Ofício...")
    template_word = baixar_do_supabase("Script ACAC.docx")

    callback_progresso(45, "Lendo e processando o Polígono...")
    gdf_poligono = kml_poligono_para_gdf(caminho_poligono)
    if gdf_poligono.empty:
        raise Exception("Nenhum polígono válido encontrado no arquivo fornecido.")
        
    poligono_alvo = gdf_poligono.iloc[0]
    nome_area_completo = poligono_alvo['Name']
    if nome_area_completo == 'Area_ACAC': 
        nome_area_completo = os.path.basename(caminho_poligono).split('.')[0]

    termo_bairro = extrair_termo_busca(nome_area_completo)
    area_km2 = round(poligono_alvo.geometry.area * 10000, 2) 

    callback_progresso(60, "Cruzando polígono com base de clientes Z94...")
    qtd_clientes, lista_ordens = processar_clientes_faturados(poligono_alvo, kml_clientes)
    
    callback_progresso(70, "Buscando ocorrências policiais e criminais...")
    noticias, tem_ocorrencias = buscar_ocorrencias(poligono_alvo, xlsx_crimes, xlsx_policia)
    
    callback_progresso(80, "Verificando restrições de entrega via API (Correios)...")
    tem_restricao_correios, cep_area, msg_correios, bairro_viacep, municipio_viacep = verificar_restricao_correios(poligono_alvo.geometry, lista_ordens)
    
    municipio_final = municipio_viacep if municipio_viacep else "Não identificado"
    bairro_final = bairro_viacep if bairro_viacep else termo_bairro.capitalize()
    regional_final = descobrir_regional(municipio_final)
    
    callback_progresso(90, "Desenhando mapa e gerando arquivo Word...")
    caminho_imagem_mapa = gerar_print_mapa(poligono_alvo)
    
    doc = DocxTemplate(template_word)
    imagem_mapa_inline = InlineImage(doc, caminho_imagem_mapa, width=Mm(160))
    
    contexto = {
        'nome_area': nome_area_completo, 'area_km2': f"{area_km2} km²", 'qtd_clientes': qtd_clientes,
        'regional': regional_final, 'municipio': municipio_final, 'bairro': bairro_final,
        'mapa_poligono': imagem_mapa_inline, 'x_gestao_nula': ' ', 'x_gestao_parcial': 'X', 
        'x_geo_sim': 'X', 'x_geo_nao': ' ', 'x_trafico_sim': 'X', 'x_trafico_nao': ' ',
        'x_barricada_sim': ' ', 'x_barricada_nao': 'X',
        'x_crimes_sim': 'X' if tem_ocorrencias else ' ', 'x_crimes_nao': ' ' if tem_ocorrencias else 'X',
        'x_policia_sim': 'X' if tem_ocorrencias else ' ', 'x_policia_nao': ' ' if tem_ocorrencias else 'X',
        'x_evento_colab_sim': ' ', 'x_evento_colab_nao': 'X', 'x_impedimento_sim': ' ', 'x_impedimento_nao': 'X',
        'x_reincidencia_sim': ' ', 'x_reincidencia_nao': 'X', 'x_perdas_sim': 'X', 'x_perdas_nao': ' ',
        'noticias': noticias, 'ordens': lista_ordens,
        'print_correios': "Restrição Domiciliar Identificada" if tem_restricao_correios else "Sem restrições Correios."
    }
    
    doc.render(contexto)
    nome_arquivo_limpo = "".join(x for x in nome_area_completo if x.isalnum() or x in " -_")
    caminho_final = os.path.join(pasta_saida, f"ACAC_{nome_arquivo_limpo}.docx")
    doc.save(caminho_final)
    
    callback_progresso(100, "Ofício gerado com sucesso!")
    return caminho_final
