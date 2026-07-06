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

def salvar_e_formatar_excel(df, caminho, nome_planilha="Planilha1"):
    """Salva o DataFrame em Excel aplicando formatação visual com a biblioteca OpenPyXL."""
    with pd.ExcelWriter(caminho, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=nome_planilha)
        worksheet = writer.sheets[nome_planilha]

        # Estilo do cabeçalho
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
# MOTOR PRINCIPAL (Roda em memória)
# ==========================================
def executar_automacao_ccp_backend(arquivo_poligonos, arquivo_clientes, callback_progresso):
    
    callback_progresso(10, "Lendo polígonos...")
    poligonos = abrir_kml_ou_kmz(arquivo_poligonos)

    callback_progresso(30, "Lendo clientes...")
    clientes = abrir_kml_ou_kmz(arquivo_clientes)

    callback_progresso(50, "Padronizando sistemas e limpando geometrias...")
    poligonos = poligonos.to_crs(4326)
    clientes = clientes.to_crs(4326)

    if hasattr(poligonos.geometry, 'force_2d'): poligonos.geometry = poligonos.geometry.force_2d()
    if hasattr(clientes.geometry, 'force_2d'): clientes.geometry = clientes.geometry.force_2d()
    poligonos.geometry = poligonos.geometry.make_valid()

    # Identificação inteligente
    def encontrar_melhor_coluna(df):
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
    col_cli_orig = encontrar_melhor_coluna(clientes)

    callback_progresso(70, "Realizando junção espacial (isso pode demorar)...")
    resultado = gpd.sjoin(clientes, poligonos, how="left", predicate="intersects")

    col_cli_nome = f"{col_cli_orig}_left" if f"{col_cli_orig}_left" in resultado.columns else col_cli_orig
    col_poli_nome = f"{col_poli_orig}_right" if f"{col_poli_orig}_right" in resultado.columns else col_poli_orig

    def extrair_coordenadas(geom):
        if not geom or geom.is_empty: return ""
        try: return f"{geom.y:.6f}, {geom.x:.6f}"
        except Exception: return f"{geom.centroid.y:.6f}, {geom.centroid.x:.6f}"
            
    resultado["Coord_Lat_Lon"] = resultado.geometry.apply(extrair_coordenadas)

    callback_progresso(85, "Preparando dados e resultados...")
    resumo = resultado[[col_poli_nome, col_cli_nome, "Coord_Lat_Lon"]].copy()
    resumo[col_poli_nome] = resumo[col_poli_nome].fillna("Fora dos Polígonos")
    resumo = resumo.rename(columns={col_poli_nome: "Poligono", col_cli_nome: "Cliente", "Coord_Lat_Lon": "Coordenadas (Lat, Lon)"})
    resumo = resumo.sort_values(by=["Poligono", "Cliente"])

    # Agrupa para a interface (mas sem salvar no PC!)
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
    # Agora a função retorna o dataframe 'resumo' para a memória do aplicativo
    return mensagem, dados_interface, resumo