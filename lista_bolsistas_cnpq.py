# -*- coding: utf-8 -*-
"""
Created on Thu Oct 22 11:19:12 2026

@author: levi1
"""

import time
import pandas as pd
import requests
from io import BytesIO

# Dicionário com os principais códigos de Áreas do Conhecimento Globais/Básicas (CNPq)
AREAS_CONHECIMENTO = {
    # Ciências Exatas e da Terra
    "10100008": "MATEMÁTICA",
    "10200002": "PROBABILIDADE E ESTATÍSTICA",
    "10300007": "CIÊNCIA DA COMPUTAÇÃO",
    "10400001": "ASTRONOMIA",
    "10500006": "FÍSICA",
    "10600000": "QUÍMICA",
    "10700005": "GEOCIÊNCIAS",
    "10800000": "OCEANOGRAFIA",
    # Ciências Biológicas
    "20100000": "BIOLOGIA GERAL",
    "20200005": "GENÉTICA",
    "20300000": "BOTÂNICA",
    "20400004": "ZOOLOGIA",
    "20500009": "ECOLOGIA",
    "20600003": "MORFOLOGIA",
    "20700008": "FISIOLOGIA",
    "20800002": "BIOQUÍMICA",
    "20900007": "BIOFÍSICA",
    "21000000": "FARMACOLOGIA",
    "21100004": "IMUNOLOGIA",
    "21200009": "MICROBIOLOGIA",
    "21300003": "PARASITOLOGIA",
    # Engenharias
    "30100003": "ENGENHARIA CIVIL",
    "30200008": "ENGENHARIA DE MINAS",
    "30300002": "ENGENHARIA DE MATERIAIS E METALÚRGICA",
    "30400007": "ENGENHARIA ELÉTRICA",
    "30500001": "ENGENHARIA MECÂNICA",
    "30600006": "ENGENHARIA QUÍMICA",
    "30700000": "ENGENHARIA SANITÁRIA",
    "30800005": "ENGENHARIA DE PRODUÇÃO",
    "30900000": "ENGENHARIA NUCLEAR",
    "31000002": "ENGENHARIA DE TRANSPORTES",
    "31100007": "ENGENHARIA NAVAL E OCEÂNICA",
    "31200001": "ENGENHARIA AEROESPACIAL",
    "31300006": "ENGENHARIA BIOMÉDICA",
    # Ciências da Saúde
    "40100006": "MEDICINA",
    "40200000": "ODONTOLOGIA",
    "40300005": "FARMÁCIA",
    "40400000": "ENFERMAGEM",
    "40500004": "NUTRIÇÃO",
    "40600009": "SAÚDE COLETIVA",
    "40700003": "FONOAUDIOLOGIA",
    "40800008": "FISIOTERAPIA E TERAPIA OCUPACIONAL",
    "40900002": "EDUCAÇÃO FÍSICA",
    # Ciências Agrárias
    "50100009": "AGRONOMIA",
    "50200003": "RECURSOS FLORESTAIS E ENGENHARIA FLORESTAL",
    "50300008": "ENGENHARIA AGRÍCOLA",
    "50400002": "ZOOTECNIA",
    "50500007": "MEDICINA VETERINÁRIA",
    "50600001": "RECURSOS PESQUEIROS E ENGENHARIA DE PESCA",
    "50700006": "CIÊNCIA E TECNOLOGIA DE ALIMENTOS",
    # Ciências Sociais Aplicadas
    "60100001": "DIREITO",
    "60200006": "ADMINISTRAÇÃO",
    "60300000": "ECONOMIA",
    "60400005": "ARQUITETURA E URBANISMO",
    "60500000": "PLANEJAMENTO URBANO E REGIONAL",
    "60600004": "DEMOGRAFIA",
    "60700009": "CIÊNCIA DA INFORMAÇÃO",
    "60800003": "MUSEOLOGIA",
    "60900008": "COMUNICAÇÃO",
    "61000000": "SERVIÇO SOCIAL",
    "61100005": "ECONOMIA DOMÉSTICA",
    "61200000": "DESENHO INDUSTRIAL",
    "61300004": "TURISMO",
    # Ciências Humanas
    "70100004": "FILOSOFIA",
    "70200009": "SOCIOLOGIA",
    "70300003": "ANTROPOLOGIA",
    "70400008": "ARQUEOLOGIA",
    "70500002": "HISTÓRIA",
    "70600007": "GEOGRAFIA",
    "70700001": "PSICOLOGIA",
    "70800006": "EDUCAÇÃO",
    "70900000": "CIÊNCIA POLÍTICA",
    "71000003": "TEOLOGIA",
    # Lingüística, Letras e Artes
    "80100007": "LINGÜÍSTICA",
    "80200001": "LETRAS",
    "80300006": "ARTES",
}

BASE_URL = "http://plsql1.cnpq.br/divulg/RESULTADO_PQ_102003.prc_comp_cmt_links"

headers = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
}


def baixar_dados_bolsistas():
  lista_dfs = []

  for cod_area, nome_area in AREAS_CONHECIMENTO.items():
    url = f"{BASE_URL}?V_COD_DEMANDA=200310&V_TPO_RESULT=CURSO&V_COD_AREA_CONHEC={cod_area}&V_COD_CMT_ASSESSOR=PE"
    print(f"Coletando: [{cod_area}] {nome_area}...")

    try:
      response = requests.get(url, headers=headers, timeout=15)

      if response.status_code == 200:
        # Lê tabelas HTML da página
        
        tabelas = pd.read_html(
            BytesIO(response.content),
            encoding="ISO-8859-1"
        )
        

        if tabelas and len(tabelas) > 0:
          df_area = tabelas[5]

          # Adiciona colunas identificadoras
          df_area["Código Área"] = cod_area
          df_area["Área do Conhecimento"] = nome_area
          df_area["URL"] = url

          lista_dfs.append(df_area)
        else:
          print(f"  -> Nenhuma tabela encontrada para {nome_area}.")
      else:
        print(f"  -> Erro HTTP {response.status_code}")

    except Exception as e:
      print(f"  -> Falha na requisição para {nome_area}: {e}")

    # Intervalo para não sobrecarregar o servidor do CNPq
    time.sleep(1)

  if lista_dfs:
    df_consolidado = pd.concat(lista_dfs, ignore_index=True)

    # Exporta para Excel
    arquivo_saida = "bolsistas_pq_cnpq_todas_areas.xlsx"
    df_consolidado.to_excel(arquivo_saida, index=False, engine="openpyxl")
    print(f"\nColeta finalizada! Planilha gerada com sucesso: {arquivo_saida}")
  else:
    print("\nNenhum dado pôde ser coletado.")


if __name__ == "__main__":
  baixar_dados_bolsistas()