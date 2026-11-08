"""
ranking_pesquisadores_po.py

Gera um ranking (Excel) de pesquisadores de instituições brasileiras com
fator h (Scopus) >= H_MINIMO que publicaram em temas de Pesquisa
Operacional (VRP, production scheduling, bin packing, cutting stock,
lot sizing, etc.).

Como funciona
-------------
1. Para cada tema (dicionário TEMAS) roda uma busca ScopusSearch restrita
   a documentos com ao menos uma afiliação no Brasil (AFFILCOUNTRY).
2. Em cada documento, identifica quais autores têm afiliação brasileira
   (cruzando author_afids com afid/affiliation_country) e acumula, por
   autor: artigos de PO, citações desses artigos, temas e instituição.
3. Mantém só autores com pelo menos MIN_ARTIGOS_PO artigos de PO (reduz
   o número de consultas de perfil, que é o que consome a cota da API).
4. Consulta o perfil (AuthorRetrieval) de cada candidato para obter o
   fator h, e filtra h >= H_MINIMO.
5. Para quem passou no filtro de h, lista TODOS os documentos do perfil e
   classifica cada um como "de PO" ou não. Só ficam no ranking autores com
   pelo menos FRACAO_MINIMA_PO (1/3) dos documentos do perfil em PO.
6. Remove autores com FRACAO_MAXIMA_TEMAS_EXCLUIDOS (30%) ou mais dos
   documentos do perfil nos TEMAS_EXCLUIDOS (ex.: gestão de energia, fluxo
   de potência, sistemas de distribuição elétrica).
7. Mantém só autores com ao menos um documento de PO publicado nos
   últimos ANOS_JANELA_PO anos (exclui quem não publica mais na área).
8. Mantém só autores cuja afiliação no documento mais recente do perfil
   é no Brasil (exclui quem já não está no país).
9. Para quem entra no ranking, procura a bolsa de produtividade (CNPq)
   vigente na planilha de bolsistas PQ (ARQUIVO_BOLSAS_PQ), pelo nome
   (tolerando grafias diferentes, sobrenome em outra posição e nomes
   abreviados), trazendo também a área do conhecimento.
10. Ordena e salva o Excel (aba "Ranking" + aba "Metodologia").

Classificação de um documento do perfil como "de PO"
----------------------------------------------------
Um documento conta como PO se (a) o EID dele apareceu nas buscas por tema
(passo 1), ou (b) o título, o resumo ou as palavras-chave do autor contêm
algum dos termos de TEMAS (palavra inteira, sem acentos, sem diferenciar
maiúsculas e tratando hífen como espaço). Temas com contexto em
o campo "contexto" em TEMAS (ex.: simulação de eventos discretos em
indústria, cadeia de suprimentos ou saúde) exigem também um termo de
contexto no mesmo documento. O critério (b) captura trabalhos
que a busca do passo 1 não alcança (ex.: sem coautor afiliado ao Brasil,
congressos, anos fora do filtro).

Definição das colunas
---------------------
- Fator h ............ h-index do perfil Scopus do autor (carreira toda)
- Nº de artigos ...... total de documentos do perfil Scopus (carreira toda)
- Citações totais .... total de citações do perfil Scopus (carreira toda)
- Citações máx. ...... maior nº de citações entre todos os documentos do
                       perfil (exige listar os documentos do autor: 1
                       busca extra por autor aprovado)
- Artigos de PO ...... documentos do perfil classificados como PO
- % em PO ............ Artigos de PO / total de documentos do perfil
- Universidade ....... afiliação brasileira do autor no documento mais
                       recente do perfil

Requisitos:
    pip install pybliometrics pandas openpyxl
"""

import re
import time
import unicodedata
from collections import defaultdict

from datetime import date, datetime
from difflib import SequenceMatcher
from functools import lru_cache

import pandas as pd
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties
from pybliometrics.scopus import AuthorRetrieval, ScopusSearch, init

init()

# =========================================================================
# CONFIGURAÇÃO
# =========================================================================
ARQUIVO_SAIDA = "ranking_pesquisadores_po_brasil_completo.xlsx"

H_MINIMO = 5            # fator h mínimo para entrar no ranking
MIN_ARTIGOS_PO = 2      # mínimo de artigos de PO para consultar o perfil
                        # (1 = mais completo, porém MUITO mais consultas)
ANO_INICIO = None       # ex.: 2015  (None = sem limite)
ANO_FIM = None          # ex.: 2026  (None = sem limite)
TIPOS_DOCUMENTO = ["ar", "re"]   # ar = artigo, re = revisão, cp = congresso
PAUSA_PERFIL = 0.1      # segundos entre consultas de perfil
ARQUIVO_BOLSAS_PQ = "bolsistas_pq_cnpq_todas_areas.xlsx"   # planilha de bolsistas PQ (CNPq)
DATA_REFERENCIA = None  # data usada para escolher a vigência da bolsa (None = hoje)
# Áreas do CNPq consideradas plausíveis para pesquisadores deste ranking. Ajudam
# a aceitar um bolsista cujo nome confere mas cuja sigla de instituição não
# pôde ser relacionada à universidade do Scopus (comparação sem acentos, por
# trecho do nome da área).
AREAS_CNPQ_PLAUSIVEIS = (
    "ENGENHARIA", "COMPUTAÇÃO", "MATEMÁTICA", "ESTATÍSTICA",
    "ADMINISTRAÇÃO", "ECONOMIA",
)
# Sobrenomes que aparecem em mais bolsistas do que isto (Silva, Santos, Souza...)
# são "comuns": nesses casos a área plausível NÃO basta para aceitar uma
# bolsa quando a instituição não confere (evita atribuir a bolsa de um
# homônimo). Só a instituição compatível desempata.
LIMITE_SOBRENOME_COMUM = 20
REFRESH_DIAS = 30       # reconsulta a API se o cache tiver mais de N dias
FRACAO_MINIMA_PO = 1 / 3   # mínimo de documentos do perfil que devem ser de PO
ANOS_JANELA_PO = 7      # exige ao menos 1 documento de PO nos últimos N anos (None = sem esse filtro)
FRACAO_MAXIMA_TEMAS_EXCLUIDOS = 0.30   # remove autores com esta fração (ou mais) dos documentos do perfil em TEMAS_EXCLUIDOS (None = sem esse filtro)

# Temas de PO. Cada tema é um dicionário com:
#   "sigla":    rótulo curto usado na coluna "Temas de PO" da planilha.
#   "termos":   termos de busca (TITLE-ABS-KEY), combinados com OR.
#   "contexto": (opcional) termos de contexto de aplicação, combinados com OR.
#               Quando existe, o documento precisa ter ao menos um termo do
#               tema E ao menos um termo de contexto
#               (TITLE-ABS-KEY(termos) AND TITLE-ABS-KEY(contexto)).
#               Serve para restringir temas amplos, como a simulação de
#               eventos discretos.
# Para incluir um novo tema, basta adicionar uma entrada neste dicionário.
TEMAS = {
    "Roteamento de veículos (VRP)": {
        "sigla": "VRP",
        "termos": [
            '"vehicle routing"', '"arc routing"', '"inventory routing"',
            '"traveling salesman"', '"travelling salesman"',
        ],
    },
    "Sequenciamento da produção": {
        "sigla": "Sequenciamento",
        "termos": [
            '"production scheduling"', '"flow shop"', '"flowshop"',
            '"job shop"', '"jobshop"', '"open shop"', '"openshop"', '"machine scheduling"',
            '"parallel machine scheduling"', '"single machine scheduling"',
            '"scheduling"',
        ],
    },
    "Bin packing / empacotamento": {
        "sigla": "Bin packing",
        "termos": [
            '"bin packing"', '"bin-packing"', '"container loading"',
            '"pallet loading"', '"packing problem"',
        ],
    },
    "Cutting stock / corte": {
        "sigla": "Cutting stock",
        "termos": [
            '"cutting stock"', '"cutting-stock"', '"cutting and packing"',
            '"trim loss"', '"guillotine cutting"',
        ],
    },
    "Dimensionamento de lotes (lot sizing)": {
        "sigla": "Lot sizing",
        "termos": ['"lot sizing"', '"lot-sizing"', '"lotsizing"'],
    },
    "Localização de instalações": {
        "sigla": "Localização",
        "termos": ['"facility location"', '"location-allocation"'],
    },
    "Pesquisa operacional (geral)": {
        "sigla": "PO geral",
        "termos": [
            '"operations research"', '"operational research"',
            '"pesquisa operacional"',
        ],
    },
    "Decisão multicritério (MCDM)": {
        "sigla": "MCDM",
        "termos": [
            '"multi-criteria decision"', '"multicriteria decision"',
            '"multiple criteria decision"', '"multi-attribute decision"',
            '"multiattribute decision"', '"MCDM"',
            '"analytic hierarchy process"', '"analytic network process"',
            '"TOPSIS"', '"PROMETHEE"', '"ELECTRE"', '"VIKOR"', '"outranking"',
        ],
    },
    "Análise por Envoltória de Dados (DEA)": {
        "sigla": "DEA",
        "termos": [
            '"data envelopment"', '"envoltória de dados"',
            '"slacks-based measure"', '"Malmquist productivity"',
        ],
    },
    "Programação linear/inteira": {
        "sigla": "Prog. linear/inteira",
        "termos": [
            '"linear programming"', '"integer programming"',
            '"mixed integer"', '"MILP"', '"MINLP"', '"linear optimization"',
            '"programação linear"', '"programação inteira"',
        ],
    },
    "Programação por restrições": {
        "sigla": "Prog. restrições",
        "termos": [
            '"constraint programming"', '"constraint satisfaction"',
            '"constraint logic programming"', '"CP-SAT"',
            '"programação por restrições"', '"programação com restrições"',
        ],
    },
    "Métodos de decomposição": {
        "sigla": "Decomposição",
        "termos": [
            '"branch and bound"', '"branch-and-bound"', '"branch and cut"',
            '"branch-and-cut"', '"branch and price"', '"branch-and-price"',
            '"cutting plane method"', '"cutting plane algorithm"',
            '"cutting planes"', '"valid inequality"', '"valid inequalities"',
            '"Benders decomposition"', '"Lagrangian relaxation"',
            '"Lagrangean relaxation"', '"Lagrangian decomposition"',
            '"column generation"', '"Dantzig-Wolfe"', '"dual decomposition"',
            '"primal-dual"', '"relaxação lagrangeana"', '"geração de colunas"',
        ],
    },
    "Teoria dos grafos": {
        "sigla": "Grafos",
        "termos": [
            '"graph theory"', '"graph coloring"', '"graph colouring"',
            '"vertex coloring"', '"edge coloring"', '"chromatic number"',
            '"graph partitioning"', '"spanning tree"', '"vertex cover"',
            '"maximum clique"', '"maximum independent set"',
            '"Hamiltonian cycle"', '"Hamiltonian path"',
            '"teoria dos grafos"', '"coloração de grafos"',
        ],
    },
    "Metaheurísticas": {
        "sigla": "Metaheurísticas",
        "termos": [
            '"metaheuristic"', '"metaheuristics"', '"meta-heuristic"',
            '"meta-heuristics"', '"hyper-heuristic"', '"matheuristic"',
            '"matheuristics"',
            '"hyper-heuristics"', '"algoritmo genético"', '"metaheurística"',
        ],
    },

    "Otimização combinatória": {
        "sigla": "Ot. combinatória",
        "termos": [
            '"combinatorial optimization"', '"combinatorial optimisation"',
            '"combinatorial problem"', '"knapsack problem"',
            '"generalized assignment"', '"quadratic assignment"',
            '"set covering"', '"set partitioning"',
            '"otimização combinatória"',
        ],
    },
    "Simulação de eventos discretos - indústria": {
        "sigla": "SED indústria",
        "termos": [
            '"discrete event simulation"', '"discrete-event simulation"',
            '"simulação de eventos discretos"',
        ],
        "contexto": [
            '"manufacturing"', '"production line"', '"production system"',
            '"assembly line"', '"factory"', '"shop floor"', '"industrial"',
            '"industry"', '"lean manufacturing"', '"lean production"',
            '"manufatura"', '"indústria"', '"linha de produção"',
            '"fábrica"',
        ],
    },
    "Simulação de eventos discretos - cadeia de suprimentos": {
        "sigla": "SED suprimentos",
        "termos": [
            '"discrete event simulation"', '"discrete-event simulation"',
            '"simulação de eventos discretos"',
        ],
        "contexto": [
            '"supply chain"', '"logistics"', '"warehouse"', '"warehousing"',
            '"inventory"', '"distribution center"', '"container terminal"',
            '"port terminal"', '"cadeia de suprimentos"', '"logística"',
            '"armazém"',
            # modal aéreo / aeroportos
            '"airport"', '"airline"', '"aviation"', '"air transport"',
            '"air cargo"', '"aircraft"', '"aeroporto"', '"aviação"',
            '"transporte aéreo"', '"modal aéreo"',
        ],
    },
    "Simulação de eventos discretos - saúde e hospitais": {
        "sigla": "SED saúde",
        "termos": [
            '"discrete event simulation"', '"discrete-event simulation"',
            '"simulação de eventos discretos"',
        ],
        "contexto": [
            '"hospital"', '"healthcare"', '"health care"',
            '"emergency department"', '"emergency room"', '"patient flow"',
            '"operating room"', '"clinic"', '"outpatient"',
            '"intensive care"', '"health system"', '"saúde"', '"hospitalar"',
            '"pronto atendimento"',
        ],
    },
}


# Temas EXCLUÍDOS: autores com FRACAO_MAXIMA_TEMAS_EXCLUIDOS (ou mais) dos
# documentos do perfil nestes temas são removidos do ranking (ex.: quem
# publica sobretudo em gestão de energia, fluxo de potência e sistemas de
# distribuição elétrica). Mesma estrutura de TEMAS ("sigla", "termos" e
# "contexto" opcional), mas a classificação é feita só sobre o título, o
# resumo e as palavras-chave dos documentos do perfil: não há buscas extras
# no Scopus, portanto nenhum consumo adicional de cota. Um documento pode
# casar ao mesmo tempo com temas de PO e com temas excluídos.
# Para incluir um tema excluído, basta adicionar uma entrada neste dicionário.
TEMAS_EXCLUIDOS = {
    "Gestão de energia e fontes renováveis": {
        "sigla": "Energia",
        "termos": [
            '"energy management"', '"demand response"',
            '"demand side management"', '"smart grid"', '"microgrid"',
            '"energy storage"', '"renewable energy"', '"photovoltaic"',
            '"wind power"', '"wind energy"', '"solar energy"',
            '"gerenciamento de energia"', '"gestão de energia"',
            '"energia renovável"', '"energias renováveis"',
        ],
    },
    "Fluxo de potência e sistemas elétricos de potência": {
        "sigla": "Sist. elétricos",
        "termos": [
            '"power flow"', '"load flow"', '"power system"',
            '"electric power"', '"electrical power"', '"power grid"',
            '"unit commitment"', '"economic dispatch"',
            '"distributed generation"', '"reactive power"',
            '"voltage control"', '"power quality"', '"capacitor placement"',
            '"power distribution"', '"electrical distribution"',
            '"electric distribution"', '"fluxo de potência"',
            '"sistemas elétricos de potência"',
            '"sistema elétrico de potência"', '"despacho econômico"',
            '"geração distribuída"',
        ],
    },
    # "Distribution system/network" também é termo de logística: só conta
    # junto de um termo de energia elétrica.
    "Sistemas de distribuição (elétrica)": {
        "sigla": "Distribuição elétrica",
        "termos": [
            '"distribution system"', '"distribution network"',
            '"distribution feeder"', '"sistema de distribuição"',
            '"sistemas de distribuição"', '"rede de distribuição"',
            '"redes de distribuição"',
        ],
        "contexto": [
            '"power"', '"voltage"', '"substation"', '"electricity"',
            '"potência"', '"energia elétrica"', '"eletricidade"',
        ],
    },
}

# =========================================================================
# FUNÇÕES AUXILIARES
# =========================================================================
def eh_brasil(pais):
    return str(pais).strip().lower() in {"brazil", "brasil"}


def montar_query(termos, contexto=None):
    """Monta a query do ScopusSearch para um tema (com contexto opcional)."""
    q = "TITLE-ABS-KEY(" + " OR ".join(termos) + ")"
    if contexto:
        q += " AND TITLE-ABS-KEY(" + " OR ".join(contexto) + ")"
    q += " AND AFFILCOUNTRY(Brazil)"
    q += " AND DOCTYPE(" + " OR ".join(TIPOS_DOCUMENTO) + ")"
    if ANO_INICIO:
        q += f" AND PUBYEAR > {int(ANO_INICIO) - 1}"
    if ANO_FIM:
        q += f" AND PUBYEAR < {int(ANO_FIM) + 1}"
    return q


def normalizar_texto(texto):
    """Minúsculas, sem acentos, hífen/underscore como espaço, espaços únicos."""
    texto = unicodedata.normalize("NFKD", str(texto or ""))
    texto = texto.encode("ASCII", "ignore").decode("ASCII").lower()
    texto = re.sub(r"[-_|;/]", " ", texto)
    return re.sub(r"\s+", " ", texto).strip()


def compilar_termo(termo):
    """
    Padrão de palavra inteira para a classificação local, aceitando plural
    simples (hospital/hospitals) e y/ies (factory/factories). Evita, por
    exemplo, que "electre" case com "electret".
    """
    t = normalizar_texto(termo.strip('"'))
    if t.endswith("y"):
        corpo = re.escape(t[:-1]) + r"(?:y|ies)"
    else:
        corpo = re.escape(t) + r"(?:s|es)?"
    return re.compile(r"(?<!\w)" + corpo + r"(?!\w)")


# Padrões de cada tema (e de seu contexto, quando houver) para a
# classificação local dos documentos do perfil.
PADROES_TEMAS = {
    tema: [compilar_termo(t) for t in cfg["termos"]]
    for tema, cfg in TEMAS.items()
}
PADROES_CONTEXTO = {
    tema: [compilar_termo(t) for t in cfg["contexto"]]
    for tema, cfg in TEMAS.items() if cfg.get("contexto")
}
PADROES_EXCLUIDOS = {
    tema: [compilar_termo(t) for t in cfg["termos"]]
    for tema, cfg in TEMAS_EXCLUIDOS.items()
}
PADROES_CONTEXTO_EXCLUIDOS = {
    tema: [compilar_termo(t) for t in cfg["contexto"]]
    for tema, cfg in TEMAS_EXCLUIDOS.items() if cfg.get("contexto")
}


def _temas_no_documento(doc, padroes_temas, padroes_contexto):
    """Temas (de um conjunto de padrões) cujos termos aparecem no título/resumo/palavras-chave."""
    texto = normalizar_texto(
        " ".join(
            str(getattr(doc, campo, "") or "")
            for campo in ("title", "description", "authkeywords")
        )
    )
    encontrados = set()
    for tema, padroes in padroes_temas.items():
        if not any(p.search(texto) for p in padroes):
            continue
        contexto = padroes_contexto.get(tema)
        if contexto and not any(p.search(texto) for p in contexto):
            continue  # tema exige também um termo de contexto
        encontrados.add(tema)
    return encontrados


def temas_do_documento(doc):
    """Temas de PO cujos termos aparecem no título/resumo/palavras-chave."""
    return _temas_no_documento(doc, PADROES_TEMAS, PADROES_CONTEXTO)


def temas_excluidos_do_documento(doc):
    """Temas excluídos (TEMAS_EXCLUIDOS) presentes no título/resumo/palavras-chave."""
    return _temas_no_documento(doc, PADROES_EXCLUIDOS, PADROES_CONTEXTO_EXCLUIDOS)


def dividir(texto):
    """Divide campos do ScopusSearch unidos por ';'."""
    if not texto:
        return []
    return [p.strip() for p in str(texto).split(";")]


def afiliacoes_do_documento(d):
    """
    Afiliação de cada autor de um documento do ScopusSearch (visão COMPLETE).

    Retorna uma lista de tuplas (posicao, author_id, instituicao_brasileira),
    uma por autor cuja afiliação tem país identificável: 'instituicao_brasileira'
    é o nome da primeira afiliação do autor localizada no Brasil, ou None se
    nenhuma estiver. Autores sem afiliação identificável não aparecem.
    Retorna None se os campos de afiliação do documento forem inconsistentes.
    """
    ids = dividir(getattr(d, "author_ids", None))
    afids_autores = dividir(getattr(d, "author_afids", None))   # "afid-afid;afid;..."
    afids_doc = dividir(getattr(d, "afid", None))
    nomes_inst = dividir(getattr(d, "affilname", None))
    paises = dividir(getattr(d, "affiliation_country", None))

    # Os campos de afiliação do documento são alinhados entre si
    if not (len(afids_doc) == len(nomes_inst) == len(paises)) \
            or len(ids) != len(afids_autores):
        return None

    mapa_inst = {a: (n, p) for a, n, p in zip(afids_doc, nomes_inst, paises)}
    resultado = []
    for i, auid in enumerate(ids):
        if not auid:
            continue
        identificavel = False
        inst_brasil = None
        for af in (a for a in afids_autores[i].split("-") if a):
            nome_inst, pais = mapa_inst.get(af, (None, None))
            if not pais:
                continue
            identificavel = True
            if eh_brasil(pais):
                inst_brasil = nome_inst
                break
        if identificavel:
            resultado.append((i, auid, inst_brasil))
    return resultado


def afiliacao_mais_recente(docs, auid):
    """
    Afiliação do autor no documento mais recente do perfil em que ela é
    identificável. Se mais de um documento tiver a data mais recente, conta
    como Brasil se ao menos um deles indicar afiliação brasileira.

    Retorna (data, instituicao_brasileira_ou_None), ou None se nenhum
    documento permitir identificar a afiliação do autor.
    """
    por_data = defaultdict(list)
    for d in docs:
        data = getattr(d, "coverDate", None)
        if not data:
            continue
        for _, aid, inst in afiliacoes_do_documento(d) or []:
            if aid == auid:
                por_data[data].append(inst)
    if not por_data:
        return None
    data = max(por_data)
    brasileiras = [i for i in por_data[data] if i]
    return data, (brasileiras[0] if brasileiras else None)


def coletar_documentos():
    """
    Roda as buscas por tema e acumula dados por autor brasileiro.

    Retorna (autores, eids_po):
      autores: dict author_id -> {
          "nome": str, "docs": {eid: (citacoes, data, instituicao)},
          "temas": set
      }
      eids_po: conjunto com os EIDs de todos os documentos de PO achados
    """
    autores = defaultdict(lambda: {"nome": "", "docs": {}, "temas": set()})
    eids_po = set()   # todos os documentos de PO encontrados nas buscas
    docs_ignorados = 0

    for tema, cfg in TEMAS.items():
        query = montar_query(cfg["termos"], cfg.get("contexto"))
        print(f"\n[{tema}]\n  Query: {query}")
        try:
            s = ScopusSearch(query, refresh=REFRESH_DIAS, view="COMPLETE")
        except Exception as e:
            print(f"  ERRO na busca deste tema: {e}")
            continue

        resultados = s.results or []
        print(f"  {len(resultados)} documentos encontrados.")

        for d in resultados:
            eids_po.add(d.eid)
            afiliacoes = afiliacoes_do_documento(d)
            if afiliacoes is None:   # campos de afiliação inconsistentes
                docs_ignorados += 1
                continue

            nomes = dividir(d.author_names)
            citacoes = int(d.citedby_count or 0)
            data = d.coverDate or ""

            for i, auid, inst_brasil in afiliacoes:
                if inst_brasil is None:
                    continue  # autor sem afiliação brasileira neste artigo

                reg = autores[auid]
                if not reg["nome"] and i < len(nomes):
                    reg["nome"] = nomes[i]
                reg["docs"][d.eid] = (citacoes, data, inst_brasil)
                reg["temas"].add(tema)

    if docs_ignorados:
        print(
            f"\nAviso: {docs_ignorados} documentos foram ignorados por "
            "campos de afiliação inconsistentes."
        )
    return autores, eids_po


def resumir_autor(auid, reg):
    """Calcula as métricas de PO de um autor a partir dos documentos."""
    citacoes = [c for c, _, _ in reg["docs"].values()]
    # Instituição do artigo de PO mais recente
    _, _, inst_recente = max(reg["docs"].values(), key=lambda x: x[1])
    return {
        "auid": auid,
        "nome_doc": reg["nome"],
        "universidade": inst_recente,
        "n_po": len(reg["docs"]),
        "cit_po": sum(citacoes),  # só para priorizar a ordem das consultas
        "temas": "; ".join(sorted(reg["temas"])),
    }


# -------------------------------------------------------------------------
# Bolsas de produtividade em pesquisa (PQ) do CNPq, lidas da planilha de
# bolsistas (colunas: Nome, Nível, Início, Término, Instituição, Situação,
# Área do Conhecimento)
# -------------------------------------------------------------------------
PARTICULAS_NOME = {"de", "da", "do", "dos", "das", "e", "di", "du", "del", "van", "von"}
PARTICULAS_INSTITUICAO = PARTICULAS_NOME | {"of", "the", "and", "at"}
# Palavras de nome de instituição que a sigla pode "pular" (UNICAMP não
# precisa de "Estadual"); as demais palavras do nome precisam aparecer na sigla.
PALAVRAS_GENERICAS_INSTITUICAO = {
    "universidade", "university", "universidad", "federal", "estadual",
    "state", "instituto", "institute", "faculdade", "centro", "center",
    "centre", "escola", "school", "fundacao", "foundation", "tecnologico",
    "tecnologica", "technology", "technological", "nacional", "national",
}
# Abreviações de estado que não são prefixo da palavra (UFPR, UFPB)
ABREVIACOES_ESPECIAIS = {"parana": ("pr",), "paraiba": ("pb",)}
SUFIXOS_NOME = {"junior", "jr", "filho", "neto", "sobrinho", "segundo"}
AREAS_PLAUSIVEIS_NORM = tuple(normalizar_texto(a) for a in AREAS_CNPQ_PLAUSIVEIS)


def tokens_nome(nome):
    """Palavras do nome sem acentos/pontuação, sem partículas (de, da...) e sem sufixo (Júnior, Filho...)."""
    tokens = [
        t for t in re.findall(r"[a-z0-9]+", normalizar_texto(nome))
        if t not in PARTICULAS_NOME
    ]
    while len(tokens) > 1 and tokens[-1] in SUFIXOS_NOME:
        tokens.pop()
    return tokens


def chave_fonetica(token):
    """
    Forma simplificada de uma palavra para comparar grafias equivalentes
    (Luis/Luiz, Souza/Sousa, Christina/Cristina, Marcello/Marcelo,
    Mathias/Matias, Hélio/Elio): ph->f, th->t, k->c, z->s, y->i, w->v,
    'h' inicial removido e letras repetidas colapsadas.
    """
    t = token.replace("chr", "cr").replace("ph", "f").replace("th", "t").replace("ck", "c")
    t = t.translate(str.maketrans({"k": "c", "z": "s", "y": "i", "w": "v"}))
    t = re.sub(r"(.)\1+", r"\1", t)
    if len(t) > 2 and t.startswith("h"):
        t = t[1:]
    return t


def _similares(a, b):
    """
    Palavras parecidas (erro de digitação) por similaridade de texto. Só vale
    para palavras de 5+ letras e nunca para nomes que diferem apenas na vogal
    final (Rafael/Rafaela, Marcelo/Marcela, Adriano/Adriana: pessoas distintas).
    """
    if min(len(a), len(b)) < 5:
        return False
    if a != b and a.rstrip("aeiou") == b.rstrip("aeiou"):
        return False
    return SequenceMatcher(None, a, b).ratio() >= 0.85


def _tok_compativel(a, b):
    """Iguais, inicial de uma delas, mesma grafia fonética ou muito parecidas."""
    if a == b:
        return True
    if (len(a) == 1 and b.startswith(a)) or (len(b) == 1 and a.startswith(b)):
        return True
    return chave_fonetica(a) == chave_fonetica(b) or _similares(a, b)


def _subsequencia(curto, longo):
    """Todas as palavras de 'curto' aparecem, em ordem, em 'longo' (aceitando iniciais e grafias parecidas)."""
    i = 0
    for t in longo:
        if i < len(curto) and _tok_compativel(curto[i], t):
            i += 1
    return i == len(curto)


def _nome_aproximado(tq, tc):
    """
    Correspondência aproximada entre o nome buscado (tq) e o nome no CNPq (tc).
    Primeiro nome igual, inicial ou de grafia equivalente e mais uma destas
    estruturas:
      (a) mesmo sobrenome na ÚLTIMA posição, com nomes do meio omitidos ou
          abreviados ('Marcelo Nagano' ~ 'Marcelo Seido Nagano');
      (b) o nome buscado é o COMEÇO do nome do CNPq, que só tem sobrenomes a
          mais no final ('Eduardo Uchoa' ~ 'Eduardo Uchoa Barboza').
    O sobrenome buscado no meio do nome do CNPq não vale: 'Marcos dos Santos'
    não é 'Marcos Vinícius dos Santos Rezende'.
    """
    if len(tq) < 2 or len(tc) < 2:
        return False
    if not _tok_compativel(tq[0], tc[0]):
        return False
    if chave_fonetica(tq[-1]) == chave_fonetica(tc[-1]):                  # (a)
        return _subsequencia(tq, tc) or _subsequencia(tc, tq)
    return len(tq) < len(tc) and all(                                     # (b)
        _tok_compativel(a, b) for a, b in zip(tq, tc)
    )


def sigla_compativel(sigla, universidade):
    """
    Aproximação para saber se a sigla do CNPq (UFC, USP, UNICAMP...) pode
    corresponder ao nome da instituição no Scopus. A sigla precisa ser formada
    por PREFIXOS de palavras diferentes do nome (UFC ~ Universidade Federal do
    Ceará; UNICAMP ~ Universidade Estadual de Campinas: "uni" + "camp";
    UTFPR ~ ... do Paraná, com a abreviação "pr"), e a última palavra
    distintiva do nome (não genérica, como universidade/federal/estadual)
    precisa entrar na sigla: UFS é compatível com Universidade Federal de
    Sergipe, mas não com ... de Santa Catarina nem ... de São Carlos; UFC não
    é compatível com ... de Campina Grande. A ordem das palavras não importa,
    para cobrir nomes em inglês ("Federal University of Technology - Paraná").
    Serve só para desempatar homônimos e validar correspondências.
    """
    s = re.sub(r"[^a-z0-9]", "", normalizar_texto(sigla))
    palavras = [
        p for p in re.findall(r"[a-z0-9]+", normalizar_texto(universidade))
        if p not in PARTICULAS_INSTITUICAO
    ]
    if not s or not palavras:
        return False
    if s in palavras:   # a própria sigla aparece no nome (ex.: "... (UFC)")
        return True

    distintivas = [
        j for j, w in enumerate(palavras) if w not in PALAVRAS_GENERICAS_INSTITUICAO
    ]
    ultima = distintivas[-1] if distintivas else None

    @lru_cache(maxsize=None)
    def formavel(i, usadas):
        if i == len(s):
            # a última palavra distintiva do nome precisa ter entrado na sigla
            return ultima is None or bool(usadas >> ultima & 1)
        for j, w in enumerate(palavras):
            if usadas >> j & 1:
                continue
            for k in range(1, min(len(w), len(s) - i) + 1):
                trecho = s[i:i + k]
                if (w.startswith(trecho) or trecho in ABREVIACOES_ESPECIAIS.get(w, ())) \
                        and formavel(i + k, usadas | (1 << j)):
                    return True
        return False

    return formavel(0, 0)


def area_plausivel(area):
    """A área do CNPq combina com o perfil do ranking (AREAS_CNPQ_PLAUSIVEIS)?"""
    a = normalizar_texto(area)
    return any(trecho in a for trecho in AREAS_PLAUSIVEIS_NORM)


def carregar_bolsistas_pq(caminho):
    """
    Lê a planilha de bolsistas PQ e a indexa pela grafia fonética de CADA
    palavra do nome (o sobrenome buscado pode estar em qualquer posição).
    Cada "pessoa" agrupa as linhas de um mesmo nome + instituição (a planilha
    traz uma linha por período de vigência).

    Retorna {"pessoas": [...], "indice": {chave_fonetica: [posições]}}.
    """
    df = pd.read_excel(caminho)
    obrigatorias = ["Nome", "Nível", "Início", "Término", "Instituição", "Situação"]
    faltando = [c for c in obrigatorias if c not in df.columns]
    if faltando:
        raise ValueError(f"A planilha de bolsistas não tem as colunas: {faltando}")

    pessoas = {}
    for linha in df.to_dict("records"):
        tokens = tokens_nome(linha["Nome"])
        if not tokens:
            continue
        inst = str(linha["Instituição"]).strip()
        pessoa = pessoas.setdefault(
            (tuple(tokens), inst),
            {
                "nome": str(linha["Nome"]).strip(),
                "tokens": tokens,
                "chaves": [chave_fonetica(t) for t in tokens],
                "inst": inst,
                "area": str(linha.get("Área do Conhecimento", "") or "").strip(),
                "linhas": [],
            },
        )
        pessoa["linhas"].append(linha)

    lista = list(pessoas.values())
    indice = defaultdict(list)
    for pos, pessoa in enumerate(lista):
        for chave in set(pessoa["chaves"]):
            indice[chave].append(pos)
    print(f"Planilha de bolsistas PQ carregada: {len(df)} linhas, {len(lista)} bolsistas.")
    return {"pessoas": lista, "indice": dict(indice)}


def _data_br(texto):
    try:
        return datetime.strptime(str(texto).strip(), "%d/%m/%Y").date()
    except ValueError:
        return None


def escolher_vigencia(linhas, ref):
    """
    Entre as linhas (períodos) de um mesmo bolsista, escolhe a vigente na
    data de referência; se não houver, a próxima a começar; senão a que
    terminou mais recentemente.
    """
    def inicio(l):
        return _data_br(l["Início"]) or date.min

    def fim(l):
        return _data_br(l["Término"]) or date.max

    vigentes = [l for l in linhas if inicio(l) <= ref <= fim(l)]
    if vigentes:
        return max(vigentes, key=inicio)
    futuras = [l for l in linhas if inicio(l) > ref]
    if futuras:
        return min(futuras, key=inicio)
    return max(linhas, key=fim)


def bolsa_cnpq(nome_pesquisador, universidade, bolsistas, ref=None, diagnostico=None):
    """
    Procura o pesquisador na planilha de bolsistas PQ.

    Retorna:
      - dict (nivel, vigencia_inicio, vigencia_fim, instituicao, situacao,
        area e, se a correspondência não foi idêntica ou precisou de
        ressalva, 'obs') quando encontra;
      - None quando não há bolsa para o nome;
      - str com um aviso quando há homônimos que não foi possível separar.

    O nome é comparado em três níveis, do mais ao menos seguro, e vale o
    primeiro que achar candidatos:
      0. idêntico (sem acentos, maiúsculas, partículas);
      1. mesma grafia fonética (Luis ~ Luiz, Souza ~ Sousa);
      2. aproximado (sobrenome em qualquer posição, iniciais, nome do meio
         omitido, palavras parecidas).
    Havendo homônimos, a sigla da instituição desempata. Para um único
    candidato: níveis 0 e 1 com 3+ palavras são aceitos direto; nos demais
    casos exige-se instituição compatível OU primeiro nome por extenso com
    área plausível (AREAS_CNPQ_PLAUSIVEIS) e sobrenome incomum
    (LIMITE_SOBRENOME_COMUM). Candidatos descartados por esses critérios são
    descritos na lista 'diagnostico', se fornecida.
    """
    ref = ref or date.today()
    tq = tokens_nome(nome_pesquisador)
    if not tq:
        return None

    chaves_q = [chave_fonetica(t) for t in tq]
    base = bolsistas["pessoas"]
    candidatos = [base[i] for i in bolsistas["indice"].get(chaves_q[-1], [])]

    pessoas = [p for p in candidatos if p["tokens"] == tq]
    nivel = 0
    if not pessoas:
        pessoas = [p for p in candidatos if p["chaves"] == chaves_q]
        nivel = 1
    if not pessoas:
        pessoas = [p for p in candidatos if _nome_aproximado(tq, p["tokens"])]
        nivel = 2
    if not pessoas:
        return None

    def inst_ok(p):
        return sigla_compativel(p["inst"], universidade)

    if len(pessoas) > 1:
        compativeis = [p for p in pessoas if inst_ok(p)]
        if len(compativeis) == 1:
            pessoas = compativeis
        elif not compativeis and nivel == 2:
            # vários nomes parecidos e nenhum da mesma instituição: sem
            # evidência para atribuir bolsa a ninguém
            if diagnostico is not None:
                diagnostico.extend(
                    f"'{p['nome']}' ({p['inst']}, {p['area']})" for p in pessoas[:5]
                )
            return None
        else:
            insts = sorted({p["inst"] for p in pessoas})
            lista = ", ".join(insts[:5]) + (", ..." if len(insts) > 5 else "")
            return f"Homônimos no CNPq ({len(pessoas)}: {lista}) - verificar"

    pessoa = pessoas[0]
    ressalva = None
    if not (nivel <= 1 and len(tq) >= 3) and not inst_ok(pessoa):
        prenome_completo = len(tq[0]) > 1
        sobrenome_raro = len(candidatos) <= LIMITE_SOBRENOME_COMUM
        if prenome_completo and sobrenome_raro and area_plausivel(pessoa["area"]):
            ressalva = (
                f"instituição divergente (CNPq: {pessoa['inst']}; Scopus: "
                f"{universidade}), aceito pela área {pessoa['area']}"
            )
        else:
            if diagnostico is not None:
                diagnostico.append(
                    f"'{pessoa['nome']}' ({pessoa['inst']}, {pessoa['area']})"
                )
            return None

    linha = escolher_vigencia(pessoa["linhas"], ref)
    info = {
        "nivel": str(linha["Nível"]).strip(),
        "vigencia_inicio": str(linha["Início"]).strip(),
        "vigencia_fim": str(linha["Término"]).strip(),
        "instituicao": pessoa["inst"] or None,
        "situacao": str(linha["Situação"]).strip(),
        "area": pessoa["area"] or None,
    }
    obs = []
    if nivel == 1:
        obs.append(f"grafia diferente de '{pessoa['nome']}'")
    elif nivel == 2:
        obs.append(f"correspondência aproximada com '{pessoa['nome']}' ({pessoa['inst']})")
    if ressalva:
        obs.append(ressalva)
    if obs:
        info["obs"] = "; ".join(obs)
    return info


def colunas_bolsa(info):
    """
    Converte o retorno de bolsa_cnpq nas colunas da planilha. Quando não há
    bolsa (ou há homônimos), a mensagem vai em 'Bolsa CNPq - Nível' e as
    demais colunas ficam vazias.
    """
    vazio = {"bolsa_nivel": None, "bolsa_inicio": None, "bolsa_fim": None,
             "bolsa_instituicao": None, "bolsa_situacao": None,
             "bolsa_area": None}
    if info is None:
        return {**vazio, "bolsa_nivel": "Sem bolsa em curso"}
    if isinstance(info, str):  # aviso de homônimos
        return {**vazio, "bolsa_nivel": info}
    return {
        "bolsa_nivel": info["nivel"],
        "bolsa_inicio": info["vigencia_inicio"],
        "bolsa_fim": info["vigencia_fim"],
        "bolsa_instituicao": info["instituicao"],
        "bolsa_situacao": info["situacao"],
        "bolsa_area": info["area"],
    }


def nome_para_cnpq(au, fallback):
    """
    Nome na ordem natural ("Bruno de Athayde Prata"), que é como aparece na
    planilha do CNPq. Se o perfil trouxer nome e sobrenome, usa-os; senão
    inverte o nome do perfil ("Prata, Bruno" -> "Bruno Prata").
    """
    sobrenome = getattr(au, "surname", None)
    nome = getattr(au, "given_name", None)
    if sobrenome and nome:
        return f"{nome} {sobrenome}"
    if "," in fallback:
        sob, _, nomes = fallback.partition(",")
        return f"{nomes.strip()} {sob.strip()}"
    return fallback.strip()


def nome_do_perfil(au, fallback):
    """Nome 'Sobrenome, Nome' do perfil; usa o nome do documento se faltar."""
    try:
        sobrenome = getattr(au, "surname", None)
        nome = getattr(au, "given_name", None)
        if sobrenome and nome:
            return f"{sobrenome}, {nome}"
        indexado = getattr(au, "indexed_name", None)
        if indexado:
            return indexado
    except Exception:
        pass
    return fallback


def data_de_corte_po():
    """
    Data mínima (ISO, AAAA-MM-DD) para um documento de PO contar como recente
    no filtro ANOS_JANELA_PO, contada a partir de DATA_REFERENCIA (ou de hoje).
    Retorna None quando o filtro está desligado.
    """
    if not ANOS_JANELA_PO:
        return None
    ref = DATA_REFERENCIA or date.today()
    try:
        corte = ref.replace(year=ref.year - ANOS_JANELA_PO)
    except ValueError:   # 29/02 em ano não bissexto
        corte = ref.replace(year=ref.year - ANOS_JANELA_PO, day=28)
    return corte.isoformat()


def analisar_documentos(au, eids_po, auid, data_corte=None):
    """
    Lista todos os documentos do perfil (AuthorRetrieval.get_documents, uma
    busca adicional - por isso só é chamada para quem já passou no filtro
    do fator h) e calcula: citação máxima, total de documentos, quantos são
    de PO, os temas de PO encontrados, a afiliação do autor no documento
    mais recente (afiliacao_mais_recente) e a atividade recente em PO
    (documentos de PO com data >= data_corte e data do último documento de PO)
    e quantos documentos do perfil caem nos TEMAS_EXCLUIDOS (e em quais).
    """
    docs = au.get_documents(refresh=REFRESH_DIAS) or []
    cit_max = 0
    n_po = 0
    n_po_janela = 0
    ultimo_po = ""
    temas = set()
    n_excluidos = 0
    por_tema_excluido = defaultdict(int)
    for d in docs:
        cit_max = max(cit_max, int(d.citedby_count or 0))
        excluidos_doc = temas_excluidos_do_documento(d)
        if excluidos_doc:
            n_excluidos += 1
            for t in excluidos_doc:
                por_tema_excluido[t] += 1
        temas_doc = temas_do_documento(d)
        if d.eid in eids_po or temas_doc:
            n_po += 1
            temas |= temas_doc
            data_doc = getattr(d, "coverDate", None) or ""
            if data_doc:
                ultimo_po = max(ultimo_po, data_doc)
                if data_corte and data_doc >= data_corte:
                    n_po_janela += 1
    return {
        "cit_max": cit_max,
        "n_docs_perfil": len(docs),
        "n_po_perfil": n_po,
        "temas_locais": temas,
        "afiliacao_recente": afiliacao_mais_recente(docs, auid),
        "n_po_janela": n_po_janela,
        "ultimo_po": ultimo_po or None,
        "n_excluidos_perfil": n_excluidos,
        "temas_excluidos": dict(por_tema_excluido),
    }


def consultar_perfis(candidatos, eids_po, bolsistas):
    """
    Consulta o perfil de cada candidato. Para quem atinge H_MINIMO, coleta
    também total de documentos, total de citações e citação máxima do
    perfil inteiro, e aplica os filtros de FRACAO_MINIMA_PO, de temas
    excluídos (FRACAO_MAXIMA_TEMAS_EXCLUIDOS), de atividade recente em PO e
    de afiliação (documento mais recente do perfil com afiliação no Brasil).

    Retorna (aprovados, excluidos), com excluidos = dict com as contagens
    'fracao', 'tema_excluido', 'sem_po_recente', 'fora_brasil' e
    'sem_afiliacao'.
    """
    aprovados = []
    excluidos = {
        "fracao": 0, "tema_excluido": 0, "sem_po_recente": 0,
        "fora_brasil": 0, "sem_afiliacao": 0,
    }
    data_corte = data_de_corte_po()
    total = len(candidatos)

    for i, c in enumerate(candidatos, start=1):
        try:
            au = AuthorRetrieval(c["auid"], refresh=REFRESH_DIAS)
            h = au.h_index
            h = int(h) if h not in (None, "") else 0

            if h < H_MINIMO:
                if i % 25 == 0 or i == total:
                    print(f"  [{i}/{total}] perfis consultados | {len(aprovados)} com h >= {H_MINIMO}")
                time.sleep(PAUSA_PERFIL)
                continue

            c = dict(c)
            c["h"] = h
            c["nome"] = nome_do_perfil(au, c["nome_doc"])
            c["n_artigos"] = int(au.document_count or 0)
            c["cit_totais"] = int(au.citation_count or 0)

            try:
                analise = analisar_documentos(au, eids_po, c["auid"], data_corte)
            except Exception as e_docs:
                if e_docs.__class__.__name__ == "Scopus429Error":
                    raise
                print(
                    f"  {c['nome']}: não foi possível listar os documentos "
                    f"do perfil ({e_docs}); autor descartado por não ser "
                    "possível verificar a fração em PO."
                )
                continue

            n_docs = analise["n_docs_perfil"]
            fracao = analise["n_po_perfil"] / n_docs if n_docs else 0.0
            if fracao < FRACAO_MINIMA_PO - 1e-9:
                excluidos["fracao"] += 1
                continue

            # Remove autores com muitos documentos nos TEMAS_EXCLUIDOS
            if FRACAO_MAXIMA_TEMAS_EXCLUIDOS is not None and TEMAS_EXCLUIDOS:
                fracao_excl = analise["n_excluidos_perfil"] / n_docs if n_docs else 0.0
                if fracao_excl >= FRACAO_MAXIMA_TEMAS_EXCLUIDOS - 1e-9:
                    excluidos["tema_excluido"] += 1
                    detalhe = ", ".join(
                        f"{TEMAS_EXCLUIDOS[t].get('sigla', t)}: {n}"
                        for t, n in sorted(
                            analise["temas_excluidos"].items(), key=lambda x: -x[1]
                        )
                    )
                    print(
                        f"  {c['nome']}: {fracao_excl:.0%} dos documentos do "
                        f"perfil em temas excluídos ({detalhe}); autor excluído."
                    )
                    continue

            # Só autores com documento de PO nos últimos ANOS_JANELA_PO anos
            if data_corte and analise["n_po_janela"] == 0:
                excluidos["sem_po_recente"] += 1
                ultimo = analise["ultimo_po"]
                if ultimo:
                    detalhe = f"último documento de PO em {ultimo}, anterior a {data_corte}"
                else:
                    detalhe = (
                        "documentos de PO sem data de publicação, que não "
                        f"contam como recentes (corte: {data_corte})"
                    )
                print(f"  {c['nome']}: {detalhe}; autor excluído.")
                continue

            # Só autores cujo documento mais recente tem afiliação no Brasil
            recente = analise["afiliacao_recente"]
            if recente is None:
                excluidos["sem_afiliacao"] += 1
                print(
                    f"  {c['nome']}: afiliação do documento mais recente "
                    "não identificável; autor descartado."
                )
                continue
            data_recente, inst_recente = recente
            if inst_recente is None:
                excluidos["fora_brasil"] += 1
                print(
                    f"  {c['nome']}: documento mais recente ({data_recente}) "
                    "sem afiliação no Brasil; autor excluído."
                )
                continue
            c["universidade"] = inst_recente   # instituição atual (documento mais recente)

            descartados = []
            info_bolsa = bolsa_cnpq(
                nome_para_cnpq(au, c["nome"]), c["universidade"],
                bolsistas, DATA_REFERENCIA, descartados,
            )
            if isinstance(info_bolsa, dict) and info_bolsa.get("obs"):
                print(f"  {c['nome']}: bolsa CNPq - {info_bolsa['obs']}")
            elif isinstance(info_bolsa, str):
                print(f"  {c['nome']}: bolsa CNPq - {info_bolsa}")
            elif descartados:
                print(
                    f"  {c['nome']}: sem bolsa; nome parecido no CNPq "
                    f"descartado (instituição/área não conferem): "
                    + "; ".join(descartados)
                )
            c.update(colunas_bolsa(info_bolsa))

            c["cit_max"] = analise["cit_max"]
            c["n_po"] = analise["n_po_perfil"]
            c["fracao_po"] = fracao
            c["temas"] = "; ".join(
                sorted(set(c["temas"].split("; ")) - {""} | analise["temas_locais"])
            )
            aprovados.append(c)
        except Exception as e:
            if e.__class__.__name__ == "Scopus429Error":
                print(
                    "\n*** Cota da API Scopus esgotada. Consultados "
                    f"{i - 1}/{total} candidatos; o ranking será gerado com "
                    "o que já foi coletado. Rode novamente depois que a "
                    "cota renovar (o cache evita repetir as consultas)."
                )
                break
            print(f"  [{i}/{total}] {c['nome_doc']}: erro no perfil ({e})")
            continue

        if i % 25 == 0 or i == total:
            print(f"  [{i}/{total}] perfis consultados | {len(aprovados)} com h >= {H_MINIMO}")
        time.sleep(PAUSA_PERFIL)

    return aprovados, excluidos


# -------------------------------------------------------------------------
# Formatação da planilha
# -------------------------------------------------------------------------
COR_GRUPOS = {
    "identificacao": "1F3864",   # azul-escuro
    "scopus": "2E75B6",          # azul
    "po": "548235",              # verde
    "bolsa": "7030A0",           # roxo
    "link": "595959",            # cinza
}
GRUPO_DA_COLUNA = {
    "Ranking": "identificacao", "Autor": "identificacao", "Universidade": "identificacao",
    "Fator h": "scopus", "Nº de artigos": "scopus", "Citações totais": "scopus",
    "Citações máx.": "scopus",
    "Artigos de PO": "po", "% em PO": "po", "Temas de PO": "po",
    "Link Scopus": "link",
}
COLUNAS_A_ESQUERDA = {
    "Autor", "Universidade", "Bolsa CNPq - Área do Conhecimento", "Temas de PO",
}
COLUNAS_COM_QUEBRA = {
    "Universidade", "Bolsa CNPq - Nível", "Bolsa CNPq - Área do Conhecimento",
    "Temas de PO",
}
LARGURA_MAXIMA = {
    "Autor": 34, "Universidade": 36, "Bolsa CNPq - Nível": 24,
    "Bolsa CNPq - Área do Conhecimento": 30, "Temas de PO": 48,
}
LARGURA_MINIMA = {"Fator h": 12}
FORMATO_NUMERICO = {
    "Ranking": "0", "Fator h": "0", "Nº de artigos": "#,##0",
    "Citações totais": "#,##0", "Citações máx.": "#,##0", "Artigos de PO": "0",
    "% em PO": "0.0%", "Bolsa CNPq - Início": "dd/mm/yyyy",
    "Bolsa CNPq - Término": "dd/mm/yyyy",
}


def _texto_exibido(valor, nome_coluna):
    """Texto aproximado de como o Excel exibirá a célula (para dimensionar a coluna)."""
    if nome_coluna == "% em PO" and isinstance(valor, (int, float)):
        return f"{valor:.1%}"
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        return f"{valor:,.0f}"
    return str(valor)


def formatar_ranking(ws, tabela):
    """
    Deixa a aba Ranking legível numa tela: cabeçalho colorido por grupo de
    colunas, linhas alternadas, larguras ajustadas ao conteúdo, formatos de
    número/data, painéis congelados (Ranking e Autor), filtros e realces
    condicionais (fator h, % em PO, nível e situação da bolsa).
    """
    n_linhas, n_colunas = ws.max_row, ws.max_column
    idx = {nome: i + 1 for i, nome in enumerate(tabela.columns)}
    letra = {nome: get_column_letter(i) for nome, i in idx.items()}

    borda = Border(bottom=Side(style="thin", color="D9D9D9"))
    zebra = PatternFill("solid", start_color="F3F6FA", end_color="F3F6FA")

    # --- cabeçalho ---
    for nome, i in idx.items():
        grupo = "bolsa" if nome.startswith("Bolsa CNPq") else GRUPO_DA_COLUNA.get(nome, "identificacao")
        cor = COR_GRUPOS[grupo]
        cel = ws.cell(row=1, column=i)
        cel.fill = PatternFill("solid", start_color=cor, end_color=cor)
        cel.font = Font(bold=True, color="FFFFFF")
        cel.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 45

    # --- corpo ---
    for linha in range(2, n_linhas + 1):
        for nome, i in idx.items():
            cel = ws.cell(row=linha, column=i)
            cel.border = borda
            if linha % 2 == 1:
                cel.fill = zebra
            cel.alignment = Alignment(
                horizontal="left" if nome in COLUNAS_A_ESQUERDA else "center",
                vertical="center",
                wrap_text=nome in COLUNAS_COM_QUEBRA,
            )
            if nome in FORMATO_NUMERICO:
                cel.number_format = FORMATO_NUMERICO[nome]
            if nome == "Autor":
                cel.font = Font(bold=True)
            if nome == "Link Scopus" and cel.value:
                cel.hyperlink = cel.value
                cel.value = "Abrir"
                cel.font = Font(color="0563C1", underline="single")

    # --- larguras: ajustadas ao conteúdo, com mínimo/máximo por coluna ---
    for nome, i in idx.items():
        if nome == "Link Scopus":
            largura = 10
        elif nome.endswith(("Início", "Término")):
            largura = 12
        else:
            conteudo = max(
                (len(_texto_exibido(ws.cell(row=r, column=i).value, nome))
                 for r in range(2, n_linhas + 1)
                 if ws.cell(row=r, column=i).value is not None),
                default=0,
            )
            minimo = max(
                max(len(palavra) for palavra in nome.split()) + 3,
                LARGURA_MINIMA.get(nome, 0),
            )
            largura = max(minimo, min(LARGURA_MAXIMA.get(nome, 28), conteudo + 2))
        ws.column_dimensions[get_column_letter(i)].width = largura

    # --- realces condicionais ---
    ultima = n_linhas
    faixa = lambda nome: f"{letra[nome]}2:{letra[nome]}{ultima}"
    ws.conditional_formatting.add(
        faixa("Fator h"),
        DataBarRule(start_type="num", start_value=0, end_type="max",
                    color="9DC3E6", showValue=True),
    )
    ws.conditional_formatting.add(
        faixa("% em PO"),
        ColorScaleRule(start_type="min", start_color="FFFFFF",
                       end_type="max", end_color="A9D18E"),
    )

    nivel = letra["Bolsa CNPq - Nível"]
    primeira = f"{nivel}2"
    regras_nivel = [
        (f'OR({primeira}="PQ-SR",{primeira}="PQ-1A",{primeira}="PQ-A")',
         PatternFill("solid", start_color="A9D18E", end_color="A9D18E"),
         Font(bold=True, color="375623")),
        (f'OR({primeira}="PQ-1B",{primeira}="PQ-1C",{primeira}="PQ-1D",{primeira}="PQ-B")',
         PatternFill("solid", start_color="E2F0D9", end_color="E2F0D9"),
         Font(color="375623")),
        (f'OR({primeira}="PQ-2",{primeira}="PQ-C")',
         PatternFill("solid", start_color="FFF2CC", end_color="FFF2CC"),
         Font(color="7F6000")),
        (f'{primeira}="Sem bolsa em curso"', None, Font(italic=True, color="7F7F7F")),
        (f'LEFT({primeira},9)="Homônimos"',
         PatternFill("solid", start_color="F8CBAD", end_color="F8CBAD"),
         Font(bold=True, color="833C0B")),
    ]
    for formula, preenchimento, fonte in regras_nivel:
        ws.conditional_formatting.add(
            faixa("Bolsa CNPq - Nível"),
            FormulaRule(formula=[formula], fill=preenchimento, font=fonte),
        )
    sit = letra["Bolsa CNPq - Situação"]
    ws.conditional_formatting.add(
        faixa("Bolsa CNPq - Situação"),
        FormulaRule(
            formula=[f'{sit}2="Suspenso"'],
            fill=PatternFill("solid", start_color="FFC7CE", end_color="FFC7CE"),
            font=Font(bold=True, color="9C0006"),
        ),
    )

    # --- visualização e impressão ---
    ws.freeze_panes = "C2"                      # Ranking e Autor sempre visíveis
    ws.auto_filter.ref = f"A1:{get_column_letter(n_colunas)}{n_linhas}"
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    ws.sheet_properties.tabColor = COR_GRUPOS["identificacao"]
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.print_title_rows = "1:1"


def formatar_metodologia(ws):
    """Aba Metodologia: cabeçalho destacado, texto com quebra de linha e itens em negrito."""
    azul = COR_GRUPOS["identificacao"]
    for cel in ws[1]:
        cel.fill = PatternFill("solid", start_color=azul, end_color=azul)
        cel.font = Font(bold=True, color="FFFFFF")
        cel.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 24
    ws.column_dimensions["A"].width = 44
    ws.column_dimensions["B"].width = 120
    borda = Border(bottom=Side(style="thin", color="D9D9D9"))
    for linha in range(2, ws.max_row + 1):
        item, descricao = ws.cell(row=linha, column=1), ws.cell(row=linha, column=2)
        item.font = Font(bold=True)
        item.alignment = Alignment(vertical="top", wrap_text=True)
        descricao.alignment = Alignment(vertical="top", wrap_text=True)
        item.border = descricao.border = borda
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "7F7F7F"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)


def siglas_dos_temas(temas):
    """'Tema A; Tema B' -> 'SiglaA, SiglaB', na ordem em que os temas aparecem em TEMAS."""
    marcados = {t for t in str(temas).split("; ") if t}
    return ", ".join(cfg.get("sigla", t) for t, cfg in TEMAS.items() if t in marcados)


def descricao_tema_excluido(cfg):
    """Texto da aba Metodologia para um tema excluído (termos e contexto)."""
    termos = ", ".join(t.strip('"') for t in cfg["termos"])
    texto = f"Termos: {termos}."
    if cfg.get("contexto"):
        contexto = ", ".join(t.strip('"') for t in cfg["contexto"])
        texto += f" Só conta junto de pelo menos um destes: {contexto}."
    return texto


def salvar_excel(aprovados):
    df = pd.DataFrame(aprovados)
    if df.empty:
        print("Nenhum pesquisador atendeu aos critérios.")
        return

    df = df.sort_values(
        ["h", "cit_totais", "n_artigos"], ascending=False
    ).reset_index(drop=True)
    df.insert(0, "Ranking", df.index + 1)
    df["temas"] = df["temas"].apply(siglas_dos_temas)
    df["Link Scopus"] = df["auid"].apply(
        lambda a: f"https://www.scopus.com/authid/detail.uri?authorId={a}"
    )

    tabela = df[
        ["Ranking", "nome", "universidade", "h", "n_artigos",
         "cit_totais", "cit_max", "n_po", "fracao_po",
         "bolsa_nivel", "bolsa_inicio", "bolsa_fim", "bolsa_instituicao",
         "bolsa_situacao", "bolsa_area", "temas", "Link Scopus"]
    ].rename(columns={
        "nome": "Autor", "universidade": "Universidade", "h": "Fator h",
        "n_artigos": "Nº de artigos", "cit_totais": "Citações totais",
        "cit_max": "Citações máx.", "n_po": "Artigos de PO",
        "fracao_po": "% em PO",
        "bolsa_nivel": "Bolsa CNPq - Nível",
        "bolsa_inicio": "Bolsa CNPq - Início",
        "bolsa_fim": "Bolsa CNPq - Término",
        "bolsa_instituicao": "Bolsa CNPq - Instituição",
        "bolsa_situacao": "Bolsa CNPq - Situação",
        "bolsa_area": "Bolsa CNPq - Área do Conhecimento",
        "temas": "Temas de PO",
    })
    for col in ("Bolsa CNPq - Início", "Bolsa CNPq - Término"):
        tabela[col] = tabela[col].apply(
            lambda t: (_data_br(t) or t) if isinstance(t, str) else None
        )

    metodologia = pd.DataFrame({"Item": [
        "Critério de inclusão",
        "Fator h",
        "Nº de artigos / Citações totais / Citações máx.",
        "Artigos de PO / % em PO",
        "Filtro de PO",
        "Filtro de temas excluídos",
        "Filtro de atividade recente em PO",
        "Filtro de afiliação atual",
        "Bolsa CNPq (nível, início, término, instituição, situação, área)",
        "Universidade",
        "Ordenação",
        "Tipos de documento",
        "Período",
        "Cores do cabeçalho",
        "Cores do nível da bolsa",
        "Siglas dos temas (coluna Temas de PO)",
    ] + [f"Busca - {t}" for t in TEMAS]
      + [f"Tema excluído - {t}" for t in TEMAS_EXCLUIDOS], "Descrição": [
        f"Autores com afiliação brasileira em ao menos {MIN_ARTIGOS_PO} "
        f"artigos de PO nas buscas, fator h >= {H_MINIMO} e fração de PO "
        f">= {FRACAO_MINIMA_PO:.0%}.",
        "h-index do perfil Scopus (carreira completa), via AuthorRetrieval.",
        "Do perfil Scopus inteiro (todos os documentos do autor, não só os "
        "de PO). Citações máx. = maior nº de citações entre os documentos "
        "do perfil.",
        "Documentos do perfil classificados como PO (EID encontrado nas "
        "buscas por tema, ou termo de PO no título/resumo/palavras-chave). "
        "% em PO = Artigos de PO / total de documentos do perfil.",
        f"Só entram autores com pelo menos {FRACAO_MINIMA_PO:.0%} dos "
        "documentos do perfil em PO.",
        (
            f"Autores com {FRACAO_MAXIMA_TEMAS_EXCLUIDOS:.0%} ou mais dos "
            "documentos do perfil nos temas excluídos (linhas 'Tema excluído' "
            "abaixo) são removidos. Classificação local pelo título, resumo e "
            "palavras-chave, sem buscas extras no Scopus."
            if FRACAO_MAXIMA_TEMAS_EXCLUIDOS is not None and TEMAS_EXCLUIDOS
            else "Desligado (FRACAO_MAXIMA_TEMAS_EXCLUIDOS = None)."
        ),
        (
            f"Só entram autores com ao menos um documento de PO publicado "
            f"nos últimos {ANOS_JANELA_PO} anos (data de capa do documento "
            f">= {data_de_corte_po()})."
            if ANOS_JANELA_PO else "Desligado (ANOS_JANELA_PO = None)."
        ),
        "Só entram autores cuja afiliação no documento mais recente do "
        "perfil (data de capa mais recente em que a afiliação é "
        "identificável) é no Brasil. Quem não está mais no país é excluído.",
        "Bolsa de produtividade em pesquisa (PQ) do CNPq, buscada pelo nome "
        f"na planilha '{ARQUIVO_BOLSAS_PQ}'. Havendo mais de um período, vale "
        "o vigente na data de referência. O nome é comparado em três níveis: "
        "idêntico (sem acentos), grafia equivalente (Luis/Luiz) e aproximado "
        "(sobrenome em qualquer posição, nome do meio omitido ou abreviado). "
        "Nome pouco específico exige sigla de instituição compatível com a "
        "universidade do Scopus ou primeiro nome por extenso com área "
        "plausível. Homônimos não resolvidos aparecem sinalizados. Situação "
        "'Suspenso' = bolsa suspensa.",
        "Afiliação brasileira do autor no documento mais recente do perfil "
        "(a mesma usada no filtro de afiliação atual).",
        "Fator h, depois citações totais, depois nº de artigos (decrescente), "
        "todos do perfil.",
        ", ".join(TIPOS_DOCUMENTO) + " (ar = artigo, re = revisão, cp = congresso).",
        f"{ANO_INICIO or 'sem limite'} a {ANO_FIM or 'sem limite'}.",
        "Azul-escuro = identificação; azul = métricas do perfil Scopus; "
        "verde = pesquisa operacional; roxo = bolsa CNPq; cinza = link.",
        "Verde-escuro = PQ-SR, PQ-1A, PQ-A; verde-claro = PQ-1B, PQ-1C, "
        "PQ-1D, PQ-B; amarelo = PQ-2, PQ-C; cinza itálico = sem bolsa; "
        "laranja = aviso (homônimos); vermelho = bolsa suspensa.",
        "\n".join(f"{cfg.get('sigla', t)} = {t}" for t, cfg in TEMAS.items()),
    ] + [
        montar_query(cfg["termos"], cfg.get("contexto"))
        for cfg in TEMAS.values()
    ] + [
        descricao_tema_excluido(cfg) for cfg in TEMAS_EXCLUIDOS.values()
    ]})

    with pd.ExcelWriter(ARQUIVO_SAIDA, engine="openpyxl") as writer:
        tabela.to_excel(writer, sheet_name="Ranking", index=False)
        metodologia.to_excel(writer, sheet_name="Metodologia", index=False)
        formatar_ranking(writer.sheets["Ranking"], tabela)
        formatar_metodologia(writer.sheets["Metodologia"])

    print(f"\nRanking salvo em '{ARQUIVO_SAIDA}' com {len(tabela)} pesquisadores.")


def main():
    # Carrega a planilha de bolsistas antes de gastar cota da API
    try:
        bolsistas = carregar_bolsistas_pq(ARQUIVO_BOLSAS_PQ)
    except FileNotFoundError:
        raise SystemExit(
            f"Planilha de bolsistas PQ não encontrada: '{ARQUIVO_BOLSAS_PQ}'. "
            "Coloque o arquivo na pasta do projeto ou ajuste ARQUIVO_BOLSAS_PQ."
        )

    autores, eids_po = coletar_documentos()
    print(f"\nAutores com afiliação brasileira nos artigos de PO: {len(autores)}")

    resumos = [resumir_autor(a, r) for a, r in autores.items()]
    candidatos = [r for r in resumos if r["n_po"] >= MIN_ARTIGOS_PO]
    # Mais produtivos em PO primeiro: se a cota acabar, os mais relevantes
    # já foram consultados.
    candidatos.sort(key=lambda r: (r["n_po"], r["cit_po"]), reverse=True)
    print(
        f"Candidatos com >= {MIN_ARTIGOS_PO} artigos de PO "
        f"(perfis a consultar): {len(candidatos)}\n"
    )

    aprovados, excluidos = consultar_perfis(candidatos, eids_po, bolsistas)
    sem_po_recente = (
        f"{excluidos['sem_po_recente']} sem documento de PO nos últimos "
        f"{ANOS_JANELA_PO} anos; " if ANOS_JANELA_PO else ""
    )
    tema_excl = (
        f"{excluidos['tema_excluido']} com {FRACAO_MAXIMA_TEMAS_EXCLUIDOS:.0%} "
        "ou mais dos documentos em temas excluídos; "
        if FRACAO_MAXIMA_TEMAS_EXCLUIDOS is not None and TEMAS_EXCLUIDOS else ""
    )
    print(
        f"\nExcluídos entre os pesquisadores com h >= {H_MINIMO}: "
        f"{excluidos['fracao']} com menos de {FRACAO_MINIMA_PO:.0%} dos "
        f"documentos em PO; {tema_excl}{sem_po_recente}"
        f"{excluidos['fora_brasil']} cujo documento mais "
        f"recente não tem afiliação no Brasil; {excluidos['sem_afiliacao']} "
        "sem afiliação identificável no perfil."
    )
    salvar_excel(aprovados)


if __name__ == "__main__":
    main()