# Ranking de Pesquisadores em Pesquisa Operacional no Brasil

Conjunto de scripts em Python para coletar informações de bolsas de produtividade em pesquisa do **CNPq** e identificar pesquisadores vinculados a instituições brasileiras que publicam em temas de **Pesquisa Operacional (PO)**. O ranking combina dados bibliográficos e perfis de autor da **Scopus** com informações de bolsas PQ do CNPq.

## Fluxo de execução

O projeto possui duas etapas:

1. **Coleta das bolsas PQ:** `lista_bolsistas_cnpq.py` consulta as páginas de resultados do CNPq para os códigos de áreas do conhecimento definidos no script, consolida os registros e gera `bolsistas_pq_cnpq_todas_areas.xlsx`.
2. **Identificação e ranking:** `ranking_pesquisadores_po.py` realiza buscas temáticas na Scopus, identifica autores com afiliação brasileira, consulta os perfis dos candidatos, aplica os critérios de elegibilidade e gera `ranking_pesquisadores_po_brasil_completo.xlsx`.

## Critérios de seleção

As buscas na Scopus abrangem **18 temas de PO**: roteamento de veículos (VRP), sequenciamento da produção, bin packing, cutting stock, dimensionamento de lotes, localização de instalações, PO geral, decisão multicritério, análise por envoltória de dados (DEA), programação linear/inteira, programação por restrições, métodos de decomposição, teoria dos grafos, metaheurísticas, otimização combinatória e simulação de eventos discretos aplicada à indústria, à cadeia de suprimentos e à saúde.

Com os parâmetros padrão do script, são considerados pesquisadores que atendam a todos os critérios:

- Pelo menos **2 artigos de PO** encontrados nas buscas iniciais;
- **Fator h ≥ 5** no perfil Scopus;
- Pelo menos **1/3 dos documentos** do perfil classificados como trabalhos de PO;
- Menos de **30% dos documentos** classificados nos temas excluídos: gestão de energia, fluxo de potência/sistemas elétricos e sistemas de distribuição elétrica;
- Pelo menos **1 documento de PO nos últimos 7 anos**;
- Afiliação brasileira identificável no documento mais recente do perfil.

A classificação de um documento como PO considera os identificadores recuperados nas buscas iniciais e a ocorrência de termos temáticos no título, resumo ou palavras-chave. Para temas que exigem contexto de aplicação, o termo temático e o termo de contexto devem aparecer no mesmo documento.

A associação com bolsas PQ é feita após a aprovação nos filtros. O script compara nomes de forma exata, fonética ou aproximada e utiliza a instituição e, em situações específicas, a área do conhecimento para auxiliar na identificação de homônimos. **Ter bolsa PQ não é requisito para entrar no ranking.**

## Arquivos

| Arquivo | Descrição |
|---|---|
| `lista_bolsistas_cnpq.py` | Consulta as páginas de resultados do CNPq e consolida os registros de bolsas PQ. |
| `ranking_pesquisadores_po.py` | Realiza as buscas na Scopus, analisa os perfis, aplica os filtros e gera o ranking. |
| `bolsistas_pq_cnpq_todas_areas.xlsx` | Base de entrada do ranking, gerada pelo script de coleta do CNPq. |
| `ranking_pesquisadores_po_brasil_completo.xlsx` | Resultado da seleção, com as abas `Ranking` e `Metodologia`. |

## Dicionário de dados

### 1. Planilha `ranking_pesquisadores_po_brasil_completo.xlsx`

A aba **Ranking** apresenta as seguintes colunas:

| Coluna | Descrição |
|---|---|
| `Ranking` | Posição do pesquisador após a ordenação final. |
| `Autor` | Nome do pesquisador conforme o perfil de autor da Scopus. |
| `Universidade` | Instituição brasileira identificada na afiliação do documento mais recente do perfil. |
| `Fator h` | Índice h do perfil Scopus, considerando a carreira completa. |
| `Nº de artigos` | Número total de documentos registrados no perfil Scopus, incluindo documentos que não são de PO. |
| `Citações totais` | Número total de citações recebido pelos documentos do perfil Scopus. |
| `Citações máx.` | Maior número de citações entre os documentos do perfil. |
| `Artigos de PO` | Quantidade de documentos do perfil classificados como relacionados à Pesquisa Operacional. |
| `% em PO` | Proporção de documentos de PO em relação ao total de documentos do perfil (`Artigos de PO / Nº de artigos`). |
| `Bolsa CNPq - Nível` | Nível da bolsa PQ associada ao pesquisador. Quando não há associação, informa ausência de bolsa em curso ou sinaliza homônimos não resolvidos. |
| `Bolsa CNPq - Início` | Data de início do período de vigência selecionado para a bolsa. |
| `Bolsa CNPq - Término` | Data de término do período de vigência selecionado para a bolsa. |
| `Bolsa CNPq - Instituição` | Instituição vinculada ao registro da bolsa na base do CNPq. |
| `Bolsa CNPq - Situação` | Situação da bolsa conforme o registro do CNPq, por exemplo, em folha de pagamento ou suspensa. |
| `Bolsa CNPq - Área do Conhecimento` | Área do conhecimento associada ao registro da bolsa no CNPq. |
| `Temas de PO` | Temas de Pesquisa Operacional identificados nos documentos do pesquisador, apresentados por siglas. A aba `Metodologia` relaciona as siglas aos temas completos. |
| `Link Scopus` | Link para o perfil do pesquisador na Scopus. |

A aba **Metodologia** documenta os critérios de inclusão, os filtros, os parâmetros utilizados, as expressões de busca por tema e a descrição dos temas excluídos.

### 2. Planilha `bolsistas_pq_cnpq_todas_areas.xlsx`

O script `lista_bolsistas_cnpq.py` consulta os **76 códigos de área** mapeados no dicionário `AREAS_CONHECIMENTO`. Para cada página, lê a sexta tabela HTML retornada pelo CNPq, preserva seus campos originais e acrescenta três colunas de identificação da coleta.

O script de ranking utiliza os seguintes campos da base de bolsas:

| Coluna | Descrição |
|---|---|
| `Nome` | Nome do pesquisador conforme publicado no registro do CNPq. |
| `Nível` | Nível da bolsa de produtividade em pesquisa. |
| `Início` | Data de início do período de vigência da bolsa. |
| `Término` | Data de término do período de vigência da bolsa. |
| `Instituição` | Instituição associada ao registro da bolsa. |
| `Situação` | Situação da bolsa no registro consultado. |
| `Código Área` | Código da área do conhecimento utilizado na consulta ao CNPq. |
| `Área do Conhecimento` | Nome da área associado ao código consultado, conforme o dicionário `AREAS_CONHECIMENTO` do script. |
| `URL` | Endereço da página de resultados do CNPq utilizada para coletar o registro. |

**Nota sobre os campos de origem:** `lista_bolsistas_cnpq.py` não define nem renomeia os cabeçalhos da tabela HTML original: ele preserva as colunas retornadas pelo CNPq. Os seis campos listados acima são os campos que o script de ranking espera para localizar e caracterizar as bolsas. As três colunas de coleta (`Código Área`, `Área do Conhecimento` e `URL`) são acrescentadas pelo próprio script. Como o arquivo Excel gerado por `lista_bolsistas_cnpq.py` não foi anexado nesta conversa, eventuais colunas adicionais da tabela original não puderam ser conferidas individualmente; a estrutura pode variar conforme a página do CNPq.

## Requisitos e instalação

- Python 3;
- Acesso à internet para consultar o CNPq e a API Scopus;
- Chave de API Scopus configurada para uso com `pybliometrics`.

Instale as dependências:

```bash
pip install requests pandas openpyxl pybliometrics
```

## Como executar

Execute os scripts na ordem indicada, no mesmo diretório:

```bash
# 1. Coletar os registros de bolsas PQ do CNPq
python lista_bolsistas_cnpq.py

# 2. Gerar o ranking de pesquisadores em Pesquisa Operacional
python ranking_pesquisadores_po.py
```

O primeiro script gera `bolsistas_pq_cnpq_todas_areas.xlsx`, utilizado como entrada pelo segundo. Formate as colunas da planilha de bolsistas e esclua as linhas com dados em branco. O ranking gera `ranking_pesquisadores_po_brasil_completo.xlsx`.

Antes da execução, configure o acesso à Scopus conforme a documentação do `pybliometrics` e confira os parâmetros no início de `ranking_pesquisadores_po.py`, especialmente `H_MINIMO`, `MIN_ARTIGOS_PO`, `FRACAO_MINIMA_PO`, `FRACAO_MAXIMA_TEMAS_EXCLUIDOS`, `ANOS_JANELA_PO`, `ANO_INICIO`, `ANO_FIM` e `ARQUIVO_BOLSAS_PQ`.

## Ordenação do ranking

Os pesquisadores aprovados são ordenados em ordem decrescente por:

1. Fator h;
2. Citações totais;
3. Número total de documentos no perfil Scopus.

## Observações

Os resultados dependem da cobertura e atualização da Scopus, da qualidade dos perfis de autor, dos termos de busca e dos registros disponibilizados pelo CNPq na data da coleta. A lista do CNPq é construída a partir dos códigos de áreas definidos no dicionário do script. Correspondências aproximadas e possíveis homônimos devem ser conferidos antes do uso dos dados em análises ou decisões. Para permitir a reprodução dos resultados, registre a data da coleta e os parâmetros utilizados.
