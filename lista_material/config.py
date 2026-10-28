"""Configurações do padrão Cadari usadas pelo gerador.

Tudo que é regra de negócio editável fica aqui: mapeamento dos grupos da
tabela do Revit para as seções da lista, textos fixos, palavras usadas para
separar áreas comuns e privativas e o visual da tabela.
"""

# ---------------------------------------------------------------------------
# Seções da lista (ordem do padrão Cadari)
# Cada seção: (grupo do Revit, título na lista, [(subgrupo do Revit, título da subseção)])
# Subgrupo vazio ('') significa itens direto na seção, sem subseção.
# Subgrupos com o mesmo título são somados na mesma subseção.
# ---------------------------------------------------------------------------
SECOES_PADRAO = [
    ("EQUIPAMENTOS", "EQUIPAMENTOS", [
        ("AR-CONDICIONADO", "AR-CONDICIONADO"),
        ("VENTILADORES", "VENTILADORES"),
    ]),
    ("REDE DE DUTOS", "REDE DE DUTOS", [
        ("CHAPA DE AÇO GALVANIZADA", "CHAPA DE AÇO GALVANIZADA"),
        ("CHAPAS DE AÇO GALVANIZADA", "CHAPA DE AÇO GALVANIZADA"),
        ("CHAPA PRETA", "CHAPA PRETA"),
        ("TUBO DE PVC BRANCO", "TUBO DE PVC BRANCO"),
        ("DERIVAÇÕES", "ACESSÓRIOS EM PVC"),
    ]),
    ("REDE FRIGORÍGENA", "REDE FRIGORÍGENA", [
        ("TUBULAÇÃO DE COBRE", "TUBO DE COBRE"),
        ("CONEXÕES", "CONEXÕES"),
        ("DERIVAÇÕES E ACESSÓRIOS", "DERIVAÇÕES E ACESSÓRIOS"),
    ]),
    ("ACESSÓRIOS", "ACESSÓRIOS", [
        ("", None),
    ]),
]

# Seções fixas que fecham a lista (vêm depois das seções com itens do modelo)
ITEM_ELETRICA = "FORNECIMENTO DA INFRAESTRUTURA (ELETRODUTOS, DISJUNTORES, CABOS, SUPORTES E ETC)"
ITENS_OUTROS = [
    "TESTE E COMISSIONAMENTO DOS SISTEMAS",
    "PROJETO CONFORME EXECUTADO (AS BUILT)",
    "SUPERVISÃO E GERENCIAMENTO",
    "PROGRAMAÇÃO E CONFIGURAÇÕES",
    "PARTIDA DO SISTEMA (START-UP)",
    "TREINAMENTO OPERACIONAL E DOCUMENTAÇÃO TÉCNICA",
    "OPERAÇÃO ASSISTIDA 60 DIAS",
    "MANUTENÇÃO - PERIODO DE GARANTIA",
    "LIMPEZA DIÁRIA DA OBRA",
]
NOTAS = [
    "1 - TODOS OS DUTOS METÁLICOS DEVERÃO SER FORNECIDOS COM SUPORTES.",
    "2 - AS INTERLIGAÇÕES ELÉTRICAS DOS EQUIPAMENTOS, BEM COMO A MONTAGEM DOS PAINÉIS SERÃO À CARGO DA "
    "INSTALADORA E DEVERÃO ATENDER ÀS ESPECIFICAÇÕES DOS FABRICANTES DOS EQUIPAMENTOS.",
    "3 - É DE RESPONSABILIDADE DAS PROPONENTES CONFERIR TODOS OS QUANTITATIVOS E SERVIÇOS ANTES DA ENTREGA DA "
    "PROPOSTA DE PREÇO. CASO HAJA ALGUMA DIVERGÊNCIA OU DÚVIDA DEVERÁ SER COMUNICADO À CONTRATANTE "
    "PREVIAMENTE AO ENVIO DA PROPOSTA.",
    "4 - OS QUANTITATIVOS APRESENTADOS NESTA PLANILHA FORAM OBTIDOS A PARTIR DE INFORMAÇÕES REGISTRADAS NOS "
    "PROJETOS. PEQUENAS VARIAÇÕES NAS QUANTIDADES PODERÃO OCORRER DEVIDO A ALTERAÇÕES DE PERCURSOS, PROJETOS "
    "REFERÊNCIA MODIFICADOS NA OBRA OU OUTRAS INTERFERÊNCIAS.",
    "5 - QUALQUER ITEM OMISSO NA PLANILHA E QUE SEJA NECESSÁRIO PARA EXECUÇÃO COMPLETA DO SERVIÇO DEVERÁ SER "
    "CONSIDERADO NO PREÇO REFERENTE AO ITEM, INCLUSIVE FUROS EM VIGAS, LAJES E PAREDES PARA PASSAGEM DAS "
    "INSTALAÇÕES.",
    "6 - PARA INFORMAÇÕES COMPLEMENTARES, VER MEMORIAL DESCRITIVO.",
]
NOTA_DESENHOS = "7 - PARA DOCUMENTOS QUE FUNDAMENTARAM A ELABORAÇÃO DESTA LISTA, VER DESENHOS {desenhos}"
NOTA_DESENHOS_SEM_REF = "7 - PARA DOCUMENTOS QUE FUNDAMENTARAM A ELABORAÇÃO DESTA LISTA, VER DESENHOS DO PROJETO."

# ---------------------------------------------------------------------------
# Abas
# ---------------------------------------------------------------------------
ABA_COMUM = "ÁREAS DE USO COMUM"
ABA_PRIVATIVA = "ÁREAS PRIVATIVAS"
# Níveis cujo nome contém estes termos vão, por padrão, para áreas de uso comum.
TERMOS_AREA_COMUM = ["SUBSOLO", "TÉRREO", "TERREO", "PILOTIS", "MEZANINO", "LAZER", "GARAGEM",
                     "PLAYGROUND", "BARRILETE", "CASA DE MÁQUINAS", "CASA DE MAQUINAS", "ÁTICO", "ATICO"]

# ---------------------------------------------------------------------------
# Locais atendidos
# ---------------------------------------------------------------------------
# Local atendido com vários ambientes: nomes separados por "-" sem espaço (ex.: SAT-MOTORISTA).
# Hífen com espaço (" - ") faz parte do nome (ex.: APARTAMENTO FINAL 02 - OPÇÃO 1).
# Prefixos que não separam ambientes (ex.: PRÉ-ESCOLA continua um ambiente só).
PREFIXOS_HIFENIZADOS = {"PRÉ", "PRE", "PÓS", "POS", "SEMI", "SUB", "MEIA", "GUARDA", "PORTA", "BEM",
                        "MINI", "MULTI", "AUTO", "ANTE", "SUPER", "CO", "EX", "VICE", "RECÉM", "RECEM", "MEIO"}
COLUNA_GERAL = "GERAL"
ABREVIACOES_LOCAL = [("APARTAMENTO ", "APTO. ")]
# Peças com nível de referência diferente do ambiente atendido (ex.: condensadora no 2º subsolo
# atendendo a academia do 1º subsolo) vão para o ambiente atendido quando são poucas.
REALOCAR_MAX_INSTANCIAS = 3
REALOCAR_MAX_FRACAO = 0.10

# ---------------------------------------------------------------------------
# Validação do CSV
# ---------------------------------------------------------------------------
MARCADOR_NF = "NF"

# ---------------------------------------------------------------------------
# Folha de rosto (células do modelo MODELO_-FOLHA_DE_ROSTO.xlsx)
# ---------------------------------------------------------------------------
ABA_CAPA = "F. DE ROSTO"
ABA_FOTOS = "FOTOS"
CELULAS_CAPA = {
    "ca": "L4",
    "obra": "A5",
    "sistema": "A6",
    "etapa": "A7",
    "documento": "A8",
    "rev": "A39", "te": "B39", "descricao": "C39",
    "elab": "I39", "verif": "J39", "resp": "K39", "data": "L39",
}
PADRAO_DOCUMENTO = "{obra}-0000-CLEX-LM-TT-GER-LMT-R00"
TIPOS_EMISSAO = ["A", "B", "C", "D", "E"]

# ---------------------------------------------------------------------------
# Visual da tabela (padrão das listas A056 / 648)
# ---------------------------------------------------------------------------
FONTE = "Franklin Gothic Book"
TAM_FONTE = 15
COR_CABECALHO = "FFBFBFBF"
COR_SECAO = "FFD9D9D9"
COR_ABA = "FF10475C"
LARGURAS_FIXAS = {"A": 11.875, "B": 58.125, "C": 74.32, "D": 15.625, "E": 15.625}
UNIDADES_DECIMAIS = {"m", "kg", "m²", "m2", "m³", "m3", "l"}
