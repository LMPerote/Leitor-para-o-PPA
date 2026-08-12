"""Mapa de campos: o "contrato" entre a ficha Word e as colunas do Excel.

Cada :class:`Campo` descreve uma coluna da planilha final e como localizá-la
no documento. Os padrões são expressões regulares aplicadas sobre a *chave
canônica* do texto (ver :mod:`extrator.texto`), ou seja: sem acentos, sem
espaços e sem pontuação. Por isso os padrões são escritos "grudados"
(``formuladecalculo``) e toleram naturalmente erros de digitação de espaços
que aparecem nos documentos reais.

Para adaptar o extrator a um novo modelo de ficha, normalmente basta editar
este arquivo — o motor de extração não precisa ser alterado.
"""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass(frozen=True)
class Campo:
    """Descreve uma coluna do Excel e como extraí-la do documento."""

    coluna: str
    padroes: tuple[str, ...] = ()
    secao: str | None = None
    #: Qual ocorrência do rótulo usar quando ele se repete na seção (0 = 1ª).
    ocorrencia: int = 0
    #: Se True, concatena todas as células seguintes até o próximo rótulo.
    multiplo: bool = False
    #: Nome de um extrator especial em :mod:`extrator.parser` (layouts que
    #: não seguem o padrão "rótulo -> célula à direita/abaixo").
    extrator: str | None = None
    descricao: str = ""


# ---------------------------------------------------------------------------
# Colunas de auditoria (preenchidas pelo pipeline, não pelo parser)
# ---------------------------------------------------------------------------
COLUNA_ARQUIVO = "Nome_do_Arquivo"
COLUNAS_AUDITORIA_INICIO = (COLUNA_ARQUIVO, "Caminho_Relativo")
COLUNAS_AUDITORIA_FIM = ("Status", "Campos_Nao_Encontrados", "Observacoes")

DESCRICOES_AUDITORIA = {
    COLUNA_ARQUIVO: "Nome do arquivo .docx de origem (rastreabilidade).",
    "Caminho_Relativo": "Caminho do arquivo em relação à pasta de entrada.",
    "Status": "OK, OK_COM_PENDENCIAS ou ERRO_DE_LEITURA.",
    "Campos_Nao_Encontrados": "Colunas cujo rótulo não foi localizado no documento.",
    "Observacoes": "Mensagem de erro ou aviso gerado durante o processamento.",
}


# ---------------------------------------------------------------------------
# Mapa de campos da ficha de Indicador de Compromisso (PPA)
# ---------------------------------------------------------------------------
CAMPOS: tuple[Campo, ...] = (
    # --- 1. VÍNCULO DO INDICADOR DE COMPROMISSO ---------------------------
    Campo(
        coluna="Eixo",
        padroes=(r"eixo",),
        secao="VINCULO",
        descricao="Eixo estruturante ao qual o indicador está vinculado.",
    ),
    Campo(
        coluna="Programa",
        padroes=(r"programa",),
        secao="VINCULO",
        descricao="Programa do PPA vinculado ao indicador.",
    ),
    Campo(
        coluna="Compromisso",
        padroes=(r"compromisso",),
        secao="VINCULO",
        descricao="Compromisso ao qual o indicador se refere.",
    ),
    Campo(
        coluna="Problemas_Vinculados",
        padroes=(r"problemas?vinculados?aocompromisso", r"problemasvinculados?"),
        secao="VINCULO",
        multiplo=True,
        descricao="Problema(s) vinculado(s) ao Compromisso.",
    ),
    Campo(
        coluna="Causas_Criticas",
        padroes=(r"causas?criticas?",),
        secao="VINCULO",
        multiplo=True,
        descricao="Lista de causas críticas (itens unidos no mesmo campo).",
    ),
    # --- 2. ATRIBUTOS DO INDICADOR ---------------------------------------
    Campo(
        coluna="Descricao",
        padroes=(r"descricao",),
        secao="ATRIBUTOS",
        descricao="Descrição do indicador.",
    ),
    Campo(
        coluna="Formula_de_Calculo",
        padroes=(r"formulade?calculo",),
        secao="ATRIBUTOS",
        descricao="Fórmula de cálculo do indicador.",
    ),
    Campo(
        coluna="Memoria_de_Calculo",
        padroes=(r"memoriade?calculo",),
        secao="ATRIBUTOS",
        descricao="Memória de cálculo do indicador.",
    ),
    Campo(
        coluna="Unidade_de_Medida",
        padroes=(r"unidadede?medida",),
        secao="ATRIBUTOS",
        descricao="Unidade de medida do indicador.",
    ),
    Campo(
        coluna="Valor_de_Referencia",
        padroes=(r"valorde?referencia",),
        secao="ATRIBUTOS",
        descricao="Valor de referência (linha de base).",
    ),
    Campo(
        coluna="Ano_de_Referencia",
        padroes=(r"anode?referencia",),
        secao="ATRIBUTOS",
        descricao="Ano do valor de referência.",
    ),
    Campo(
        coluna="Valor_da_Meta",
        padroes=(r"valordameta", r"valorde?meta"),
        secao="ATRIBUTOS",
        descricao="Valor da meta previsto.",
    ),
    Campo(
        coluna="Periodicidade_da_Apuracao",
        padroes=(r"periodicidade(da)?apuracao", r"periodicidade"),
        secao="ATRIBUTOS",
        descricao="Periodicidade da apuração (mensal, semestral, anual...).",
    ),
    Campo(
        coluna="Polaridade",
        padroes=(r"polaridade",),
        secao="ATRIBUTOS",
        descricao="Polaridade do indicador (positiva/negativa).",
    ),
    Campo(
        coluna="Classificacao",
        padroes=(r"classificacao",),
        secao="ATRIBUTOS",
        descricao="Classificação do indicador (produto, resultado, etc.).",
    ),
    Campo(
        coluna="Fonte",
        padroes=(r"fontes?",),
        secao="ATRIBUTOS",
        multiplo=True,
        descricao="Fonte dos dados.",
    ),
    Campo(
        coluna="Meios_de_Verificacao",
        padroes=(r"meiosde?verificacao", r"meiode?verificacao"),
        secao="ATRIBUTOS",
        multiplo=True,
        descricao="Meios de verificação do indicador.",
    ),
    Campo(
        coluna="Responsavel_Sigla_Orgao",
        padroes=(r"sigladoorgao", r"siglaorgao"),
        secao="ATRIBUTOS",
        descricao="Sigla do órgão responsável pelo indicador.",
    ),
    Campo(
        coluna="Responsavel_UO",
        padroes=(r"uo",),
        secao="ATRIBUTOS",
        descricao="Unidade Orçamentária responsável.",
    ),
    Campo(
        coluna="Responsavel_USP",
        padroes=(r"usp",),
        secao="ATRIBUTOS",
        descricao="Unidade do Sistema de Planejamento responsável.",
    ),
    # --- 3. DESAGREGAÇÃO TERRITORIAL --------------------------------------
    Campo(
        coluna="Desagregacao_Estado",
        extrator="marcacao_estado",
        secao="TERRITORIAL",
        descricao="Sim/Não — desagregação no nível Estado está marcada.",
    ),
    Campo(
        coluna="Desagregacao_Territorio_Identidade",
        extrator="marcacao_territorio",
        secao="TERRITORIAL",
        descricao="Sim/Não — desagregação por Território de Identidade marcada.",
    ),
    Campo(
        coluna="Formula_Calculo_Territorial",
        padroes=(r"formulade?calculoterritorial",),
        secao="TERRITORIAL",
        descricao="Fórmula de cálculo da desagregação territorial.",
    ),
    Campo(
        coluna="Unidade_Medida_Territorial",
        padroes=(r"unidadede?medida",),
        secao="TERRITORIAL",
        descricao="Unidade de medida da desagregação territorial.",
    ),
    Campo(
        coluna="Memoria_Calculo_Territorial",
        padroes=(r"memoriade?calculo",),
        secao="TERRITORIAL",
        descricao="Memória de cálculo da desagregação territorial.",
    ),
    Campo(
        coluna="Territorios_de_Identidade",
        extrator="lista_territorios",
        secao="TERRITORIAL",
        descricao="Territórios de Identidade listados na ficha.",
    ),
    Campo(
        coluna="Memoria_Calculo_Por_Territorio",
        extrator="lista_memoria_territorial",
        secao="TERRITORIAL",
        descricao="Memória de cálculo por território (uma linha por território).",
    ),
    Campo(
        coluna="Metas_Territoriais",
        extrator="lista_metas_territoriais",
        secao="TERRITORIAL",
        descricao="Metas territoriais (uma linha por território).",
    ),
    Campo(
        coluna="Outras_Possibilidades_Regionalizacao",
        padroes=(r"outraspossibilidadesde?regionalizacao", r"outraspossibilidades.*"),
        secao="TERRITORIAL",
        descricao="Outras possibilidades de regionalização.",
    ),
    # --- 4. INFORMAÇÕES COMPLEMENTARES ------------------------------------
    Campo(
        coluna="Objetivo_Interpretacao_Uso",
        padroes=(r"objetivo(interpretacaoeuso)?", r"objetivointerpretacao.*"),
        secao="COMPLEMENTARES",
        multiplo=True,
        descricao="Objetivo / interpretação e uso do indicador.",
    ),
    Campo(
        coluna="Limitacoes_do_Indicador",
        padroes=(r"limitacoesdoindicador",),
        secao="COMPLEMENTARES",
        multiplo=True,
        descricao="Limitações do indicador.",
    ),
    Campo(
        coluna="Fragilidades_para_Apuracao",
        padroes=(r"fragilidadespara?apuracao.*",),
        secao="COMPLEMENTARES",
        multiplo=True,
        descricao="Fragilidades para apuração e ações de superação.",
    ),
    Campo(
        coluna="Limitacoes_Meta_Operacionais",
        padroes=(r"operacionais",),
        secao="COMPLEMENTARES",
        descricao="Limitações operacionais para definição do valor da meta.",
    ),
    Campo(
        coluna="Limitacoes_Meta_Orcamentarias",
        padroes=(r"orcamentarias?financeiras?", r"orcamentarias?"),
        secao="COMPLEMENTARES",
        descricao="Limitações orçamentárias/financeiras para definição da meta.",
    ),
    Campo(
        coluna="Limitacoes_Meta_Institucionais",
        padroes=(r"institucionaisou?politicas", r"institucionais.*"),
        secao="COMPLEMENTARES",
        descricao="Limitações institucionais ou políticas para definição da meta.",
    ),
    Campo(
        coluna="Possibilidade_Desagregacao_Populacional",
        padroes=(r"possibilidadede?desagregacaopopulacional",),
        secao="COMPLEMENTARES",
        multiplo=True,
        descricao="Possibilidade de desagregação populacional.",
    ),
    Campo(
        coluna="Programas_Especiais",
        extrator="programas_especiais",
        secao="COMPLEMENTARES",
        descricao="Programas Especiais (nome, memória de cálculo e meta).",
    ),
    Campo(
        coluna="Indicadores_Programa_Sensibilizado",
        padroes=(r"indicadores?do?programasensibilizados?",),
        secao="COMPLEMENTARES",
        multiplo=True,
        descricao="Indicador(es) do Programa Sensibilizado(s).",
    ),
)


#: Rótulos estruturais que não viram coluna, mas que o motor precisa conhecer
#: para não confundi-los com *valores* (ex.: em "Unidade de medida | Valor de
#: referência", a célula à direita é outro rótulo, não o valor procurado).
ROTULOS_ESTRUTURAIS: tuple[str, ...] = (
    r"vinculodoindicadordecompromisso",
    r"atributosdo?indicadordecompromisso",
    r"desagregacaoterritorial",
    r"informacoescomplementares",
    r"responsavelpeloindicador",
    r"limitacoesparade?finicaodovalordameta",
    r"limitacoespara.*meta",
    r"programasespeciais",
    r"nomedoprograma",
    r"meta",
    r"metaterritorial",
    r"memoriade?calculoterritorial",
    r"territoriode?identidade",
    r"estado",
    r"anodameta",
)


def colunas_de_dados() -> list[str]:
    return [campo.coluna for campo in CAMPOS]


def todas_as_colunas() -> list[str]:
    return [*COLUNAS_AUDITORIA_INICIO, *colunas_de_dados(), *COLUNAS_AUDITORIA_FIM]


def dicionario_de_campos() -> list[dict[str, str]]:
    """Linhas da aba 'Dicionário' da planilha final."""
    linhas = [
        {"Coluna": nome, "Seção": "Auditoria", "Descrição": DESCRICOES_AUDITORIA[nome]}
        for nome in COLUNAS_AUDITORIA_INICIO
    ]
    linhas += [
        {
            "Coluna": campo.coluna,
            "Seção": campo.secao or "Documento",
            "Descrição": campo.descricao,
        }
        for campo in CAMPOS
    ]
    linhas += [
        {"Coluna": nome, "Seção": "Auditoria", "Descrição": DESCRICOES_AUDITORIA[nome]}
        for nome in COLUNAS_AUDITORIA_FIM
    ]
    return linhas
