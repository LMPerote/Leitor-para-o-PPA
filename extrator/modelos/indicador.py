"""Modelo da ficha de INDICADOR DE COMPROMISSO."""

from __future__ import annotations

from .base import Campo, Modelo

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
        padroes=(r"problemas?vinculados?aocompromisso", r"problemas?vinculados?"),
        secao="VINCULO",
        multiplo=True,
        descritivo=True,
        descricao="Problema(s) vinculado(s) ao Compromisso.",
    ),
    Campo(
        coluna="Causas_Criticas",
        padroes=(r"causas?criticas?",),
        secao="VINCULO",
        multiplo=True,
        descritivo=True,
        descricao="Causa(s) crítica(s), unidas na mesma célula.",
    ),
    # --- 2. ATRIBUTOS DO INDICADOR ----------------------------------------
    Campo(
        coluna="Atributos_Descricao",
        padroes=(r"descricao",),
        secao="ATRIBUTOS",
        descricao="Descrição do indicador.",
    ),
    Campo(
        coluna="Formula_Calculo",
        padroes=(r"formulade?calculo",),
        secao="ATRIBUTOS",
        descricao="Fórmula de cálculo do indicador.",
    ),
    Campo(
        coluna="Memoria_Calculo",
        padroes=(r"memoriade?calculo",),
        secao="ATRIBUTOS",
        descricao="Memória de cálculo do indicador.",
    ),
    Campo(
        coluna="Unidade_Medida",
        padroes=(r"unidadede?medida",),
        secao="ATRIBUTOS",
        descricao="Unidade de medida do indicador.",
    ),
    Campo(
        coluna="Valor_Referencia",
        padroes=(r"valorde?referencia",),
        secao="ATRIBUTOS",
        descricao="Valor de referência (linha de base).",
    ),
    Campo(
        coluna="Ano_Referencia",
        padroes=(r"anode?referencia",),
        secao="ATRIBUTOS",
        descricao="Ano do valor de referência.",
    ),
    Campo(
        coluna="Valor_Meta",
        padroes=(r"valordameta", r"valorde?meta"),
        secao="ATRIBUTOS",
        descricao="Valor da meta previsto.",
    ),
    Campo(
        coluna="Periodicidade_Apuracao",
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
        coluna="Meios_Verificacao",
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
        coluna="Desagregacao_Territorial",
        extrator="desagregacao_territorial",
        secao="TERRITORIAL",
        descricao=(
            "Bloco consolidado da desagregação territorial: marcação de Estado e "
            "Território de Identidade, fórmula, unidade, memória de cálculo, "
            "territórios/metas e outras regionalizações."
        ),
    ),
    # --- 4. INFORMAÇÕES COMPLEMENTARES ------------------------------------
    Campo(
        coluna="Info_Complementares_Objetivo",
        padroes=(r"objetivo(interpretacaoeuso)?", r"objetivointerpretacao.*"),
        secao="COMPLEMENTARES",
        multiplo=True,
        descricao="Objetivo / interpretação e uso do indicador.",
    ),
    Campo(
        coluna="Limitacoes_Indicador",
        padroes=(r"limitacoesdoindicador",),
        secao="COMPLEMENTARES",
        multiplo=True,
        descricao="Limitações do indicador.",
    ),
    Campo(
        coluna="Fragilidades_Apuracao",
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
        # Há versões do modelo em que o rótulo é apenas "Financeiros".
        padroes=(
            r"orcamentari[oa]s?e?financeir[oa]s?",
            r"orcamentari[oa]s?",
            r"financeir[oa]s?",
        ),
        secao="COMPLEMENTARES",
        descricao="Limitações orçamentárias/financeiras para definição da meta.",
    ),
    Campo(
        coluna="Limitacoes_Meta_Institucionais",
        padroes=(r"institucionaisou?politic[oa]s", r"institucionais.*"),
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
        extrator="tabela_estruturada",
        parametro=("nomedoprograma",),
        secao="COMPLEMENTARES",
        descricao="Programas Especiais (nome, memória de cálculo e meta).",
    ),
)

#: Rótulos estruturais desta ficha que não viram coluna, mas que o motor
#: precisa conhecer para não confundi-los com *valores*.
ROTULOS_ESTRUTURAIS: tuple[str, ...] = (
    r"vinculodoindicadord[eo]compromisso",
    r"atributosdo?indicadord[eo]compromisso",
    r"desagregacaoterritorial",
    r"informacoescomplementares",
    r"responsavelpeloindicador",
    r"limitacoespara.*meta",
    r"programasespeciais",
    r"nomedoprograma",
    r"meta",
    r"metaterritorial",
    r"memoriade?calculoterritorial",
    r"territoriode?identidade",
    r"estado",
    r"anodameta",
    r"outraspossibilidadesde?regionalizacao",
    r"criteriospara?distribuicaoterritorial",
    r"desagregacaoterritorialregional.*",
    r"indicadores?do?programasensibilizados?",
)

MODELO = Modelo(
    codigo="INDICADOR",
    rotulo="Indicador de Compromisso",
    arquivo_saida="Indicadores.xlsx",
    deteccao=(
        r"vinculodoindicadord[eo]compromisso",
        r"atributosdo?indicadord[eo]compromisso",
    ),
    campos=CAMPOS,
)
