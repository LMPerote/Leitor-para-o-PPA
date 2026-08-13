"""Modelo da ficha de INICIATIVA."""

from __future__ import annotations

from .base import Campo, Modelo

CAMPOS: tuple[Campo, ...] = (
    # --- 1. VÍNCULO DA INICIATIVA -----------------------------------------
    Campo(
        coluna="Eixo",
        padroes=(r"eixo",),
        secao="VINCULO",
        descricao="Eixo estruturante ao qual a iniciativa está vinculada.",
    ),
    Campo(
        coluna="Programa",
        padroes=(r"programa",),
        secao="VINCULO",
        descricao="Programa do PPA vinculado à iniciativa.",
    ),
    Campo(
        coluna="Compromisso",
        padroes=(r"compromisso",),
        secao="VINCULO",
        descricao="Compromisso ao qual a iniciativa se refere.",
    ),
    Campo(
        coluna="Problemas_Vinculados",
        padroes=(r"problemas?vinculados?aocompromisso", r"problemas?vinculados?"),
        secao="VINCULO",
        multiplo=True,
        descricao="Problema(s) vinculado(s) ao Compromisso.",
    ),
    Campo(
        coluna="Causas_Criticas",
        padroes=(r"causas?criticas?",),
        secao="VINCULO",
        multiplo=True,
        descricao="Causa(s) crítica(s), unidas na mesma célula.",
    ),
    Campo(
        coluna="Acoes_Criticas",
        # "Ação(ões) Crítica(s)" vira a chave "acaooescriticas"; os demais
        # ramos cobrem "Ações Críticas" e "Ação Crítica".
        padroes=(r"ac(aooes|oes|ao)s?criticas?",),
        secao="VINCULO",
        multiplo=True,
        descricao="Ação(ões) crítica(s), unidas na mesma célula.",
    ),
    Campo(
        coluna="Propostas_Escuta_Social",
        extrator="tabela_estruturada",
        parametro=(r"propostas?de?escutasocialassociadas?aocompromisso",),
        secao="VINCULO",
        descricao=(
            "Propostas de escuta social associadas ao compromisso, com status "
            "para atendimento e justificativa (uma proposta por linha)."
        ),
    ),
    # --- 2. ATRIBUTOS DA INICIATIVA ---------------------------------------
    Campo(
        coluna="Atributos_Descricao",
        padroes=(r"descricao",),
        secao="ATRIBUTOS",
        descricao="Descrição da iniciativa.",
    ),
    Campo(
        coluna="Entregas_Vinculadas",
        padroes=(r"entregas?vinculadas?",),
        secao="ATRIBUTOS",
        multiplo=True,
        descricao="Entrega(s) vinculada(s) à iniciativa.",
    ),
    Campo(
        coluna="Responsavel_Sigla_Orgao",
        padroes=(r"sigladoorgao", r"siglaorgao"),
        secao="ATRIBUTOS",
        descricao="Sigla do órgão responsável pela iniciativa.",
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
    Campo(
        coluna="Orgaos_Parceiros",
        padroes=(r"orgaos?parceiros?",),
        secao="ATRIBUTOS",
        multiplo=True,
        descricao="Órgãos parceiros na execução da iniciativa.",
    ),
    # --- 3. FATORES CRÍTICOS DE CONTEXTO ----------------------------------
    Campo(
        coluna="Fatores_Criticos_Operacionais",
        padroes=(r"operacionais",),
        secao="FATORES",
        descricao="Fatores críticos operacionais.",
    ),
    Campo(
        coluna="Fatores_Criticos_Orcamentarios",
        # Há versões do modelo em que o rótulo é apenas "Financeiros".
        padroes=(
            r"orcamentari[oa]s?e?financeir[oa]s?",
            r"orcamentari[oa]s?",
            r"financeir[oa]s?",
        ),
        secao="FATORES",
        descricao="Fatores críticos orçamentários/financeiros.",
    ),
    Campo(
        coluna="Fatores_Criticos_Institucionais",
        padroes=(r"institucionaisou?politic[oa]s", r"institucionais.*"),
        secao="FATORES",
        descricao="Fatores críticos institucionais ou políticos.",
    ),
    # --- 4. RECURSOS ORÇAMENTÁRIOS ----------------------------------------
    Campo(
        coluna="Recursos_Orcamentarios",
        extrator="recursos_orcamentarios",
        secao="RECURSOS",
        descricao=(
            "Estimativa dos recursos: código e nome da fonte, montante em R$ "
            "por fonte e total dos recursos."
        ),
    ),
    # --- 5. INDICADORES VINCULADOS ----------------------------------------
    Campo(
        coluna="Indicador_Compromisso_Vinculado",
        padroes=(r"indicadorde?compromissovinculado",),
        secao="INDICADORES",
        multiplo=True,
        descricao="Indicador de Compromisso vinculado à iniciativa.",
    ),
    Campo(
        coluna="Indicadores_Sensibilizados",
        padroes=(r"indicadores?de?compromissosensibilizados?",),
        secao="INDICADORES",
        multiplo=True,
        descricao="Indicadores de Compromisso sensibilizados pela iniciativa.",
    ),
    # --- 6. PROGRAMAÇÃO ORÇAMENTÁRIA (ETAPA II) ---------------------------
    Campo(
        coluna="Acoes_Orcamentarias_Vinculadas",
        extrator="tabela_estruturada",
        parametro=(r"orgao",),
        secao="ACOES_ORCAMENTARIAS",
        descricao=(
            "Ações orçamentárias vinculadas: órgão, UO, código e nome da ação "
            "(uma ação por linha)."
        ),
    ),
    Campo(
        coluna="Produtos_Acoes_Orcamentarias",
        extrator="tabela_estruturada",
        parametro=(r"orgao",),
        secao="PRODUTOS_ORCAMENTARIOS",
        descricao=(
            "Produtos das ações orçamentárias: código e nome do produto "
            "(um produto por linha)."
        ),
    ),
)

#: Rótulos estruturais desta ficha que não viram coluna, mas que o motor
#: precisa conhecer para não confundi-los com *valores*.
ROTULOS_ESTRUTURAIS: tuple[str, ...] = (
    r"vinculodainiciativa",
    r"atributosdainiciativa",
    r"responsavelpelainiciativa",
    r"fatorescriticosde?contextodainiciativa",
    r"estimativadosrecursos.*",
    r"etapaii.*",
    r"statuspara?atendimento",
    r"justificativas?",
    r"orgao",
    r"codigodafonte",
    r"nomedafonte",
    r"montanteemrs?",
    r"totaldosrecursos",
    r"codigoda?acaoorcamentaria",
    r"nomeda?acaoorcamentaria",
    r"codigodo?produtoda?acaoorcamentaria",
    r"nomeprodutoda?acaoorcamentaria",
)

MODELO = Modelo(
    codigo="INICIATIVA",
    rotulo="Iniciativa",
    arquivo_saida="Iniciativas.xlsx",
    deteccao=(r"vinculodainiciativa", r"atributosdainiciativa"),
    campos=CAMPOS,
)
