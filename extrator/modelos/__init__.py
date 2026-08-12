"""Registro dos modelos de ficha reconhecidos pelo extrator.

Para dar suporte a um novo tipo de documento basta criar um módulo com um
:class:`~extrator.modelos.base.Modelo` (a exemplo de ``indicador.py`` e
``iniciativa.py``), acrescentar suas seções em :data:`SECOES` e registrá-lo
em :data:`MODELOS`.
"""

from __future__ import annotations

from . import indicador, iniciativa
from .base import (
    COLUNA_ARQUIVO,
    COLUNAS_AUDITORIA_FIM,
    COLUNAS_AUDITORIA_INICIO,
    Campo,
    Modelo,
)

#: Modelos reconhecidos, na ordem em que a classificação os testa.
MODELOS: tuple[Modelo, ...] = (indicador.MODELO, iniciativa.MODELO)

MODELOS_POR_CODIGO: dict[str, Modelo] = {modelo.codigo: modelo for modelo in MODELOS}

#: Cabeçalhos que delimitam as seções das fichas, na forma
#: ``(regex sobre a chave canônica, identificador da seção)``.
#:
#: A lista é a **união** dos dois modelos: a leitura do documento acontece
#: antes da classificação, então o leitor precisa reconhecer as seções de
#: qualquer tipo de ficha. Seções equivalentes entre os modelos compartilham
#: o mesmo identificador (ex.: "VÍNCULO DO INDICADOR..." e "VÍNCULO DA
#: INICIATIVA" viram ambos ``VINCULO``), o que mantém o mapa de campos legível.
SECOES: tuple[tuple[str, str], ...] = (
    (r"vinculodoindicadordecompromisso", "VINCULO"),
    (r"vinculodainiciativa", "VINCULO"),
    (r"atributosdo?indicadordecompromisso", "ATRIBUTOS"),
    (r"atributosdainiciativa", "ATRIBUTOS"),
    (r"desagregacaoterritorial", "TERRITORIAL"),
    (r"informacoescomplementares", "COMPLEMENTARES"),
    (r"fatorescriticosde?contextodainiciativa", "FATORES"),
    (r"estimativadosrecursos.*", "RECURSOS"),
    (r"indicadorde?compromissovinculado", "INDICADORES"),
    # A ordem importa: vence o primeiro padrão que casar. O cabeçalho dos
    # produtos ("... Produtos das ações orçamentárias vinculadas ...") também
    # contém o texto do cabeçalho das ações, então precisa ser testado antes.
    (r"etapaii.*produtosdasacoesorcamentarias.*", "PRODUTOS_ORCAMENTARIOS"),
    (r"etapaii.*acoesorcamentariasvinculadas.*", "ACOES_ORCAMENTARIAS"),
)

#: Todos os rótulos conhecidos de todos os modelos: rótulos de campo, âncoras
#: de cabeçalho dos extratores de tabela, rótulos estruturais e cabeçalhos de
#: seção. Serve para o motor distinguir "isto é um rótulo" de "isto é um
#: valor" — inclusive para saber onde uma lista termina.
ROTULOS_CONHECIDOS: tuple[str, ...] = (
    tuple(
        padrao
        for modelo in MODELOS
        for campo in modelo.campos
        for padrao in (*campo.padroes, *campo.parametro)
    )
    + indicador.ROTULOS_ESTRUTURAIS
    + iniciativa.ROTULOS_ESTRUTURAIS
    + tuple(padrao for padrao, _ in SECOES)
)

__all__ = [
    "COLUNAS_AUDITORIA_FIM",
    "COLUNAS_AUDITORIA_INICIO",
    "COLUNA_ARQUIVO",
    "MODELOS",
    "MODELOS_POR_CODIGO",
    "ROTULOS_CONHECIDOS",
    "SECOES",
    "Campo",
    "Modelo",
]
