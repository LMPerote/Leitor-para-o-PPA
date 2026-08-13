"""Modelo da FICHA DE CONTROLE (capa do diretório do compromisso).

É a folha que acompanha cada pasta de compromisso e registra quem digitou no
Fiplan, quando, quantas fichas de indicador e de iniciativa existem e o que
ficou pendente. Não descreve indicador nem iniciativa — por isso vira uma
planilha própria.

O layout é uma tabela única em que rótulo e valor moram na mesma célula,
separados por quebra de linha ("EIXO:" na primeira linha, o valor na segunda).
"""

from __future__ import annotations

from .base import Campo, Modelo

CAMPOS: tuple[Campo, ...] = (
    Campo(
        coluna="Nome_Diretorio_Compromisso",
        padroes=(r"nomedodiretoriodocompromisso", r"nomedodiretorio"),
        descricao="Nome do diretório (pasta) do compromisso.",
    ),
    Campo(
        coluna="Eixo",
        padroes=(r"eixo",),
        descricao="Eixo estruturante.",
    ),
    Campo(
        coluna="Programa",
        padroes=(r"programa",),
        descricao="Programa do PPA.",
    ),
    Campo(
        coluna="Compromisso",
        padroes=(r"compromisso",),
        descricao="Compromisso a que a pasta se refere.",
    ),
    Campo(
        coluna="Nome_Digitador_Fiplan",
        padroes=(r"nomedigitador[ao]s?fiplan", r"digitador[ao]fiplan"),
        descricao="Quem digitou as informações no Fiplan.",
    ),
    Campo(
        coluna="Data_Insercao_Fiplan",
        padroes=(r"datainsercaonofiplan", r"datade?insercao.*"),
        descricao="Data de inserção no Fiplan.",
    ),
    Campo(
        coluna="Qtd_Indicadores_Compromisso",
        padroes=(r"n?de?indicadoresdecompromisso", r"n?de?indicadores"),
        descricao="Quantidade de fichas de Indicador de Compromisso na pasta.",
    ),
    Campo(
        coluna="Qtd_Fichas_Iniciativas",
        padroes=(r"n?de?fichasde?iniciativas?", r"n?de?iniciativas?"),
        descricao="Quantidade de fichas de Iniciativa na pasta.",
    ),
    Campo(
        coluna="Pendencias_Observacoes",
        padroes=(r"pendenciase?observacoes",),
        multiplo=True,
        descricao="Pendências e observações anotadas na ficha de controle.",
    ),
)

ROTULOS_ESTRUTURAIS: tuple[str, ...] = (
    r"nomedodiretoriodocompromisso",
    r"pendenciase?observacoes",
)

MODELO = Modelo(
    codigo="CONTROLE",
    rotulo="Ficha de Controle",
    arquivo_saida="Controles.xlsx",
    # O nome do diretório só existe nesta ficha; as fichas de indicador e de
    # iniciativa são reconhecidas antes, pelas suas seções de vínculo.
    deteccao=(r"nomedodiretoriodocompromisso", r"pendenciase?observacoes"),
    campos=CAMPOS,
)
