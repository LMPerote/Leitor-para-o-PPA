"""Estruturas que descrevem um *modelo de ficha* (tipo de documento).

Um :class:`Modelo` amarra três coisas:

* como **reconhecer** o tipo do documento (padrões da seção de vínculo);
* **quais campos** extrair e como localizá-los (lista de :class:`Campo`);
* **para qual planilha** os registros vão (``arquivo_saida``).

Os padrões são expressões regulares aplicadas sobre a *chave canônica* do
texto (ver :mod:`extrator.texto`): minúsculas, sem acentos, sem espaços e sem
pontuação. Por isso são escritos "grudados" (``formuladecalculo``) e toleram
naturalmente as variações de digitação que aparecem nos documentos reais.
"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class Campo:
    """Descreve uma coluna da planilha e como extraí-la do documento."""

    coluna: str
    #: Regex (sobre a chave canônica) que identificam o rótulo do campo.
    padroes: tuple[str, ...] = ()
    #: Seção da ficha onde procurar; ``None`` procura no documento inteiro.
    secao: str | None = None
    #: Qual ocorrência do rótulo usar quando ele se repete na seção (0 = 1ª).
    ocorrencia: int = 0
    #: Se True, concatena todas as células seguintes até o próximo rótulo.
    multiplo: bool = False
    #: Nome de um extrator especial de :class:`extrator.parser.Extrator`,
    #: para layouts que fogem do padrão "rótulo -> célula à direita/abaixo".
    extrator: str | None = None
    #: Argumentos do extrator especial (ex.: a chave-âncora do cabeçalho).
    parametro: tuple[str, ...] = ()
    descricao: str = ""


# ---------------------------------------------------------------------------
# Colunas de auditoria — comuns aos dois modelos, preenchidas pelo pipeline
# ---------------------------------------------------------------------------
COLUNA_ARQUIVO = "Nome_Arquivo"
COLUNAS_AUDITORIA_INICIO: tuple[str, ...] = (COLUNA_ARQUIVO, "Caminho_Relativo")
COLUNAS_AUDITORIA_FIM: tuple[str, ...] = (
    "Status",
    "Campos_Nao_Encontrados",
    "Observacoes",
)

DESCRICOES_AUDITORIA: dict[str, str] = {
    COLUNA_ARQUIVO: "Nome do arquivo .docx de origem (rastreabilidade).",
    "Caminho_Relativo": "Caminho do arquivo em relação à pasta de entrada.",
    "Status": "OK, OK_COM_PENDENCIAS ou ERRO_DE_LEITURA.",
    "Campos_Nao_Encontrados": (
        "Colunas sem valor: rótulo não localizado no documento ou localizado "
        "sem resposta preenchida."
    ),
    "Observacoes": (
        "Avisos do processamento: campos sem valor e itens que remetem a outro "
        "item em vez de descrever um (nesse caso o texto continua na planilha)."
    ),
}


@dataclass(frozen=True)
class Modelo:
    """Um tipo de ficha (Indicador de Compromisso ou Iniciativa)."""

    codigo: str  # "INDICADOR" | "INICIATIVA"
    rotulo: str  # nome legível, usado nos relatórios
    arquivo_saida: str  # "Indicadores.xlsx" | "Iniciativas.xlsx"
    #: Regex (sobre a chave canônica) que identificam o tipo do documento.
    deteccao: tuple[str, ...]
    campos: tuple[Campo, ...]

    def colunas_de_dados(self) -> list[str]:
        return [campo.coluna for campo in self.campos]

    def colunas(self) -> list[str]:
        """Todas as colunas da planilha, na ordem final."""
        return [
            *COLUNAS_AUDITORIA_INICIO,
            *self.colunas_de_dados(),
            *COLUNAS_AUDITORIA_FIM,
        ]

    def dicionario(self) -> list[dict[str, str]]:
        """Linhas da aba 'Dicionário' da planilha."""
        auditoria_inicio = [
            {"Coluna": nome, "Seção": "Auditoria", "Descrição": DESCRICOES_AUDITORIA[nome]}
            for nome in COLUNAS_AUDITORIA_INICIO
        ]
        dados = [
            {
                "Coluna": campo.coluna,
                "Seção": campo.secao or "Documento",
                "Descrição": campo.descricao,
            }
            for campo in self.campos
        ]
        auditoria_fim = [
            {"Coluna": nome, "Seção": "Auditoria", "Descrição": DESCRICOES_AUDITORIA[nome]}
            for nome in COLUNAS_AUDITORIA_FIM
        ]
        return [*auditoria_inicio, *dados, *auditoria_fim]
