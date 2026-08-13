"""Fluxo completo da extração: pasta de entrada -> planilhas gravadas.

Este módulo é o "miolo" do aplicativo, compartilhado pela linha de comando
(:mod:`extrair_indicadores`) e pela janela gráfica (:mod:`interface`), para
que as duas façam exatamente a mesma coisa.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass
from pathlib import Path

from .modelos import MODELOS
from .pipeline import Estatisticas, listar_documentos, processar_lote
from .planilha import montar_dataframe, salvar_csv, salvar_excel

logger = logging.getLogger(__name__)

#: Formas de unir itens de lista dentro de uma mesma célula.
SEPARADORES: dict[str, str] = {
    "quebra": "\n",
    "ponto-virgula": "; ",
    "barra": " | ",
}


class PastaSemDocumentos(Exception):
    """Nenhum .docx encontrado na pasta de entrada."""


@dataclass(frozen=True)
class Opcoes:
    """Parâmetros de uma execução."""

    entrada: Path
    saida: Path
    recursivo: bool = True
    separador: str = "\n"
    gerar_csv: bool = False
    limite: int = 0


@dataclass
class Execucao:
    """O que uma execução produziu."""

    arquivos: list[Path]
    estatisticas: Estatisticas
    destinos: dict[str, Path]


def resolver_separador(nome: str) -> str:
    """Aceita um apelido ('ponto-virgula') ou o próprio texto separador."""
    return SEPARADORES.get(nome, nome)


def executar(opcoes: Opcoes, criar_progresso=None) -> Execucao:
    """Processa a pasta e grava uma planilha por tipo de ficha.

    ``criar_progresso`` recebe o total de arquivos e devolve um objeto com
    ``update(n)`` (a barra do tqdm, no terminal, ou a da janela gráfica).
    """
    arquivos = listar_documentos(opcoes.entrada, recursivo=opcoes.recursivo)
    if opcoes.limite > 0:
        arquivos = arquivos[: opcoes.limite]
    if not arquivos:
        raise PastaSemDocumentos(
            f"Nenhum arquivo .docx encontrado em {opcoes.entrada}"
        )

    logger.info("Arquivos encontrados: %d", len(arquivos))
    progresso = criar_progresso(len(arquivos)) if criar_progresso else None
    try:
        registros, estatisticas = processar_lote(
            arquivos,
            opcoes.entrada,
            separador=opcoes.separador,
            barra_de_progresso=progresso,
        )
    finally:
        fechar = getattr(progresso, "close", None)
        if fechar:
            fechar()

    # Uma planilha por tipo de ficha. São geradas mesmo quando vazias, para
    # que a saída do aplicativo seja sempre previsível.
    destinos: dict[str, Path] = {}
    for modelo in MODELOS:
        quadro = montar_dataframe(registros[modelo.codigo], modelo)
        destino = opcoes.saida / modelo.arquivo_saida
        salvar_excel(quadro, destino, modelo)
        destinos[modelo.codigo] = destino.resolve()
        if opcoes.gerar_csv:
            salvar_csv(quadro, destino.with_suffix(".csv"))

    return Execucao(arquivos=arquivos, estatisticas=estatisticas, destinos=destinos)
