#!/usr/bin/env python3
"""Extrai fichas de Indicadores de Compromisso (.docx) para uma planilha Excel.

Cada arquivo Word da pasta de entrada gera exatamente UMA linha na planilha.

Exemplos de uso:

    # Processa todos os .docx da pasta (inclusive subpastas)
    python extrair_indicadores.py -e ./documentos -s ./indicadores.xlsx

    # Sem varrer subpastas, unindo listas com ponto e vírgula, e gerando CSV
    python extrair_indicadores.py -e ./documentos --sem-recursao --separador ";" --csv

    # Inspeciona a estrutura de um arquivo (útil para ajustar o mapa de campos)
    python extrair_indicadores.py --inspecionar ./documentos/ficha.docx
"""

from __future__ import annotations

import argparse
import logging
import sys
import time
from pathlib import Path

from extrator import __version__, listar_documentos, montar_dataframe, processar_lote
from extrator.planilha import salvar_csv, salvar_excel

try:  # A barra de progresso é opcional: sem tqdm, usa um contador simples.
    from tqdm import tqdm
except ImportError:  # pragma: no cover
    tqdm = None

logger = logging.getLogger("extrator")

SEPARADORES = {"quebra": "\n", "ponto-virgula": "; ", "barra": " | "}


def configurar_log(arquivo_log: Path | None, verboso: bool) -> None:
    """Log no console e, opcionalmente, em arquivo."""
    formato = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%H:%M:%S")

    console = logging.StreamHandler(sys.stderr)
    console.setFormatter(formato)

    raiz = logging.getLogger()
    raiz.setLevel(logging.DEBUG if verboso else logging.INFO)
    raiz.handlers.clear()
    raiz.addHandler(console)

    if arquivo_log:
        arquivo_log.parent.mkdir(parents=True, exist_ok=True)
        em_arquivo = logging.FileHandler(arquivo_log, encoding="utf-8")
        em_arquivo.setFormatter(
            logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s")
        )
        em_arquivo.setLevel(logging.DEBUG)
        raiz.addHandler(em_arquivo)


def inspecionar(caminho: Path) -> int:
    """Imprime seções, tabelas e células de um documento (modo diagnóstico)."""
    from extrator import ler_documento
    from extrator.parser import eh_rotulo

    documento = ler_documento(str(caminho))
    print(f"# Documento: {caminho.name}\n")
    for no in documento.nos:
        if not no.texto:
            continue
        origem = (
            "parágrafo"
            if no.tipo == "paragrafo"
            else f"tabela {no.tabela} [{no.linha},{no.coluna}]"
        )
        marca = "RÓTULO" if eh_rotulo(no.texto) else "valor "
        texto = no.texto.replace("\n", " ⏎ ")
        texto = texto if len(texto) <= 110 else texto[:107] + "..."
        print(f"[{no.secao:<15}] {marca} {origem:<24} {texto}")
    return 0


def montar_argumentos() -> argparse.ArgumentParser:
    analisador = argparse.ArgumentParser(
        description="Consolida fichas de Indicadores de Compromisso (.docx) em um Excel.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    analisador.add_argument(
        "-e", "--entrada", type=Path, help="Pasta com os arquivos .docx."
    )
    analisador.add_argument(
        "-s",
        "--saida",
        type=Path,
        default=Path("indicadores_consolidados.xlsx"),
        help="Arquivo .xlsx de saída (padrão: indicadores_consolidados.xlsx).",
    )
    analisador.add_argument(
        "--sem-recursao",
        action="store_true",
        help="Não varre subpastas da pasta de entrada.",
    )
    analisador.add_argument(
        "--separador",
        default="quebra",
        help="Como unir itens de listas na mesma célula: "
        "'quebra' (padrão), 'ponto-virgula', 'barra' ou um texto literal.",
    )
    analisador.add_argument(
        "--csv", action="store_true", help="Gera também um .csv com os mesmos dados."
    )
    analisador.add_argument(
        "--limite",
        type=int,
        default=0,
        help="Processa apenas os N primeiros arquivos (útil para testes).",
    )
    analisador.add_argument(
        "--log", type=Path, default=None, help="Grava o log detalhado neste arquivo."
    )
    analisador.add_argument(
        "-v", "--verboso", action="store_true", help="Mostra mensagens de depuração."
    )
    analisador.add_argument(
        "--inspecionar",
        type=Path,
        metavar="ARQUIVO.docx",
        help="Diagnostica a estrutura de um único documento e encerra.",
    )
    analisador.add_argument("--versao", action="version", version=f"%(prog)s {__version__}")
    return analisador


def main(argumentos: list[str] | None = None) -> int:
    analisador = montar_argumentos()
    opcoes = analisador.parse_args(argumentos)
    configurar_log(opcoes.log, opcoes.verboso)

    if opcoes.inspecionar:
        if not opcoes.inspecionar.is_file():
            logger.error("Arquivo não encontrado: %s", opcoes.inspecionar)
            return 2
        return inspecionar(opcoes.inspecionar)

    if not opcoes.entrada:
        analisador.error("informe a pasta de entrada com -e/--entrada")
    if not opcoes.entrada.is_dir():
        logger.error("Pasta de entrada inválida: %s", opcoes.entrada)
        return 2

    arquivos = listar_documentos(opcoes.entrada, recursivo=not opcoes.sem_recursao)
    if opcoes.limite > 0:
        arquivos = arquivos[: opcoes.limite]
    if not arquivos:
        logger.error("Nenhum arquivo .docx encontrado em %s", opcoes.entrada)
        return 1

    separador = SEPARADORES.get(opcoes.separador, opcoes.separador)
    logger.info("Arquivos encontrados: %d", len(arquivos))

    inicio = time.perf_counter()
    barra = (
        tqdm(total=len(arquivos), unit="doc", desc="Extraindo", ncols=88)
        if tqdm
        else None
    )
    try:
        registros, estatisticas = processar_lote(
            arquivos, opcoes.entrada, separador=separador, barra_de_progresso=barra
        )
    finally:
        if barra:
            barra.close()

    quadro = montar_dataframe(registros)
    salvar_excel(quadro, opcoes.saida)
    if opcoes.csv:
        salvar_csv(quadro, opcoes.saida.with_suffix(".csv"))

    imprimir_relatorio(estatisticas, opcoes.saida, time.perf_counter() - inicio)
    return 0 if estatisticas.erros == 0 else 3


def imprimir_relatorio(estatisticas, saida: Path, duracao: float) -> None:
    """Resumo final do lote."""
    linhas = [
        "",
        "=" * 62,
        "RELATÓRIO DE EXTRAÇÃO",
        "=" * 62,
        f"Arquivos encontrados .......: {estatisticas.total}",
        f"Processados com sucesso ....: {estatisticas.sucesso}",
        f"  - completos ..............: {estatisticas.sucesso - estatisticas.com_pendencias}",
        f"  - com campos pendentes ...: {estatisticas.com_pendencias}",
        f"Falhas de leitura ..........: {estatisticas.erros}",
        f"Tempo total ................: {duracao:.1f}s",
        f"Planilha gerada ............: {saida.resolve()}",
    ]

    if estatisticas.campos_ausentes:
        linhas.append("-" * 62)
        linhas.append("Campos mais ausentes (rótulo não localizado no documento):")
        mais_ausentes = sorted(
            estatisticas.campos_ausentes.items(), key=lambda item: -item[1]
        )[:10]
        linhas += [f"  {contagem:>5}x  {coluna}" for coluna, contagem in mais_ausentes]

    if estatisticas.arquivos_com_erro:
        linhas.append("-" * 62)
        linhas.append("Arquivos com erro de leitura:")
        linhas += [f"  - {nome}" for nome in estatisticas.arquivos_com_erro[:20]]
        if len(estatisticas.arquivos_com_erro) > 20:
            linhas.append(f"  ... e mais {len(estatisticas.arquivos_com_erro) - 20}.")

    linhas.append("=" * 62)
    print("\n".join(linhas))


if __name__ == "__main__":
    raise SystemExit(main())
