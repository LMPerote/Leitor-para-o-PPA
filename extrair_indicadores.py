#!/usr/bin/env python3
"""Extrai fichas do PPA (.docx) para planilhas Excel consolidadas.

A pasta de entrada pode misturar os dois tipos de ficha; o tipo de cada
documento é identificado automaticamente e cada arquivo Word vira exatamente
UMA linha na planilha do seu tipo:

* "VÍNCULO DO INDICADOR DE COMPROMISSO" -> ``Indicadores.xlsx``
* "VÍNCULO DA INICIATIVA"               -> ``Iniciativas.xlsx``

Exemplos de uso:

    # Processa todos os .docx da pasta (inclusive subpastas)
    python extrair_indicadores.py -e ./documentos -s ./saida

    # Sem varrer subpastas, unindo listas com ponto e vírgula, e gerando CSV
    python extrair_indicadores.py -e ./documentos --sem-recursao --separador ponto-virgula --csv

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
from extrator.modelos import MODELOS, MODELOS_POR_CODIGO
from extrator.pipeline import Estatisticas
from extrator.planilha import salvar_csv, salvar_excel

try:  # A barra de progresso é opcional: sem tqdm, usa um contador simples.
    from tqdm import tqdm
except ImportError:  # pragma: no cover
    tqdm = None

logger = logging.getLogger("extrator")

SEPARADORES = {"quebra": "\n", "ponto-virgula": "; ", "barra": " | "}
LARGURA_RELATORIO = 66


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
    """Imprime tipo, seções, tabelas e células de um documento (diagnóstico)."""
    from extrator import ler_documento
    from extrator.parser import classificar, eh_rotulo

    documento = ler_documento(str(caminho))
    modelo = classificar(documento)
    print(f"# Documento: {caminho.name}")
    print(f"# Tipo detectado: {modelo.rotulo if modelo else 'NÃO RECONHECIDO'}\n")

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
        print(f"[{no.secao:<22}] {marca} {origem:<24} {texto}")
    return 0


def montar_argumentos() -> argparse.ArgumentParser:
    analisador = argparse.ArgumentParser(
        description=(
            "Consolida fichas do PPA (.docx) em Indicadores.xlsx e Iniciativas.xlsx, "
            "identificando o tipo de cada documento automaticamente."
        ),
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
        default=Path("."),
        help="Pasta onde gravar as planilhas (padrão: pasta atual).",
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
        "--csv", action="store_true", help="Gera também um .csv de cada planilha."
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
        tqdm(total=len(arquivos), unit="doc", desc="Extraindo", ncols=88) if tqdm else None
    )
    try:
        registros, estatisticas = processar_lote(
            arquivos, opcoes.entrada, separador=separador, barra_de_progresso=barra
        )
    finally:
        if barra:
            barra.close()

    # Uma planilha por tipo de ficha. São geradas mesmo quando vazias, para
    # que a saída do aplicativo seja sempre previsível.
    destinos: dict[str, Path] = {}
    for modelo in MODELOS:
        quadro = montar_dataframe(registros[modelo.codigo], modelo)
        destino = opcoes.saida / modelo.arquivo_saida
        salvar_excel(quadro, destino, modelo)
        destinos[modelo.codigo] = destino
        if opcoes.csv:
            salvar_csv(quadro, destino.with_suffix(".csv"))

    imprimir_relatorio(estatisticas, destinos, time.perf_counter() - inicio)
    return 0 if estatisticas.total_ignorados == 0 else 3


def imprimir_relatorio(
    estatisticas: Estatisticas, destinos: dict[str, Path], duracao: float
) -> None:
    """Resumo final do lote, por tipo de ficha."""
    regua = "=" * LARGURA_RELATORIO
    linhas = ["", regua, "RELATÓRIO DE EXTRAÇÃO", regua]
    linhas.append(f"Arquivos encontrados .......: {estatisticas.total}")

    for modelo in MODELOS:
        quantidade = estatisticas.processados.get(modelo.codigo, 0)
        pendentes = estatisticas.com_pendencias.get(modelo.codigo, 0)
        linhas.append(f"{quantidade} arquivos de {modelo.rotulo} processados.")
        if pendentes:
            linhas.append(f"  - {pendentes} com algum campo não localizado.")
        linhas.append(f"  - planilha: {destinos[modelo.codigo].resolve()}")

    linhas.append(f"{estatisticas.total_ignorados} arquivos ignorados ou com erro.")
    linhas += [
        f"  - {ocorrencia.arquivo}: {ocorrencia.motivo}"
        for ocorrencia in estatisticas.ignorados
    ]
    linhas.append(f"Tempo total ................: {duracao:.1f}s")

    for codigo, ausentes in estatisticas.campos_ausentes.items():
        if not ausentes:
            continue
        linhas.append("-" * LARGURA_RELATORIO)
        linhas.append(
            f"Campos mais ausentes em {MODELOS_POR_CODIGO[codigo].rotulo} "
            "(rótulo não localizado no documento):"
        )
        mais_ausentes = sorted(ausentes.items(), key=lambda item: -item[1])[:10]
        linhas += [f"  {contagem:>5}x  {coluna}" for coluna, contagem in mais_ausentes]

    linhas.append(regua)
    print("\n".join(linhas))


if __name__ == "__main__":
    raise SystemExit(main())
