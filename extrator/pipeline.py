"""Orquestração do processamento em lote (pasta -> registros)."""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from pathlib import Path

from .campos import COLUNA_ARQUIVO, colunas_de_dados
from .documento import ler_documento
from .parser import Extrator

logger = logging.getLogger(__name__)

EXTENSOES = (".docx",)
#: Arquivos temporários do Word ("~$relatorio.docx") nunca são documentos reais.
PREFIXO_TEMPORARIO = "~$"


@dataclass
class Estatisticas:
    """Contadores do lote, usados no relatório final."""

    total: int = 0
    sucesso: int = 0
    com_pendencias: int = 0
    erros: int = 0
    arquivos_com_erro: list[str] = field(default_factory=list)
    campos_ausentes: dict[str, int] = field(default_factory=dict)


def listar_documentos(pasta: Path, recursivo: bool = True) -> list[Path]:
    """Lista os .docx da pasta, ignorando temporários do Word."""
    padrao = "**/*" if recursivo else "*"
    arquivos = [
        caminho
        for caminho in sorted(pasta.glob(padrao))
        if caminho.is_file()
        and caminho.suffix.lower() in EXTENSOES
        and not caminho.name.startswith(PREFIXO_TEMPORARIO)
    ]
    return arquivos


def processar_arquivo(
    caminho: Path, extrator: Extrator, pasta_base: Path
) -> tuple[dict[str, str], bool]:
    """Extrai um documento. Devolve (registro, houve_erro).

    Erros de leitura nunca interrompem o lote: o arquivo entra na planilha
    com ``Status = ERRO_DE_LEITURA`` e a mensagem em ``Observacoes``.
    """
    try:
        relativo = str(caminho.relative_to(pasta_base))
    except ValueError:  # pragma: no cover - caminho fora da base
        relativo = str(caminho)

    registro: dict[str, str] = {
        COLUNA_ARQUIVO: caminho.name,
        "Caminho_Relativo": relativo,
    }

    try:
        documento = ler_documento(str(caminho))
        resultado = extrator.extrair(documento)
    except Exception as erro:  # noqa: BLE001 - resiliência é requisito
        logger.warning("Falha ao processar '%s': %s", caminho.name, erro)
        registro.update({coluna: "" for coluna in colunas_de_dados()})
        registro["Status"] = "ERRO_DE_LEITURA"
        registro["Campos_Nao_Encontrados"] = ""
        registro["Observacoes"] = f"{type(erro).__name__}: {erro}"
        return registro, True

    registro.update(resultado.valores)
    if resultado.nao_encontrados:
        registro["Status"] = "OK_COM_PENDENCIAS"
        registro["Campos_Nao_Encontrados"] = "; ".join(resultado.nao_encontrados)
        registro["Observacoes"] = (
            f"{len(resultado.nao_encontrados)} campo(s) não localizado(s) no documento."
        )
        logger.debug(
            "'%s': campos não encontrados -> %s",
            caminho.name,
            ", ".join(resultado.nao_encontrados),
        )
    else:
        registro["Status"] = "OK"
        registro["Campos_Nao_Encontrados"] = ""
        registro["Observacoes"] = ""
    return registro, False


def processar_lote(
    arquivos: list[Path],
    pasta_base: Path,
    separador: str = "\n",
    barra_de_progresso=None,
) -> tuple[list[dict[str, str]], Estatisticas]:
    """Processa sequencialmente todos os arquivos informados."""
    extrator = Extrator(separador=separador)
    estatisticas = Estatisticas(total=len(arquivos))
    registros: list[dict[str, str]] = []

    for caminho in arquivos:
        registro, houve_erro = processar_arquivo(caminho, extrator, pasta_base)
        registros.append(registro)

        if houve_erro:
            estatisticas.erros += 1
            estatisticas.arquivos_com_erro.append(caminho.name)
        elif registro["Status"] == "OK_COM_PENDENCIAS":
            estatisticas.com_pendencias += 1
            estatisticas.sucesso += 1
        else:
            estatisticas.sucesso += 1

        for coluna in filter(None, registro["Campos_Nao_Encontrados"].split("; ")):
            estatisticas.campos_ausentes[coluna] = (
                estatisticas.campos_ausentes.get(coluna, 0) + 1
            )

        if barra_de_progresso is not None:
            barra_de_progresso.update(1)

    return registros, estatisticas
