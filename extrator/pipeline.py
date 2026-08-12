"""Orquestração do processamento em lote (pasta -> registros por tipo de ficha)."""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from pathlib import Path

from .documento import ler_documento
from .modelos import COLUNA_ARQUIVO, MODELOS, Modelo
from .parser import Extrator, classificar

logger = logging.getLogger(__name__)

EXTENSOES = (".docx",)
#: Arquivos temporários do Word ("~$relatorio.docx") nunca são documentos reais.
PREFIXO_TEMPORARIO = "~$"


@dataclass
class Ocorrencia:
    """Um arquivo que não entrou em nenhuma planilha, e o porquê."""

    arquivo: str
    motivo: str


@dataclass
class Estatisticas:
    """Contadores do lote, usados no relatório final."""

    total: int = 0
    #: código do modelo -> quantidade de arquivos processados.
    processados: dict[str, int] = field(default_factory=dict)
    #: código do modelo -> quantidade com algum campo não localizado.
    com_pendencias: dict[str, int] = field(default_factory=dict)
    #: arquivos ignorados (tipo não reconhecido) ou com erro de leitura.
    ignorados: list[Ocorrencia] = field(default_factory=list)
    #: código do modelo -> coluna -> nº de arquivos em que o rótulo faltou.
    campos_ausentes: dict[str, dict[str, int]] = field(default_factory=dict)

    @property
    def total_ignorados(self) -> int:
        return len(self.ignorados)


def listar_documentos(pasta: Path, recursivo: bool = True) -> list[Path]:
    """Lista os .docx da pasta, ignorando temporários do Word."""
    padrao = "**/*" if recursivo else "*"
    return [
        caminho
        for caminho in sorted(pasta.glob(padrao))
        if caminho.is_file()
        and caminho.suffix.lower() in EXTENSOES
        and not caminho.name.startswith(PREFIXO_TEMPORARIO)
    ]


def processar_arquivo(
    caminho: Path, extratores: dict[str, Extrator], pasta_base: Path
) -> tuple[Modelo | None, dict[str, str] | None, str | None]:
    """Lê, classifica e extrai um documento.

    Devolve ``(modelo, registro, motivo_da_falha)``. Erro de leitura e tipo não
    reconhecido nunca interrompem o lote: devolvem ``motivo`` preenchido, para
    o arquivo ser listado no relatório final.
    """
    try:
        documento = ler_documento(str(caminho))
    except Exception as erro:  # noqa: BLE001 - resiliência é requisito
        logger.warning("Falha ao ler '%s': %s", caminho.name, erro)
        return None, None, f"erro de leitura ({type(erro).__name__}: {erro})"

    modelo = classificar(documento)
    if modelo is None:
        logger.warning("Tipo não reconhecido em '%s'", caminho.name)
        return None, None, "tipo não reconhecido (sem seção de vínculo conhecida)"

    try:
        resultado = extratores[modelo.codigo].extrair(documento)
    except Exception as erro:  # noqa: BLE001 - resiliência é requisito
        logger.warning("Falha ao extrair '%s': %s", caminho.name, erro)
        return None, None, f"erro de extração ({type(erro).__name__}: {erro})"

    try:
        relativo = str(caminho.relative_to(pasta_base))
    except ValueError:  # pragma: no cover - caminho fora da base
        relativo = str(caminho)

    registro: dict[str, str] = {
        COLUNA_ARQUIVO: caminho.name,
        "Caminho_Relativo": relativo,
        **resultado.valores,
    }
    if resultado.nao_encontrados:
        registro["Status"] = "OK_COM_PENDENCIAS"
        registro["Campos_Nao_Encontrados"] = "; ".join(resultado.nao_encontrados)
        registro["Observacoes"] = (
            f"{len(resultado.nao_encontrados)} campo(s) não localizado(s) no documento."
        )
        logger.debug(
            "'%s' (%s): campos não encontrados -> %s",
            caminho.name,
            modelo.rotulo,
            ", ".join(resultado.nao_encontrados),
        )
    else:
        registro["Status"] = "OK"
        registro["Campos_Nao_Encontrados"] = ""
        registro["Observacoes"] = ""
    return modelo, registro, None


def processar_lote(
    arquivos: list[Path],
    pasta_base: Path,
    separador: str = "\n",
    barra_de_progresso=None,
) -> tuple[dict[str, list[dict[str, str]]], Estatisticas]:
    """Processa sequencialmente todos os arquivos, separando-os por tipo."""
    extratores = {
        modelo.codigo: Extrator(modelo, separador=separador) for modelo in MODELOS
    }
    registros: dict[str, list[dict[str, str]]] = {modelo.codigo: [] for modelo in MODELOS}
    estatisticas = Estatisticas(total=len(arquivos))

    for caminho in arquivos:
        modelo, registro, motivo = processar_arquivo(caminho, extratores, pasta_base)

        if modelo is None or registro is None:
            estatisticas.ignorados.append(
                Ocorrencia(caminho.name, motivo or "motivo desconhecido")
            )
        else:
            registros[modelo.codigo].append(registro)
            estatisticas.processados[modelo.codigo] = (
                estatisticas.processados.get(modelo.codigo, 0) + 1
            )
            if registro["Status"] == "OK_COM_PENDENCIAS":
                estatisticas.com_pendencias[modelo.codigo] = (
                    estatisticas.com_pendencias.get(modelo.codigo, 0) + 1
                )
            ausentes = estatisticas.campos_ausentes.setdefault(modelo.codigo, {})
            for coluna in filter(None, registro["Campos_Nao_Encontrados"].split("; ")):
                ausentes[coluna] = ausentes.get(coluna, 0) + 1

        if barra_de_progresso is not None:
            barra_de_progresso.update(1)

    return registros, estatisticas
