"""Orquestração do processamento em lote (pasta -> registros por tipo de ficha)."""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from pathlib import Path

from .documento import ler_documento
from .modelos import COLUNA_ARQUIVO, MODELOS, Modelo
from .parser import Extrator, classificar

logger = logging.getLogger(__name__)

EXTENSOES = (".docx",)
#: Arquivos temporários do Word ("~$relatorio.docx") nunca são documentos reais.
PREFIXO_TEMPORARIO = "~$"
#: Chave auxiliar do registro, fora das colunas do modelo (não vai à planilha).
CHAVE_ROTULOS_AUSENTES = "_Rotulos_Ausentes"


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
    #: código do modelo -> linhas geradas na planilha. Difere de
    #: ``processados`` quando a ficha traz mais de um problema, causa crítica,
    #: ação crítica ou entrega vinculada.
    linhas: dict[str, int] = field(default_factory=dict)
    #: código do modelo -> quantidade com algum campo não localizado.
    com_pendencias: dict[str, int] = field(default_factory=dict)
    #: código do modelo -> quantidade de arquivos com algum item que remete a
    #: outro item em vez de descrever um (ver :func:`itens_sem_descricao`).
    com_itens_sem_descricao: dict[str, int] = field(default_factory=dict)
    #: arquivos ignorados (tipo não reconhecido) ou com erro de leitura.
    ignorados: list[Ocorrencia] = field(default_factory=list)
    #: código do modelo -> coluna -> nº de arquivos em que o campo ficou sem
    #: valor (rótulo ausente ou presente sem resposta).
    campos_ausentes: dict[str, dict[str, int]] = field(default_factory=dict)
    #: código do modelo -> coluna -> nº de arquivos em que o **rótulo** sequer
    #: existe. O restante de ``campos_ausentes`` é ficha sem preencher: a
    #: distinção separa "o modelo desta ficha é outro" de "não preencheram".
    rotulos_ausentes: dict[str, dict[str, int]] = field(default_factory=dict)

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
    caminho: Path,
    extratores: dict[str, Extrator],
    pasta_base: Path,
    separador: str = "\n",
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
    observacoes: list[str] = []
    # Chave auxiliar: não é coluna do modelo, então não entra na planilha —
    # serve para o lote contabilizar rótulo ausente à parte de ficha vazia.
    registro[CHAVE_ROTULOS_AUSENTES] = "; ".join(resultado.rotulos_ausentes)
    if resultado.nao_encontrados:
        registro["Status"] = "OK_COM_PENDENCIAS"
        registro["Campos_Nao_Encontrados"] = "; ".join(resultado.nao_encontrados)
        observacoes.append(
            f"{len(resultado.nao_encontrados)} campo(s) sem valor "
            "(rótulo não localizado ou sem resposta preenchida)."
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

    # Aviso (não altera nenhum dado): itens que remetem a outro item em vez de
    # descrever um. O texto continua na planilha como está na ficha.
    achados = itens_sem_descricao(registro, separador)
    if achados:
        observacoes.append(_aviso_de_itens(achados))
        logger.debug(
            "'%s' (%s): itens sem descrição -> %s",
            caminho.name,
            modelo.rotulo,
            ", ".join(f"{coluna}={item!r}" for coluna, item in achados),
        )

    registro["Observacoes"] = " ".join(observacoes)
    return modelo, registro, None


#: Colunas que definem a granularidade da planilha: uma linha por problema,
#: causa crítica, ação crítica e entrega que se correspondem dentro da mesma
#: ficha.
COLUNA_PROBLEMA = "Problemas_Vinculados"
COLUNA_CAUSA = "Causas_Criticas"
COLUNA_ACAO = "Acoes_Criticas"
COLUNA_ENTREGA = "Entregas_Vinculadas"
COLUNAS_EXPANDIDAS: tuple[str, ...] = (
    COLUNA_PROBLEMA,
    COLUNA_CAUSA,
    COLUNA_ACAO,
    COLUNA_ENTREGA,
)

#: Marcas que separam itens dentro de uma mesma célula, além do separador
#: escolhido na execução: a quebra de linha (como a ficha lista os itens no
#: Word) e o ponto e vírgula (como as listas chegam digitadas em uma linha só).
MARCAS_DE_ITEM: tuple[str, ...] = ("\n", ";")


def dividir_em_itens(valor: str, separador: str) -> list[str]:
    """Quebra a célula em itens individuais, um por linha da planilha.

    Divide pelo separador configurado e também por quebra de linha e ponto e
    vírgula, porque a mesma lista chega das duas formas nas fichas: itens em
    linhas separadas da célula do Word ("Entrega A⏎Entrega B") ou digitados
    seguidos ("P1 ...; P2 ...;"). Itens repetidos são preservados: cada
    ocorrência vale uma linha.
    """
    if not valor:
        return []

    texto = valor
    for marca in dict.fromkeys((separador, *MARCAS_DE_ITEM)):
        if marca and marca != "\n":
            texto = texto.replace(marca, "\n")
    return [parte.strip() for parte in texto.split("\n") if parte.strip()]


#: Uma "palavra de conteúdo": três letras seguidas. Um item que não tem
#: nenhuma ("AC 4,5,7,810,12,13,14", "C5P1,3CC12,13,16AC3", "P1") remete a
#: outro item em vez de descrever um: é anotação de quem preencheu a ficha.
_REGEX_PALAVRA = re.compile(r"[^\W\d_]{3,}")

#: Quantos itens citar no aviso antes de resumir o restante.
MAXIMO_ITENS_CITADOS = 3
#: Tamanho máximo de um item citado no aviso.
LIMITE_ITEM_CITADO = 60


def itens_sem_descricao(registro: dict[str, str], separador: str) -> list[tuple[str, str]]:
    """Pares ``(coluna, item)`` que remetem a outro item em vez de descrever.

    Serve só para **avisar**: nada é removido da planilha. O texto está
    digitado na ficha e vai para a linha dele como está; o aviso é para
    localizar as fichas que precisam de correção no documento.
    """
    return [
        (coluna, item)
        for coluna in COLUNAS_EXPANDIDAS
        for item in dividir_em_itens(registro.get(coluna, ""), separador)
        if not _REGEX_PALAVRA.search(item)
    ]


def _aviso_de_itens(achados: list[tuple[str, str]]) -> str:
    """Monta o aviso de ``Observacoes`` a partir dos itens sem descrição."""
    por_coluna: dict[str, list[str]] = {}
    for coluna, item in achados:
        por_coluna.setdefault(coluna, []).append(item)

    avisos = []
    for coluna, itens in por_coluna.items():
        citados = [
            item if len(item) <= LIMITE_ITEM_CITADO else item[:LIMITE_ITEM_CITADO] + "..."
            for item in itens[:MAXIMO_ITENS_CITADOS]
        ]
        amostra = ", ".join(f'"{item}"' for item in citados)
        if len(itens) > MAXIMO_ITENS_CITADOS:
            amostra += f" e mais {len(itens) - MAXIMO_ITENS_CITADOS}"
        plural = "itens" if len(itens) > 1 else "item"
        avisos.append(
            f"{coluna}: {len(itens)} {plural} sem descrição ({amostra}) — conferir a ficha."
        )
    return " ".join(avisos)


def _item_na_posicao(itens: list[str], posicao: int) -> str:
    """Item da posição pedida; passado o fim da lista, o último se repete.

    É como a ficha se comporta: um problema que vale para várias causas críticas
    aparece repetido nas linhas dessas causas. Confirmado contra a planilha de
    vínculo já usada pela equipe — uma ficha com 4 problemas, 1 causa e 3 ações
    vira P1/C1/A1, P2/C1/A2, P3/C1/A3 e P4/C1/A3.
    """
    if not itens:
        return ""
    if posicao < len(itens):
        return itens[posicao]
    return itens[-1]


def expandir_em_linhas(
    registro: dict[str, str], separador: str
) -> list[dict[str, str]]:
    """Desdobra o registro em uma linha por problema, causa, ação e entrega.

    A regra de negócio é "um por linha": cada linha traz um problema, uma causa
    crítica, uma ação crítica e uma entrega vinculada, pareados pela ordem em
    que aparecem na ficha. Quando uma das listas é menor que as outras, seu
    último item se repete nas linhas restantes — nada é agrupado nem
    deduplicado, e o mesmo item aparece em quantas linhas for preciso. As demais
    colunas se repetem em todas as linhas da ficha.

    Ficha com um único item de cada (ou nenhum) continua gerando uma linha só,
    para que todo arquivo apareça na planilha.
    """
    listas = {
        coluna: dividir_em_itens(registro.get(coluna, ""), separador)
        for coluna in COLUNAS_EXPANDIDAS
    }
    total = max(len(itens) for itens in listas.values())
    if total == 0:
        # Ficha sem nenhum desses itens continua aparecendo na planilha.
        return [registro]

    return [
        {
            **registro,
            **{
                coluna: _item_na_posicao(listas[coluna], posicao)
                for coluna in COLUNAS_EXPANDIDAS
                # Coluna ausente no modelo (a ficha de Indicador não tem ações
                # críticas) não é criada do nada.
                if coluna in registro
            },
        }
        for posicao in range(total)
    ]


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
        modelo, registro, motivo = processar_arquivo(
            caminho, extratores, pasta_base, separador
        )

        if modelo is None or registro is None:
            estatisticas.ignorados.append(
                Ocorrencia(caminho.name, motivo or "motivo desconhecido")
            )
        else:
            linhas = expandir_em_linhas(registro, separador)
            registros[modelo.codigo].extend(linhas)
            estatisticas.processados[modelo.codigo] = (
                estatisticas.processados.get(modelo.codigo, 0) + 1
            )
            estatisticas.linhas[modelo.codigo] = (
                estatisticas.linhas.get(modelo.codigo, 0) + len(linhas)
            )
            if registro["Status"] == "OK_COM_PENDENCIAS":
                estatisticas.com_pendencias[modelo.codigo] = (
                    estatisticas.com_pendencias.get(modelo.codigo, 0) + 1
                )
            if "sem descrição" in registro["Observacoes"]:
                estatisticas.com_itens_sem_descricao[modelo.codigo] = (
                    estatisticas.com_itens_sem_descricao.get(modelo.codigo, 0) + 1
                )
            ausentes = estatisticas.campos_ausentes.setdefault(modelo.codigo, {})
            for coluna in filter(None, registro["Campos_Nao_Encontrados"].split("; ")):
                ausentes[coluna] = ausentes.get(coluna, 0) + 1
            sem_rotulo = estatisticas.rotulos_ausentes.setdefault(modelo.codigo, {})
            for coluna in filter(None, registro[CHAVE_ROTULOS_AUSENTES].split("; ")):
                sem_rotulo[coluna] = sem_rotulo.get(coluna, 0) + 1

        if barra_de_progresso is not None:
            barra_de_progresso.update(1)

    return registros, estatisticas
