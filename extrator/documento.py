"""Leitura de arquivos .docx e construção de um modelo navegável.

O modelo produzido tem duas visões complementares:

* ``Documento.nos`` – sequência linear (na ordem de leitura) de todos os
  parágrafos e células do arquivo, com a seção a que cada um pertence.
  É o que permite localizar um rótulo em qualquer lugar do documento.
* ``Documento.tabelas`` – grade de cada tabela com tratamento de células
  mescladas, o que permite perguntar "qual é a célula à direita/abaixo
  deste rótulo?" — a operação central do parser.
"""

from __future__ import annotations

from dataclasses import dataclass, field

import docx
from docx.document import Document as _DocumentoDocx
from docx.oxml.ns import qn
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph

from .texto import chave, limpar

# Cabeçalhos que delimitam as grandes seções da ficha de indicador.
# A comparação é feita pela chave canônica, então variações como
# "ATRIBUTOSDO INDICADOR" (sem espaço) continuam sendo reconhecidas.
SECOES: dict[str, str] = {
    "vinculodoindicadordecompromisso": "VINCULO",
    "atributosdoindicadordecompromisso": "ATRIBUTOS",
    "desagregacaoterritorial": "TERRITORIAL",
    "informacoescomplementares": "COMPLEMENTARES",
}
SECAO_INICIAL = "CABECALHO"


_NS_COMPAT = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"

#: Subárvores que não fazem parte do texto corrente do parágrafo:
#: caixas de texto/imagens (carimbos decorativos como "IC - 2027"), códigos de
#: campo, texto excluído por controle de alterações e o ramo de compatibilidade
#: do ``mc:AlternateContent`` (que duplicaria o conteúdo já lido em mc:Choice).
_SUBARVORES_IGNORADAS = frozenset(
    {
        qn("w:pict"),
        qn("w:drawing"),
        qn("w:instrText"),
        qn("w:del"),
        f"{_NS_COMPAT}Fallback",
    }
)


def _coletar_texto(elemento, partes: list[str]) -> None:
    for filho in elemento:
        etiqueta = filho.tag
        if etiqueta in _SUBARVORES_IGNORADAS:
            continue
        if etiqueta == qn("w:t"):
            partes.append(filho.text or "")
        elif etiqueta in (qn("w:br"), qn("w:cr")):
            partes.append("\n")
        elif etiqueta == qn("w:tab"):
            partes.append(" ")
        else:
            _coletar_texto(filho, partes)


def _texto_do_paragrafo(paragrafo: Paragraph) -> str:
    """Extrai o texto de um parágrafo preservando quebras manuais.

    ``Paragraph.text`` do python-docx ignora ``<w:br/>``, ``<w:tab/>`` e
    ``<w:cr/>``; em fichas preenchidas à mão essas quebras costumam separar
    itens de listas (ex.: causas críticas), então precisam ser preservadas.
    """
    partes: list[str] = []
    _coletar_texto(paragrafo._p, partes)
    return "".join(partes)


def _blocos(container) -> list[Paragraph | Table]:
    """Itera parágrafos e tabelas de um documento/célula na ordem de leitura."""
    if isinstance(container, _DocumentoDocx):
        elemento_pai = container.element.body
    elif isinstance(container, _Cell):
        elemento_pai = container._tc
    else:  # pragma: no cover - defensivo
        raise TypeError(f"Container não suportado: {type(container)!r}")

    resultado: list[Paragraph | Table] = []
    for filho in elemento_pai.iterchildren():
        if isinstance(filho, CT_P):
            resultado.append(Paragraph(filho, container))
        elif isinstance(filho, CT_Tbl):
            resultado.append(Table(filho, container))
    return resultado


def _texto_da_celula(celula: _Cell) -> str:
    """Texto de uma célula (apenas seus parágrafos diretos).

    Tabelas aninhadas são ignoradas aqui de propósito: elas entram no modelo
    como tabelas próprias, com sua própria grade de vizinhança.
    """
    linhas = [_texto_do_paragrafo(p) for p in _blocos(celula) if isinstance(p, Paragraph)]
    return limpar("\n".join(linhas))


@dataclass
class No:
    """Um parágrafo ou uma célula, posicionado no documento."""

    tipo: str  # "paragrafo" ou "celula"
    texto: str
    secao: str
    ordem: int
    tabela: int = -1
    linha: int = -1
    coluna: int = -1

    @property
    def chave(self) -> str:
        return chave(self.texto)


@dataclass
class Tabela:
    """Grade de uma tabela com tratamento de células mescladas.

    ``ids[l][c]`` guarda a identidade do elemento XML da célula. Células
    mescladas (horizontal ou verticalmente) repetem o mesmo id em várias
    posições — é assim que descobrimos onde uma célula realmente termina.

    ``ancoras`` mantém referências vivas aos elementos ``<w:tc>``: o lxml
    recria seus objetos-proxy sob demanda, e sem essa âncora o ``id()`` de um
    elemento coletado poderia ser reaproveitado por outro, fazendo células
    distintas parecerem mescladas.
    """

    indice: int
    ids: list[list[int]] = field(default_factory=list)
    textos: dict[int, str] = field(default_factory=dict)
    ancoras: list = field(default_factory=list, repr=False)

    def texto(self, linha: int, coluna: int) -> str:
        try:
            return self.textos.get(self.ids[linha][coluna], "")
        except IndexError:
            return ""

    def _id(self, linha: int, coluna: int) -> int | None:
        try:
            return self.ids[linha][coluna]
        except IndexError:
            return None

    def celulas_a_direita(self, linha: int, coluna: int) -> list[str]:
        """Textos das células distintas à direita, na mesma linha."""
        atual = self._id(linha, coluna)
        vistos = {atual}
        resultado: list[str] = []
        for c in range(coluna + 1, len(self.ids[linha])):
            identificador = self.ids[linha][c]
            if identificador in vistos:
                continue
            vistos.add(identificador)
            resultado.append(self.textos.get(identificador, ""))
        return resultado

    def celulas_abaixo(self, linha: int, coluna: int) -> list[str]:
        """Textos das células distintas abaixo, na mesma coluna."""
        atual = self._id(linha, coluna)
        vistos = {atual}
        resultado: list[str] = []
        for l in range(linha + 1, len(self.ids)):
            identificador = self._id(l, coluna)
            if identificador is None or identificador in vistos:
                continue
            vistos.add(identificador)
            resultado.append(self.textos.get(identificador, ""))
        return resultado

    def linha_de_celulas(self, linha: int) -> list[tuple[int, str]]:
        """Pares (coluna, texto) das células distintas de uma linha."""
        vistos: set[int] = set()
        resultado: list[tuple[int, str]] = []
        for c, identificador in enumerate(self.ids[linha] if linha < len(self.ids) else []):
            if identificador in vistos:
                continue
            vistos.add(identificador)
            resultado.append((c, self.textos.get(identificador, "")))
        return resultado

    @property
    def total_linhas(self) -> int:
        return len(self.ids)


@dataclass
class Documento:
    """Representação completa de um arquivo .docx já lido."""

    caminho: str
    nos: list[No] = field(default_factory=list)
    tabelas: list[Tabela] = field(default_factory=list)

    def nos_da_secao(self, secao: str | None) -> list[No]:
        if secao is None:
            return self.nos
        return [no for no in self.nos if no.secao == secao]

    def texto_integral(self) -> str:
        return "\n".join(no.texto for no in self.nos if no.texto)


class _Leitor:
    """Percorre o documento montando nós e grades, controlando a seção atual."""

    def __init__(self, caminho: str) -> None:
        self.documento = Documento(caminho=caminho)
        self.secao = SECAO_INICIAL
        self.ordem = 0

    def ler(self) -> Documento:
        arquivo = docx.Document(self.documento.caminho)
        self._percorrer(_blocos(arquivo))
        return self.documento

    # -- internos ---------------------------------------------------------
    def _atualizar_secao(self, texto: str) -> None:
        nova = SECOES.get(chave(texto))
        if nova:
            self.secao = nova

    def _novo_no(self, **kwargs) -> No:
        no = No(secao=self.secao, ordem=self.ordem, **kwargs)
        self.ordem += 1
        self.documento.nos.append(no)
        return no

    def _percorrer(self, blocos) -> None:
        for bloco in blocos:
            if isinstance(bloco, Paragraph):
                texto = limpar(_texto_do_paragrafo(bloco))
                if not texto:
                    continue
                self._atualizar_secao(texto)
                self._novo_no(tipo="paragrafo", texto=texto)
            else:
                self._ler_tabela(bloco)

    def _ler_tabela(self, tabela: Table) -> None:
        indice = len(self.documento.tabelas)
        grade = Tabela(indice=indice)
        self.documento.tabelas.append(grade)

        for numero_linha, linha in enumerate(tabela.rows):
            try:
                celulas = list(linha.cells)
            except (IndexError, ValueError):
                # Linha malformada (acontece em tabelas editadas manualmente):
                # registra vazia para não deslocar os índices das demais.
                grade.ids.append([])
                continue

            ids_linha: list[int] = []
            aninhadas: list[Table] = []
            for numero_coluna, celula in enumerate(celulas):
                elemento = celula._tc
                grade.ancoras.append(elemento)  # ver docstring de Tabela
                identificador = id(elemento)
                ids_linha.append(identificador)
                if identificador in grade.textos:
                    continue  # célula mesclada já processada

                texto = _texto_da_celula(celula)
                grade.textos[identificador] = texto
                if texto:
                    self._atualizar_secao(texto)
                self._novo_no(
                    tipo="celula",
                    texto=texto,
                    tabela=indice,
                    linha=numero_linha,
                    coluna=numero_coluna,
                )
                aninhadas.extend(
                    bloco for bloco in _blocos(celula) if isinstance(bloco, Table)
                )

            grade.ids.append(ids_linha)
            # Tabelas aninhadas viram tabelas independentes, logo após a linha.
            for aninhada in aninhadas:
                self._ler_tabela(aninhada)


def ler_documento(caminho: str) -> Documento:
    """Lê um arquivo .docx e devolve o modelo navegável."""
    return _Leitor(caminho).ler()
