"""Motor de extração: classifica o documento e o transforma em uma linha.

Estratégia geral (independente da posição exata das tabelas no arquivo):

1. **Classificação** — :func:`classificar` procura, do topo para o fim, o
   cabeçalho de vínculo que identifica o tipo da ficha ("VÍNCULO DO INDICADOR
   DE COMPROMISSO" ou "VÍNCULO DA INICIATIVA") e devolve o modelo
   correspondente.
2. **Localização do rótulo** — a célula/parágrafo cujo texto casa com o
   rótulo do campo, preferindo ocorrências dentro da seção esperada.
3. **Resolução do valor**, nesta ordem:
   a) valor na própria célula depois de ``:`` (ex.: "Fonte: SEI");
   b) primeira célula **à direita** que não esteja vazia e que não seja,
      ela mesma, outro rótulo conhecido (layout horizontal);
   c) primeira célula **abaixo**, mesmo critério (layout vertical).
4. Campos marcados como ``multiplo`` continuam descendo pela coluna e unem
   todos os itens em um único texto — garantindo "uma célula por campo" no
   Excel mesmo quando o Word traz uma lista.

Layouts que fogem desse padrão (caixas de seleção, tabela de territórios, de
propostas de escuta social, de recursos e de ações orçamentárias) têm
extratores dedicados no fim do módulo.
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field

from .documento import Documento, No, Tabela
from .modelos import MODELOS, ROTULOS_CONHECIDOS, Campo, Modelo
from .texto import casa, chave, chaves, esta_marcado, juntar, limpar

logger = logging.getLogger(__name__)

SEPARADOR_PADRAO = "\n"
#: Separador entre as colunas de uma mesma linha de tabela consolidada.
SEPARADOR_DE_COLUNAS = " | "

_REGEX_ROTULOS = re.compile("|".join(f"(?:{p})" for p in ROTULOS_CONHECIDOS))

#: Formas de um texto anunciar um rótulo, da mais literal à mais tolerante.
EXATO, EMBUTIDO, PRIMEIRA_LINHA = 0, 1, 2

# Divisor "rótulo: valor" dentro de uma mesma célula/parágrafo (uma linha).
_DIVISOR_INLINE = re.compile(r"^([^:\n]{2,80}?)\s*[:：]\s*(.+?)\s*(?:\n|$)")


def eh_rotulo(texto: str) -> bool:
    """True se o texto é (apenas) um rótulo conhecido, e não um valor."""
    return casa(_REGEX_ROTULOS, texto)


def classificar(documento: Documento) -> Modelo | None:
    """Descobre o tipo da ficha; devolve ``None`` se não reconhecer.

    Vence o marcador que aparecer **mais no topo** do documento, como manda a
    regra de negócio ("se contiver no topo a seção ...").
    """
    padroes = [
        (modelo, re.compile("|".join(f"(?:{p})" for p in modelo.deteccao)))
        for modelo in MODELOS
    ]
    for no in documento.nos:
        if not no.texto:
            continue
        for modelo, regex in padroes:
            if casa(regex, no.texto) or casa(regex, no.texto.split("\n", 1)[0]):
                return modelo
    return None


@dataclass
class Resultado:
    """Resultado da extração de um documento."""

    valores: dict[str, str] = field(default_factory=dict)
    nao_encontrados: list[str] = field(default_factory=list)


class Extrator:
    """Aplica o mapa de campos de um :class:`Modelo` sobre um documento lido."""

    def __init__(self, modelo: Modelo, separador: str = SEPARADOR_PADRAO) -> None:
        self.modelo = modelo
        self.separador = separador
        # Estado do documento em extração (o processamento é sequencial).
        self._documento: Documento | None = None
        self._secoes_presentes: set[str] = set()

    # -- API pública -------------------------------------------------------
    def extrair(self, documento: Documento) -> Resultado:
        self._documento = documento
        self._secoes_presentes = {no.secao for no in documento.nos}

        resultado = Resultado()
        for campo in self.modelo.campos:
            try:
                if campo.extrator:
                    metodo = getattr(self, f"_extrair_{campo.extrator}")
                    valor, encontrado = metodo(campo)
                else:
                    valor, encontrado = self._extrair_por_rotulo(campo)
            except Exception:  # pragma: no cover - blindagem por campo
                logger.exception(
                    "Falha ao extrair '%s' de %s", campo.coluna, documento.caminho
                )
                valor, encontrado = "", False

            resultado.valores[campo.coluna] = limpar(valor)
            if not encontrado:
                resultado.nao_encontrados.append(campo.coluna)

        return resultado

    @property
    def documento(self) -> Documento:
        if self._documento is None:  # pragma: no cover - uso indevido da API
            raise RuntimeError("Nenhum documento em extração.")
        return self._documento

    # -- extração genérica -------------------------------------------------
    def _candidatos(self, campo: Campo) -> list[No]:
        """Nós cujo texto casa com o rótulo do campo, na ordem do documento.

        Busca primeiro na seção esperada. Se a seção sequer foi identificada
        no arquivo (documento fora do padrão), cai para uma busca global —
        assim um cabeçalho ausente não zera todos os campos daquela seção.

        A prioridade entre as formas de casar segue o layout mais comum da
        ficha: célula com *apenas* o rótulo, depois "Rótulo: valor" na mesma
        célula, e por fim rótulo na primeira linha com o valor abaixo.
        """
        regex = self._regex(campo)

        def buscar(nos: list[No]) -> list[No]:
            achados: list[list[No]] = [[], [], []]
            for no in nos:
                forma = self._forma_do_rotulo(regex, no.texto)
                # EXATO é 0: comparar com None, senão o casamento mais forte some.
                if forma is not None:
                    achados[forma].append(no)
            return [no for grupo in achados for no in grupo]

        documento = self.documento
        escopo = documento.nos_da_secao(campo.secao) if campo.secao else documento.nos
        candidatos = buscar(escopo)
        if not candidatos and campo.secao and campo.secao not in self._secoes_presentes:
            candidatos = buscar(documento.nos)
        return candidatos

    def _regex(self, campo: Campo):
        return re.compile("|".join(f"(?:{p})" for p in campo.padroes))

    def _forma_do_rotulo(self, regex, texto: str) -> int | None:
        """Como (e se) o texto anuncia este rótulo.

        Devolve ``EXATO``, ``EMBUTIDO`` (Rótulo: valor), ``PRIMEIRA_LINHA``
        (rótulo em cima, valor abaixo na mesma célula) ou ``None``.
        """
        if not texto:
            return None
        if casa(regex, texto):
            return EXATO
        partes = _DIVISOR_INLINE.match(texto)
        if partes and casa(regex, partes.group(1)):
            return EMBUTIDO
        if "\n" in texto and casa(regex, texto.split("\n", 1)[0]):
            return PRIMEIRA_LINHA
        return None

    def _extrair_por_rotulo(self, campo: Campo) -> tuple[str, bool]:
        candidatos = self._candidatos(campo)
        if not candidatos:
            return "", False

        indice = min(campo.ocorrencia, len(candidatos) - 1)
        no = candidatos[indice]
        forma = self._forma_do_rotulo(self._regex(campo), no.texto)

        if forma == EMBUTIDO:
            valor = _DIVISOR_INLINE.match(no.texto).group(2)
            if not eh_rotulo(valor):
                return valor, True
        elif forma == PRIMEIRA_LINHA:
            resto = no.texto.split("\n", 1)[1].strip()
            if resto and not eh_rotulo(resto):
                return resto, True

        if no.tipo == "celula":
            tabela = self.documento.tabelas[no.tabela]
            return self._valor_na_tabela(tabela, no, campo), True
        return self._valor_apos_paragrafo(no, campo), True

    def _valor_na_tabela(self, tabela: Tabela, no: No, campo: Campo) -> str:
        """Procura o valor à direita e, em seguida, abaixo do rótulo."""
        for texto in tabela.celulas_a_direita(no.linha, no.coluna):
            if limpar(texto) and not eh_rotulo(texto):
                return texto

        abaixo = tabela.celulas_abaixo(no.linha, no.coluna)
        if not campo.multiplo:
            for texto in abaixo:
                if limpar(texto) and not eh_rotulo(texto):
                    return texto
            return ""

        # Campo múltiplo: consome a coluna até o próximo rótulo conhecido.
        itens: list[str] = []
        for texto in abaixo:
            if eh_rotulo(texto):
                break
            if limpar(texto):
                itens.append(texto)
        return self._unir_itens(itens)

    def _valor_apos_paragrafo(self, no: No, campo: Campo) -> str:
        """Fallback para fichas escritas em parágrafos, sem tabelas."""
        itens: list[str] = []
        for seguinte in self.documento.nos[no.ordem + 1 :]:
            if seguinte.tipo != "paragrafo":
                break
            if eh_rotulo(seguinte.texto):
                break
            if limpar(seguinte.texto):
                itens.append(seguinte.texto)
                if not campo.multiplo:
                    break
        return self._unir_itens(itens)

    def _unir_itens(self, itens: list[str]) -> str:
        """Une itens de uma lista usando o separador configurado.

        Com um separador diferente de ``\\n``, as quebras de linha internas de
        uma célula também são convertidas: na ficha, várias causas/ações
        críticas costumam vir em uma única célula separadas por quebra manual,
        e quem escolhe ``--separador ponto-virgula`` espera "a; b; c".
        """
        if self.separador != "\n":
            itens = [linha for item in itens for linha in item.split("\n")]
        return juntar(itens, self.separador)

    def _valor_por_padrao(self, padrao: str, secao: str | None) -> str:
        """Atalho para extrair um rótulo avulso (usado pelos extratores)."""
        valor, _ = self._extrair_por_rotulo(
            Campo(coluna="", padroes=(padrao,), secao=secao)
        )
        return valor

    # -- localização de tabelas -------------------------------------------
    def _linha_com_rotulo(self, secao: str | None, *padroes: str) -> tuple[Tabela, int] | None:
        """Localiza a linha cuja 1ª chave casa e que contém todas as demais."""
        ancora = re.compile(padroes[0])
        restantes = [re.compile(padrao) for padrao in padroes[1:]]
        for no in self.documento.nos_da_secao(secao):
            if no.tipo != "celula" or not casa(ancora, no.texto):
                continue
            tabela = self.documento.tabelas[no.tabela]
            textos = [texto for _, texto in tabela.linha_de_celulas(no.linha)]
            if all(any(casa(regex, texto) for texto in textos) for regex in restantes):
                return tabela, no.linha
        return None

    def _linhas_da_tabela(
        self, tabela: Tabela, linha_cabecalho: int
    ) -> list[list[tuple[str, str]]]:
        """Linhas de dados abaixo do cabeçalho, como pares (cabeçalho, valor).

        Para no primeiro rótulo conhecido: é onde o bloco termina e começa
        outra parte da ficha.
        """
        cabecalhos = [
            (coluna, limpar(texto) or f"Coluna {coluna + 1}")
            for coluna, texto in tabela.linha_de_celulas(linha_cabecalho)
        ]

        linhas: list[list[tuple[str, str]]] = []
        for numero in range(linha_cabecalho + 1, tabela.total_linhas):
            celulas = dict(tabela.linha_de_celulas(numero))
            if any(eh_rotulo(texto) for texto in celulas.values()):
                break
            preenchidas = [
                (titulo, limpar(celulas.get(coluna, "")))
                for coluna, titulo in cabecalhos
                if limpar(celulas.get(coluna, ""))
            ]
            if preenchidas:
                linhas.append(preenchidas)
        return linhas

    def _consolidar_linhas(self, linhas: list[list[tuple[str, str]]]) -> list[str]:
        return [
            SEPARADOR_DE_COLUNAS.join(f"{titulo}: {valor}" for titulo, valor in linha)
            for linha in linhas
        ]

    # -- extratores especiais ---------------------------------------------
    def _extrair_tabela_estruturada(self, campo: Campo) -> tuple[str, bool]:
        """Concatena uma tabela "cabeçalho + linhas" em um único texto.

        Cada linha vira ``Cabeçalho: valor | Cabeçalho: valor``; as linhas são
        unidas pelo separador configurado. Cobre Programas Especiais,
        Propostas de Escuta Social, Ações Orçamentárias e seus Produtos.
        """
        localizacao = self._linha_com_rotulo(campo.secao, *campo.parametro)
        if localizacao is None:
            return "", False
        tabela, linha_cabecalho = localizacao
        registros = self._consolidar_linhas(self._linhas_da_tabela(tabela, linha_cabecalho))
        return juntar(registros, self.separador), True

    def _extrair_recursos_orcamentarios(self, campo: Campo) -> tuple[str, bool]:
        """Fontes de recurso mais a linha de "Total dos Recursos" que as fecha."""
        localizacao = self._linha_com_rotulo(campo.secao, r"codigoda?fonte")
        if localizacao is None:
            return "", False
        tabela, linha_cabecalho = localizacao

        registros = self._consolidar_linhas(self._linhas_da_tabela(tabela, linha_cabecalho))

        # O total fica em uma linha rotulada, onde o laço acima parou.
        for numero in range(linha_cabecalho + 1, tabela.total_linhas):
            celulas = [texto for _, texto in tabela.linha_de_celulas(numero)]
            if not any(chave(texto) == "totaldosrecursos" for texto in celulas):
                continue
            valores = [
                limpar(texto)
                for texto in celulas
                if limpar(texto) and chave(texto) != "totaldosrecursos"
            ]
            if valores:
                registros.append(f"Total dos Recursos: {valores[-1]}")
            break

        return juntar(registros, self.separador), True

    def _marcacao(self, padrao: str) -> tuple[str, bool]:
        """Lê uma caixa de seleção do tipo "Estado [x]  Território [ ]".

        A linha é localizada por "Estado"; há versões do modelo em que o
        "Território de Identidade" não aparece ao lado, e nesse caso só a
        marcação existente é reportada.
        """
        localizacao = self._linha_com_rotulo("TERRITORIAL", r"estado")
        if localizacao is None:
            return "", False
        tabela, linha = localizacao

        regex = re.compile(padrao)
        celulas = tabela.linha_de_celulas(linha)
        for posicao, (_, texto) in enumerate(celulas):
            if not casa(regex, texto):
                continue
            if esta_marcado(texto):  # marca dentro da própria célula (☒)
                return "Sim", True
            seguinte = celulas[posicao + 1][1] if posicao + 1 < len(celulas) else ""
            return ("Sim" if esta_marcado(seguinte) else "Não"), True
        return "", False

    def _coluna_da_tabela_territorial(self, padrao: str) -> str:
        """Lê uma coluna da tabela "Território | Memória | Meta"."""
        localizacao = self._linha_com_rotulo(
            "TERRITORIAL", r"territoriode?identidade", r"metaterritorial"
        )
        if localizacao is None:
            return ""
        tabela, linha_cabecalho = localizacao

        regex = re.compile(padrao)
        coluna_alvo = next(
            (
                coluna
                for coluna, texto in tabela.linha_de_celulas(linha_cabecalho)
                if casa(regex, texto)
            ),
            None,
        )
        if coluna_alvo is None:
            return ""

        itens: list[str] = []
        for texto in tabela.celulas_abaixo(linha_cabecalho, coluna_alvo):
            if eh_rotulo(texto):
                break  # começou outro bloco da ficha
            if limpar(texto):
                itens.append(texto)
        return self._unir_itens(itens)

    def _extrair_desagregacao_territorial(self, campo: Campo) -> tuple[str, bool]:
        """Consolida toda a seção de desagregação territorial em uma célula."""
        estado, encontrada = self._marcacao(r"estado")
        territorio, _ = self._marcacao(r"territoriode?identidade")

        itens = [
            ("Estado", estado),
            ("Território de Identidade", territorio),
            (
                "Fórmula de cálculo Territorial",
                self._valor_por_padrao(r"formulade?calculoterritorial", "TERRITORIAL"),
            ),
            ("Unidade de Medida", self._valor_por_padrao(r"unidadede?medida", "TERRITORIAL")),
            ("Memória de Cálculo", self._valor_por_padrao(r"memoriade?calculo", "TERRITORIAL")),
            (
                "Territórios de Identidade",
                self._coluna_da_tabela_territorial(r"territoriode?identidade"),
            ),
            (
                "Memória de Cálculo Territorial",
                self._coluna_da_tabela_territorial(r"memoriade?calculoterritorial"),
            ),
            ("Meta Territorial", self._coluna_da_tabela_territorial(r"metaterritorial")),
            (
                "Outras possibilidades de Regionalização",
                self._valor_por_padrao(r"outraspossibilidadesde?regionalizacao", "TERRITORIAL"),
            ),
        ]

        texto = juntar(
            (f"{titulo}: {valor}" for titulo, valor in itens if valor), self.separador
        )
        return texto, bool(texto) or encontrada
