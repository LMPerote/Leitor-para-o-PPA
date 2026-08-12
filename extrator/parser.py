"""Motor de extração: transforma um :class:`Documento` em uma linha de dados.

Estratégia geral (independente da posição exata das tabelas no arquivo):

1. Localiza a célula/parágrafo cujo texto casa com o rótulo do campo,
   preferindo ocorrências dentro da seção esperada.
2. Resolve o valor associado tentando, nesta ordem:
   a) valor na própria célula depois de ``:`` (ex.: "Fonte: SEI");
   b) primeira célula **à direita** que não esteja vazia e que não seja,
      ela mesma, outro rótulo conhecido (layout horizontal);
   c) primeira célula **abaixo**, mesmo critério (layout vertical).
3. Campos marcados como ``multiplo`` continuam descendo pela coluna e unem
   todos os itens em um único texto — garantindo "uma célula por campo" no
   Excel mesmo quando o Word traz uma lista.

Layouts que fogem desse padrão (caixas de seleção, tabelas de territórios e
de programas especiais) têm extratores dedicados no fim do módulo.
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field

from .campos import CAMPOS, Campo, ROTULOS_ESTRUTURAIS
from .documento import Documento, No, Tabela
from .texto import chave, esta_marcado, juntar, limpar

logger = logging.getLogger(__name__)

SEPARADOR_PADRAO = "\n"

# Rótulos conhecidos = rótulos de campos + rótulos estruturais.
_TODOS_OS_PADROES: tuple[str, ...] = tuple(
    padrao for campo in CAMPOS for padrao in campo.padroes
) + ROTULOS_ESTRUTURAIS
_REGEX_ROTULOS = re.compile("|".join(f"(?:{p})" for p in _TODOS_OS_PADROES))

# Divisor "rótulo: valor" dentro de uma mesma célula/parágrafo (uma linha).
_DIVISOR_INLINE = re.compile(r"^([^:\n]{2,80}?)\s*[:：]\s*(.+?)\s*(?:\n|$)")


def eh_rotulo(texto: str) -> bool:
    """True se o texto é (apenas) um rótulo conhecido, e não um valor."""
    canonica = chave(texto)
    if not canonica:
        return False
    return _REGEX_ROTULOS.fullmatch(canonica) is not None


@dataclass
class Resultado:
    """Resultado da extração de um documento."""

    valores: dict[str, str] = field(default_factory=dict)
    nao_encontrados: list[str] = field(default_factory=list)


class Extrator:
    """Aplica o mapa de campos sobre um documento já lido."""

    def __init__(self, separador: str = SEPARADOR_PADRAO) -> None:
        self.separador = separador

    # -- API pública -------------------------------------------------------
    def extrair(self, documento: Documento) -> Resultado:
        resultado = Resultado()
        secoes_presentes = {no.secao for no in documento.nos}

        for campo in CAMPOS:
            try:
                if campo.extrator:
                    metodo = getattr(self, f"_extrair_{campo.extrator}")
                    valor, encontrado = metodo(documento, campo)
                else:
                    valor, encontrado = self._extrair_por_rotulo(
                        documento, campo, secoes_presentes
                    )
            except Exception:  # pragma: no cover - blindagem por campo
                logger.exception(
                    "Falha ao extrair '%s' de %s", campo.coluna, documento.caminho
                )
                valor, encontrado = "", False

            resultado.valores[campo.coluna] = limpar(valor)
            if not encontrado:
                resultado.nao_encontrados.append(campo.coluna)

        return resultado

    # -- extração genérica -------------------------------------------------
    def _candidatos(
        self, documento: Documento, campo: Campo, secoes_presentes: set[str]
    ) -> list[No]:
        """Nós cujo texto casa com o rótulo do campo, na ordem do documento.

        Busca primeiro na seção esperada. Se a seção sequer foi identificada
        no arquivo (documento fora do padrão), cai para uma busca global —
        assim um cabeçalho ausente não zera todos os campos daquela seção.

        Células que contêm *apenas* o rótulo têm prioridade sobre células no
        formato "Rótulo: valor", pois são o layout padrão da ficha.
        """
        regex = re.compile("|".join(f"(?:{p})" for p in campo.padroes))

        def exato(no: No) -> bool:
            return bool(no.chave and regex.fullmatch(no.chave))

        def embutido(no: No) -> bool:
            partes = _DIVISOR_INLINE.match(no.texto)
            return bool(partes and regex.fullmatch(chave(partes.group(1))))

        def buscar(nos: list[No]) -> list[No]:
            return [no for no in nos if exato(no)] + [
                no for no in nos if not exato(no) and embutido(no)
            ]

        escopo = documento.nos_da_secao(campo.secao) if campo.secao else documento.nos
        candidatos = buscar(escopo)
        if not candidatos and campo.secao and campo.secao not in secoes_presentes:
            candidatos = buscar(documento.nos)
        return candidatos

    def _extrair_por_rotulo(
        self, documento: Documento, campo: Campo, secoes_presentes: set[str]
    ) -> tuple[str, bool]:
        candidatos = self._candidatos(documento, campo, secoes_presentes)
        if not candidatos:
            return "", False

        indice = min(campo.ocorrencia, len(candidatos) - 1)
        no = candidatos[indice]

        # (a) valor embutido na própria célula ("Fonte: SEI").
        partes = _DIVISOR_INLINE.match(no.texto)
        if partes and not eh_rotulo(partes.group(2)):
            return partes.group(2), True

        if no.tipo == "celula":
            return self._valor_na_tabela(documento.tabelas[no.tabela], no, campo), True
        return self._valor_apos_paragrafo(documento, no, campo), True

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
        return juntar(itens, self.separador)

    def _valor_apos_paragrafo(self, documento: Documento, no: No, campo: Campo) -> str:
        """Fallback para fichas escritas em parágrafos, sem tabelas."""
        itens: list[str] = []
        for seguinte in documento.nos[no.ordem + 1 :]:
            if seguinte.tipo != "paragrafo":
                break
            if eh_rotulo(seguinte.texto):
                break
            if limpar(seguinte.texto):
                itens.append(seguinte.texto)
                if not campo.multiplo:
                    break
        return juntar(itens, self.separador)

    # -- extratores especiais ---------------------------------------------
    def _linha_com_rotulo(
        self, documento: Documento, secao: str | None, *chaves: str
    ) -> tuple[Tabela, int] | None:
        """Localiza a linha de tabela que contém todas as chaves informadas."""
        for no in documento.nos_da_secao(secao):
            if no.tipo != "celula" or no.chave != chaves[0]:
                continue
            tabela = documento.tabelas[no.tabela]
            textos = {chave(texto) for _, texto in tabela.linha_de_celulas(no.linha)}
            if all(c in textos for c in chaves):
                return tabela, no.linha
        return None

    def _marcacao(self, documento: Documento, rotulo: str) -> tuple[str, bool]:
        """Lê uma caixa de seleção do tipo "Estado [x]  Território [ ]"."""
        localizacao = self._linha_com_rotulo(
            documento, "TERRITORIAL", "estado", "territoriodeidentidade"
        )
        if localizacao is None:
            return "", False
        tabela, linha = localizacao

        celulas = tabela.linha_de_celulas(linha)
        for posicao, (_, texto) in enumerate(celulas):
            if chave(texto) != rotulo:
                continue
            if esta_marcado(texto):  # marca dentro da própria célula (☒)
                return "Sim", True
            seguinte = celulas[posicao + 1][1] if posicao + 1 < len(celulas) else ""
            return ("Sim" if esta_marcado(seguinte) else "Não"), True
        return "", False

    def _extrair_marcacao_estado(self, documento: Documento, campo: Campo):
        return self._marcacao(documento, "estado")

    def _extrair_marcacao_territorio(self, documento: Documento, campo: Campo):
        return self._marcacao(documento, "territoriodeidentidade")

    def _coluna_da_tabela_territorial(
        self, documento: Documento, rotulo: str
    ) -> tuple[str, bool]:
        """Lê uma coluna da tabela "Território | Memória | Meta"."""
        localizacao = self._linha_com_rotulo(
            documento, "TERRITORIAL", "territoriodeidentidade", "metaterritorial"
        )
        if localizacao is None:
            return "", False
        tabela, linha_cabecalho = localizacao

        coluna_alvo = next(
            (
                coluna
                for coluna, texto in tabela.linha_de_celulas(linha_cabecalho)
                if chave(texto) == rotulo
            ),
            None,
        )
        if coluna_alvo is None:
            return "", False

        itens: list[str] = []
        for texto in tabela.celulas_abaixo(linha_cabecalho, coluna_alvo):
            if eh_rotulo(texto):
                break  # começou outro bloco da ficha
            if limpar(texto):
                itens.append(texto)
        return juntar(itens, self.separador), True

    def _extrair_lista_territorios(self, documento: Documento, campo: Campo):
        return self._coluna_da_tabela_territorial(documento, "territoriodeidentidade")

    def _extrair_lista_memoria_territorial(self, documento: Documento, campo: Campo):
        return self._coluna_da_tabela_territorial(documento, "memoriadecalculoterritorial")

    def _extrair_lista_metas_territoriais(self, documento: Documento, campo: Campo):
        return self._coluna_da_tabela_territorial(documento, "metaterritorial")

    def _extrair_programas_especiais(self, documento: Documento, campo: Campo):
        """Consolida a tabela "Nome do Programa | Memória de Cálculo | Meta"."""
        localizacao = self._linha_com_rotulo(
            documento, "COMPLEMENTARES", "nomedoprograma"
        )
        if localizacao is None:
            return "", False
        tabela, linha_cabecalho = localizacao

        cabecalhos = [
            (coluna, limpar(texto) or f"Coluna {coluna + 1}")
            for coluna, texto in tabela.linha_de_celulas(linha_cabecalho)
        ]

        registros: list[str] = []
        for linha in range(linha_cabecalho + 1, tabela.total_linhas):
            celulas = dict(tabela.linha_de_celulas(linha))
            if any(eh_rotulo(texto) for texto in celulas.values()):
                break
            partes = [
                f"{titulo}: {limpar(celulas.get(coluna, ''))}"
                for coluna, titulo in cabecalhos
                if limpar(celulas.get(coluna, ""))
            ]
            if partes:
                registros.append(" | ".join(partes))
        return juntar(registros, self.separador), True
