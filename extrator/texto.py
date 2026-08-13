"""Utilidades de normalização de texto.

O parser precisa comparar rótulos ("Fórmula de Cálculo", "ATRIBUTOSDO
INDICADOR...") que, na prática, aparecem com acentuação inconsistente,
espaços faltando/sobrando, quebras de linha e pontuação variável.

A estratégia é reduzir qualquer texto a uma *chave canônica*: minúsculas,
sem acentos e sem nenhum caractere que não seja letra ou dígito. Assim,
"Indicador(es) doPrograma Sensibilizado(s)" e
"Indicador(es) do Programa Sensibilizado(s)" produzem exatamente a mesma
chave, tornando o casamento de rótulos resiliente a variações de layout.
"""

from __future__ import annotations

import re
import unicodedata

# Espaços "exóticos" comuns em documentos Word (NBSP, narrow NBSP, etc.).
_ESPACOS_EXOTICOS = re.compile(r"[   -   　]")
_ESPACOS_REPETIDOS = re.compile(r"[ \t]{2,}")
_QUEBRAS_REPETIDAS = re.compile(r"\n{3,}")
_NAO_ALFANUMERICO = re.compile(r"[^a-z0-9]+")

# Caracteres usados como marcação de caixa de seleção nos formulários.
MARCAS_SELECAO = {"x", "X", "✓", "✔", "☑", "☒", "•", "*", "1", "sim"}


def limpar(texto: str | None) -> str:
    """Normaliza espaços preservando as quebras de linha significativas."""
    if not texto:
        return ""
    texto = _ESPACOS_EXOTICOS.sub(" ", texto)
    texto = texto.replace("\r\n", "\n").replace("\r", "\n")
    # Remove espaços nas pontas de cada linha antes de colapsar quebras.
    linhas = [_ESPACOS_REPETIDOS.sub(" ", linha).strip() for linha in texto.split("\n")]
    texto = "\n".join(linhas)
    texto = _QUEBRAS_REPETIDAS.sub("\n\n", texto)
    return texto.strip()


def sem_acentos(texto: str) -> str:
    """Remove diacríticos ('Cálculo' -> 'Calculo')."""
    normalizado = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in normalizado if not unicodedata.combining(c))


def chave(texto: str | None) -> str:
    """Converte um texto na chave canônica usada para casar rótulos.

    >>> chave("Fórmula de Cálculo ")
    'formuladecalculo'
    >>> chave("Indicador(es) doPrograma Sensibilizado(s)")
    'indicadoresdoprogramasensibilizados'
    """
    if not texto:
        return ""
    return _NAO_ALFANUMERICO.sub("", sem_acentos(texto).lower())


_PARENTESES = re.compile(r"\([^()]*\)")
_PREFIXO_BLOCO = re.compile(r"^\s*bloco\s*[\dIVXivx]*\s*[:.\-–—]?\s*", re.IGNORECASE)


def chaves(texto: str | None) -> tuple[str, ...]:
    """Chaves candidatas do texto, da mais literal à mais tolerante.

    Algumas versões do modelo numeram as seções e explicam os rótulos entre
    parênteses — "Bloco 1: VÍNCULO DO INDICADOR DE COMPROMISSO",
    "Compromisso (Objetivo do Compromisso)". A segunda chave remove esses
    ornamentos.

    >>> chaves("Bloco 1: VÍNCULO DO INDICADOR DE COMPROMISSO")[-1]
    'vinculodoindicadordecompromisso'
    >>> chaves("Compromisso (Objetivo do Compromisso)")[-1]
    'compromisso'

    Rótulos em que os parênteses fazem parte do nome continuam funcionando,
    porque os padrões toleram o plural opcional:

    >>> chaves("Causa(s) Crítica(s)")
    ('causascriticas', 'causacritica')
    """
    if not texto:
        return ()

    principal = chave(texto)
    resultado = [principal] if principal else []

    simplificado = _PREFIXO_BLOCO.sub("", _PARENTESES.sub(" ", texto).strip())
    alternativa = chave(simplificado)
    if alternativa and alternativa not in resultado:
        resultado.append(alternativa)
    return tuple(resultado)


def casa(regex, texto: str | None) -> bool:
    """True se alguma chave candidata do texto casa integralmente com o regex."""
    return any(regex.fullmatch(candidata) for candidata in chaves(texto))


def esta_marcado(texto: str | None) -> bool:
    """Indica se o conteúdo de uma célula representa uma caixa marcada."""
    valor = limpar(texto)
    if not valor:
        return False
    if len(valor) <= 3 and valor in MARCAS_SELECAO:
        return True
    return any(marca in valor for marca in ("☒", "✓", "✔", "☑"))


def juntar(valores, separador: str = "\n") -> str:
    """Une itens múltiplos em um único texto, sem duplicatas nem vazios."""
    vistos: list[str] = []
    for valor in valores:
        valor = limpar(valor)
        if valor and valor not in vistos:
            vistos.append(valor)
    return separador.join(vistos)
