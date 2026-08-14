"""Montagem do relatório final do lote.

Fica separado da interface para que o terminal e a janela gráfica mostrem
exatamente o mesmo texto.
"""

from __future__ import annotations

from pathlib import Path

from .modelos import MODELOS, MODELOS_POR_CODIGO
from .pipeline import Estatisticas

LARGURA = 66
#: Quantos arquivos com problema listar antes de resumir o restante.
MAXIMO_IGNORADOS_LISTADOS = 30


def montar_relatorio(
    estatisticas: Estatisticas, destinos: dict[str, Path], duracao: float
) -> str:
    """Resumo do lote, por tipo de ficha, pronto para exibição."""
    regua = "=" * LARGURA
    linhas = [regua, "RELATÓRIO DE EXTRAÇÃO", regua]
    linhas.append(f"Arquivos encontrados .......: {estatisticas.total}")

    for modelo in MODELOS:
        quantidade = estatisticas.processados.get(modelo.codigo, 0)
        pendentes = estatisticas.com_pendencias.get(modelo.codigo, 0)
        geradas = estatisticas.linhas.get(modelo.codigo, quantidade)
        linhas.append(f"{quantidade} arquivos de {modelo.rotulo} processados.")
        if geradas != quantidade:
            # Acontece quando a ficha traz mais de um problema, causa crítica,
            # ação crítica ou entrega: cada conjunto vira uma linha.
            linhas.append(
                f"  - {geradas} linhas na planilha "
                "(um problema + causa crítica + ação crítica + entrega por linha)."
            )
        if pendentes:
            linhas.append(f"  - {pendentes} com algum campo sem valor.")
        sem_descricao = estatisticas.com_itens_sem_descricao.get(modelo.codigo, 0)
        if sem_descricao:
            linhas.append(
                f"  - {sem_descricao} com item que remete a outro em vez de "
                "descrever (ver Observacoes)."
            )
        if modelo.codigo in destinos:
            linhas.append(f"  - planilha: {destinos[modelo.codigo]}")

    linhas.append(f"{estatisticas.total_ignorados} arquivos ignorados ou com erro.")
    for ocorrencia in estatisticas.ignorados[:MAXIMO_IGNORADOS_LISTADOS]:
        linhas.append(f"  - {ocorrencia.arquivo}: {ocorrencia.motivo}")
    if estatisticas.total_ignorados > MAXIMO_IGNORADOS_LISTADOS:
        restantes = estatisticas.total_ignorados - MAXIMO_IGNORADOS_LISTADOS
        linhas.append(f"  ... e mais {restantes} (veja o log completo).")

    linhas.append(f"Tempo total ................: {duracao:.1f}s")

    for codigo, ausentes in estatisticas.campos_ausentes.items():
        if not ausentes:
            continue
        linhas.append("-" * LARGURA)
        linhas.append(
            f"Campos mais ausentes em {MODELOS_POR_CODIGO[codigo].rotulo} "
            "(rótulo não localizado ou sem resposta preenchida):"
        )
        mais_ausentes = sorted(ausentes.items(), key=lambda item: -item[1])[:10]
        linhas += [f"  {contagem:>5}x  {coluna}" for coluna, contagem in mais_ausentes]

    linhas.append(regua)
    return "\n".join(linhas)
