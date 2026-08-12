"""Testes do extrator.

Os documentos de teste são gerados em tempo de execução com python-docx,
cobrindo as três situações que aparecem no acervo real:

* layout vertical  (rótulo em uma linha, valor na linha de baixo);
* layout horizontal (rótulo à esquerda, valor à direita);
* layout em parágrafos com "Rótulo: valor" na mesma linha.

Execute com:  pytest -q
"""

from __future__ import annotations

import sys
from pathlib import Path

import docx
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from extrator import ler_documento, montar_dataframe, processar_lote  # noqa: E402
from extrator.parser import Extrator, eh_rotulo  # noqa: E402
from extrator.pipeline import listar_documentos  # noqa: E402
from extrator.planilha import LIMITE_CELULA, salvar_excel, sanitizar  # noqa: E402
from extrator.texto import chave, esta_marcado  # noqa: E402


# ---------------------------------------------------------------------------
# Auxiliares de construção de documentos
# ---------------------------------------------------------------------------
def _tabela_vertical(documento, linhas: list[str]) -> None:
    """Tabela de uma coluna: rótulo, valor, rótulo, valor..."""
    tabela = documento.add_table(rows=len(linhas), cols=1)
    for indice, texto in enumerate(linhas):
        tabela.cell(indice, 0).text = texto


def _tabela_horizontal(documento, pares: list[tuple[str, str]]) -> None:
    """Tabela de duas colunas: rótulo | valor."""
    tabela = documento.add_table(rows=len(pares), cols=2)
    for indice, (rotulo, valor) in enumerate(pares):
        tabela.cell(indice, 0).text = rotulo
        tabela.cell(indice, 1).text = valor


def _tabela_grade(documento, linhas: list[list[str]]) -> None:
    """Tabela com cabeçalhos na primeira linha e valores nas seguintes."""
    tabela = documento.add_table(rows=len(linhas), cols=len(linhas[0]))
    for l, linha in enumerate(linhas):
        for c, texto in enumerate(linha):
            tabela.cell(l, c).text = texto


def criar_ficha_completa(destino: Path) -> Path:
    """Ficha no layout padrão (misto vertical/horizontal)."""
    documento = docx.Document()

    _tabela_vertical(documento, ["VÍNCULO DO INDICADOR DE COMPROMISSO"])
    _tabela_vertical(
        documento,
        [
            "Eixo",
            "EIXO DE TESTE",
            "Programa",
            "Programa de Teste",
            "Compromisso",
            "Compromisso de teste",
            "Problema(s) vinculado(s) ao Compromisso",
            "Problema de teste",
            "Causa(s) Crítica(s)",
            "Causa 1.\nCausa 2.\nCausa 3.",
            "ATRIBUTOSDO INDICADOR DE COMPROMISSO",  # erro de espaço proposital
        ],
    )
    _tabela_vertical(
        documento,
        [
            "Descrição",
            "Descrição do indicador",
            "Fórmula de Cálculo",
            "Somatório de X",
            "Memória de Cálculo",
            "Memória detalhada",
        ],
    )
    _tabela_grade(
        documento,
        [
            ["Unidade de medida", "Valor de referência", "Ano de referência", "Valor da meta"],
            ["Unidade", "6", "2023", "40"],
            ["Periodicidade da apuração", "Polaridade", "Classificação", ""],
            ["Semestral", "Positiva", "Produto", ""],
        ],
    )
    _tabela_vertical(
        documento,
        ["Fonte", "Órgão XPTO", "Meios de verificação", "Relatórios administrativos"],
    )
    _tabela_grade(
        documento,
        [["Sigla do Órgão", "UO", "USP"], ["SJDH", "APG", "O16 GASEC"]],
    )

    _tabela_grade(
        documento,
        [
            ["DESAGREGAÇÃO TERRITORIAL", "", "", ""],
            ["Estado", "x", "Território de Identidade", ""],
            ["Fórmula de cálculo Territorial", "", "Unidade de Medida", ""],
            ["Não se aplica", "", "Unidade", ""],
            ["Memória de Cálculo", "", "", ""],
            ["Sem memória", "", "", ""],
            ["Território de Identidade", "Memória de Cálculo Territorial", "Meta Territorial", ""],
            ["Metropolitano", "Soma regional", "10", ""],
            ["Recôncavo", "Soma regional", "5", ""],
            ["Outras possibilidades de Regionalização", "", "", ""],
            ["Não se aplica", "", "", ""],
        ],
    )

    _tabela_vertical(documento, ["INFORMAÇÕES COMPLEMENTARES"])
    _tabela_vertical(
        documento,
        [
            "Objetivo/ Interpretação e uso",
            "Mede a evolução de X",
            "Limitações do Indicador",
            "Não se aplica",
            "Fragilidades para apuração e ações em curso para superação",
            "Nenhuma",
        ],
    )
    _tabela_horizontal(
        documento,
        [
            ("Limitações para definição do valor da meta", ""),
            ("Operacionais", "Limitação operacional"),
            ("Orçamentárias/Financeiras", "Limitação orçamentária"),
            ("Institucionais ou políticas", "Limitação institucional"),
        ],
    )
    _tabela_vertical(documento, ["Possibilidade de desagregação populacional", "Sim, por sexo"])
    _tabela_grade(
        documento,
        [
            ["Programas Especiais", "", ""],
            ["Nome do Programa", "Memória de Cálculo", "Meta"],
            ["Programa Especial A", "Soma", "12"],
        ],
    )
    _tabela_vertical(
        documento, ["Indicador(es) doPrograma Sensibilizado(s)", "Indicador A", "Indicador B"]
    )

    documento.save(destino)
    return destino


def criar_ficha_em_paragrafos(destino: Path) -> Path:
    """Variação fora do padrão: texto corrido com "Rótulo: valor"."""
    documento = docx.Document()
    documento.add_paragraph("VÍNCULO DO INDICADOR DE COMPROMISSO")
    documento.add_paragraph("Eixo: EIXO EM PARÁGRAFO")
    documento.add_paragraph("Programa")
    documento.add_paragraph("Programa em parágrafo")
    documento.add_paragraph("ATRIBUTOS DO INDICADOR DE COMPROMISSO")
    documento.add_paragraph("Descrição: Indicador descrito em parágrafo")
    documento.add_paragraph("Fonte: Órgão do parágrafo")
    documento.save(destino)
    return destino


@pytest.fixture(scope="module")
def pasta(tmp_path_factory) -> Path:
    destino = tmp_path_factory.mktemp("fichas")
    criar_ficha_completa(destino / "ficha_completa.docx")
    criar_ficha_em_paragrafos(destino / "ficha_paragrafos.docx")
    (destino / "corrompido.docx").write_bytes(b"isto nao e um docx valido")
    (destino / "~$temporario.docx").write_bytes(b"lixo")
    (destino / "ignorar.txt").write_text("nao e docx")
    return destino


@pytest.fixture(scope="module")
def valores(pasta: Path) -> dict[str, str]:
    documento = ler_documento(str(pasta / "ficha_completa.docx"))
    return Extrator().extrair(documento).valores


# ---------------------------------------------------------------------------
# Normalização de texto
# ---------------------------------------------------------------------------
def test_chave_ignora_acentos_espacos_e_pontuacao():
    assert chave("Fórmula de Cálculo") == "formuladecalculo"
    assert chave("Indicador(es) doPrograma Sensibilizado(s)") == chave(
        "Indicador(es) do Programa Sensibilizado(s)"
    )


def test_reconhecimento_de_rotulos():
    assert eh_rotulo("Valor de referência")
    assert eh_rotulo("ATRIBUTOSDO INDICADOR DE COMPROMISSO")
    assert not eh_rotulo("Semestral")
    assert not eh_rotulo("40")


def test_marcacao_de_caixa_de_selecao():
    assert esta_marcado("x") and esta_marcado("☒")
    assert not esta_marcado("") and not esta_marcado("☐")


# ---------------------------------------------------------------------------
# Extração de campos
# ---------------------------------------------------------------------------
@pytest.mark.parametrize(
    ("coluna", "esperado"),
    [
        ("Eixo", "EIXO DE TESTE"),
        ("Programa", "Programa de Teste"),
        ("Compromisso", "Compromisso de teste"),
        ("Problemas_Vinculados", "Problema de teste"),
        ("Descricao", "Descrição do indicador"),
        ("Formula_de_Calculo", "Somatório de X"),
        ("Memoria_de_Calculo", "Memória detalhada"),
        ("Unidade_de_Medida", "Unidade"),
        ("Valor_de_Referencia", "6"),
        ("Ano_de_Referencia", "2023"),
        ("Valor_da_Meta", "40"),
        ("Periodicidade_da_Apuracao", "Semestral"),
        ("Polaridade", "Positiva"),
        ("Classificacao", "Produto"),
        ("Fonte", "Órgão XPTO"),
        ("Meios_de_Verificacao", "Relatórios administrativos"),
        ("Responsavel_Sigla_Orgao", "SJDH"),
        ("Responsavel_UO", "APG"),
        ("Responsavel_USP", "O16 GASEC"),
        ("Desagregacao_Estado", "Sim"),
        ("Desagregacao_Territorio_Identidade", "Não"),
        ("Formula_Calculo_Territorial", "Não se aplica"),
        ("Unidade_Medida_Territorial", "Unidade"),
        ("Memoria_Calculo_Territorial", "Sem memória"),
        ("Outras_Possibilidades_Regionalizacao", "Não se aplica"),
        ("Objetivo_Interpretacao_Uso", "Mede a evolução de X"),
        ("Limitacoes_do_Indicador", "Não se aplica"),
        ("Fragilidades_para_Apuracao", "Nenhuma"),
        ("Limitacoes_Meta_Operacionais", "Limitação operacional"),
        ("Limitacoes_Meta_Orcamentarias", "Limitação orçamentária"),
        ("Limitacoes_Meta_Institucionais", "Limitação institucional"),
        ("Possibilidade_Desagregacao_Populacional", "Sim, por sexo"),
    ],
)
def test_campos_simples(valores, coluna, esperado):
    assert valores[coluna] == esperado


def test_lista_permanece_em_uma_unica_celula(valores):
    assert valores["Causas_Criticas"] == "Causa 1.\nCausa 2.\nCausa 3."


def test_separador_configuravel(pasta: Path):
    documento = ler_documento(str(pasta / "ficha_completa.docx"))
    valores = Extrator(separador="; ").extrair(documento).valores
    assert valores["Indicadores_Programa_Sensibilizado"] == "Indicador A; Indicador B"


def test_tabela_territorial_agrupa_todas_as_linhas(valores):
    assert valores["Territorios_de_Identidade"] == "Metropolitano\nRecôncavo"
    assert valores["Metas_Territoriais"] == "10\n5"
    # Valores repetidos entre territórios não são duplicados no texto final.
    assert valores["Memoria_Calculo_Por_Territorio"] == "Soma regional"


def test_programas_especiais_viram_texto_estruturado(valores):
    assert valores["Programas_Especiais"] == (
        "Nome do Programa: Programa Especial A | Memória de Cálculo: Soma | Meta: 12"
    )


def test_layout_em_paragrafos(pasta: Path):
    documento = ler_documento(str(pasta / "ficha_paragrafos.docx"))
    valores = Extrator().extrair(documento).valores
    assert valores["Eixo"] == "EIXO EM PARÁGRAFO"  # "Rótulo: valor" na mesma linha
    assert valores["Programa"] == "Programa em parágrafo"  # valor no parágrafo seguinte
    assert valores["Descricao"] == "Indicador descrito em parágrafo"
    assert valores["Fonte"] == "Órgão do parágrafo"


def test_campos_ausentes_sao_reportados(pasta: Path):
    documento = ler_documento(str(pasta / "ficha_paragrafos.docx"))
    resultado = Extrator().extrair(documento)
    assert "Polaridade" in resultado.nao_encontrados
    assert "Eixo" not in resultado.nao_encontrados


# ---------------------------------------------------------------------------
# Lote, resiliência e planilha
# ---------------------------------------------------------------------------
def test_listagem_ignora_temporarios_e_outras_extensoes(pasta: Path):
    nomes = {caminho.name for caminho in listar_documentos(pasta)}
    assert nomes == {"ficha_completa.docx", "ficha_paragrafos.docx", "corrompido.docx"}


def test_arquivo_corrompido_nao_interrompe_o_lote(pasta: Path):
    arquivos = listar_documentos(pasta)
    registros, estatisticas = processar_lote(arquivos, pasta)

    assert len(registros) == len(arquivos)  # uma linha por arquivo, sempre
    assert estatisticas.erros == 1
    assert estatisticas.sucesso == 2
    with_erro = next(r for r in registros if r["Nome_do_Arquivo"] == "corrompido.docx")
    assert with_erro["Status"] == "ERRO_DE_LEITURA"
    assert with_erro["Observacoes"]


def test_planilha_gerada_tem_uma_linha_por_arquivo(pasta: Path, tmp_path: Path):
    import openpyxl

    arquivos = listar_documentos(pasta)
    registros, _ = processar_lote(arquivos, pasta)
    destino = tmp_path / "consolidado.xlsx"
    salvar_excel(montar_dataframe(registros), destino)

    planilha = openpyxl.load_workbook(destino)["Indicadores"]
    assert planilha.max_row == len(arquivos) + 1  # + cabeçalho
    assert planilha.cell(row=1, column=1).value == "Nome_do_Arquivo"
    assert planilha.cell(row=1, column=1).font.bold
    assert planilha.cell(row=2, column=1).alignment.wrap_text
    assert planilha.cell(row=2, column=1).alignment.vertical == "top"
    assert planilha.freeze_panes == "C2"


def test_sanitizacao_para_o_excel():
    assert sanitizar(None) == ""
    assert sanitizar("=1+1").startswith("'")  # não vira fórmula
    assert sanitizar("a\x07b") == "ab"  # caractere de controle removido
    assert len(sanitizar("x" * (LIMITE_CELULA + 500))) <= LIMITE_CELULA
