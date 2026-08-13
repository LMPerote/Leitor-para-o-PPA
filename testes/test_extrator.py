"""Testes do extrator.

Os documentos de teste são gerados em tempo de execução com python-docx,
cobrindo as situações que aparecem no acervo real:

* os dois tipos de ficha (Indicador de Compromisso e Iniciativa);
* layout vertical  (rótulo em uma linha, valor na linha de baixo);
* layout horizontal (rótulo à esquerda, valor à direita);
* layout em parágrafos com "Rótulo: valor" na mesma linha;
* arquivo corrompido e arquivo de tipo não reconhecido.

Execute com:  pytest -q
"""

from __future__ import annotations

import sys
from pathlib import Path

import docx
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from extrator import ler_documento, montar_dataframe, processar_lote  # noqa: E402
from extrator.modelos import MODELOS_POR_CODIGO  # noqa: E402
from extrator.parser import Extrator, classificar, eh_rotulo  # noqa: E402
from extrator.pipeline import listar_documentos  # noqa: E402
from extrator.planilha import LIMITE_CELULA, salvar_excel, sanitizar  # noqa: E402
from extrator.texto import chave, esta_marcado  # noqa: E402

INDICADOR = MODELOS_POR_CODIGO["INDICADOR"]
INICIATIVA = MODELOS_POR_CODIGO["INICIATIVA"]


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
    colunas = max(len(linha) for linha in linhas)
    tabela = documento.add_table(rows=len(linhas), cols=colunas)
    for l, linha in enumerate(linhas):
        for c, texto in enumerate(linha):
            tabela.cell(l, c).text = texto


def criar_ficha_indicador(destino: Path) -> Path:
    """Ficha de Indicador no layout padrão (misto vertical/horizontal)."""
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
    _tabela_grade(documento, [["Sigla do Órgão", "UO", "USP"], ["SJDH", "APG", "O16 GASEC"]])

    _tabela_grade(
        documento,
        [
            ["DESAGREGAÇÃO TERRITORIAL", "", "", ""],
            ["Estado", "x", "Território de Identidade", ""],
            ["Fórmula de cálculo Territorial", "", "Unidade de Medida", ""],
            ["Fórmula territorial", "", "Unidade", ""],
            ["Memória de Cálculo", "", "", ""],
            ["Sem memória", "", "", ""],
            [
                "Território de Identidade",
                "Memória de Cálculo Territorial",
                "Meta Territorial",
                "",
            ],
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

    documento.save(destino)
    return destino


def criar_ficha_iniciativa(destino: Path) -> Path:
    """Ficha de Iniciativa no layout padrão."""
    documento = docx.Document()

    _tabela_vertical(documento, ["VÍNCULO DA INICIATIVA"])
    _tabela_vertical(
        documento,
        [
            "Eixo",
            "EIXO DA INICIATIVA",
            "Programa",
            "Programa da Iniciativa",
            "Compromisso",
            "Compromisso da iniciativa",
            "Problema(s) vinculado(s) ao Compromisso",
            "Problema da iniciativa",
            "Causa(s) Crítica(s)",
            "Causa A.\nCausa B.",
            "Ação(ões) Crítica(s)",
            "Ação 1.\nAção 2.",
        ],
    )
    _tabela_grade(
        documento,
        [
            [
                "Proposta(s) de Escuta Social Associada(s) ao Compromisso",
                "Status para atendimento",
                "Justificativa(s)",
            ],
            ["Proposta A", "Atendida parcialmente", "-"],
            ["Proposta B", "Não atendida", "Sem orçamento"],
        ],
    )

    _tabela_vertical(documento, ["ATRIBUTOS DA INICIATIVA"])
    _tabela_vertical(
        documento,
        [
            "Descrição",
            "Descrição da iniciativa",
            "Entrega(s) Vinculada(s)",
            "Entrega 1\nEntrega 2",
            "Responsável pela Iniciativa",
        ],
    )
    _tabela_grade(
        documento,
        [
            ["Sigla do Órgão", "UO", "USP", "Órgãos Parceiros"],
            ["SJDH", "APG", "GASEC", "Parceiro A, Parceiro B"],
        ],
    )
    _tabela_grade(
        documento,
        [
            ["Fatores Críticos de Contexto da Iniciativa", ""],
            ["Operacionais", "Fator operacional"],
            ["Orçamentários/Financeiros", "Fator orçamentário"],
            ["Institucionais ou políticos", "Fator institucional"],
        ],
    )
    _tabela_grade(
        documento,
        [
            [
                "Estimativa dos recursos orçamentários/financeiros disponíveis para Iniciativa",
                "",
                "",
            ],
            ["Código da Fonte", "Nome da Fonte", "Montante em R$"],
            ["100", "Tesouro Estadual", "R$ 2.400.000,00"],
            ["Total dos Recursos", "", "R$ 2.400.000,00"],
        ],
    )
    _tabela_vertical(
        documento,
        [
            "Indicador de Compromisso Vinculado",
            "Número de campanhas realizadas",
            "Indicadores de Compromisso Sensibilizados",
            "Indicador sensibilizado A",
        ],
    )
    _tabela_grade(
        documento,
        [
            [
                "Etapa II – Programação Orçamentária - Ações orçamentárias vinculadas à Iniciativa",
                "",
                "",
                "",
            ],
            ["Órgão", "UO", "Código da Ação Orçamentária", "Nome da Ação Orçamentária"],
            ["SJDH", "APG", "4170", "Realização de ações de promoção"],
        ],
    )
    _tabela_grade(
        documento,
        [
            [
                "Etapa II – Programação Orçamentária – Produtos das ações orçamentárias "
                "vinculadas à Iniciativa",
                "",
                "",
                "",
                "",
            ],
            [
                "Órgão",
                "UO",
                "Código da Ação Orçamentária",
                "Código do Produto da Ação Orçamentária",
                "Nome Produto da Ação Orçamentária",
            ],
            ["SJDH", "APG", "4170", "P1", "Ações de promoção realizadas"],
        ],
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


def criar_documento_desconhecido(destino: Path) -> Path:
    """Um .docx válido que não é nenhuma das duas fichas."""
    documento = docx.Document()
    documento.add_paragraph("MEMORANDO INTERNO")
    documento.add_paragraph("Este documento não é uma ficha do PPA.")
    documento.save(destino)
    return destino


@pytest.fixture(scope="module")
def pasta(tmp_path_factory) -> Path:
    destino = tmp_path_factory.mktemp("fichas")
    criar_ficha_indicador(destino / "ficha_indicador.docx")
    criar_ficha_iniciativa(destino / "ficha_iniciativa.docx")
    criar_ficha_em_paragrafos(destino / "ficha_paragrafos.docx")
    criar_documento_desconhecido(destino / "outro_documento.docx")
    (destino / "corrompido.docx").write_bytes(b"isto nao e um docx valido")
    (destino / "~$temporario.docx").write_bytes(b"lixo")
    (destino / "ignorar.txt").write_text("nao e docx")
    return destino


@pytest.fixture(scope="module")
def indicador(pasta: Path) -> dict[str, str]:
    documento = ler_documento(str(pasta / "ficha_indicador.docx"))
    return Extrator(INDICADOR).extrair(documento).valores


@pytest.fixture(scope="module")
def iniciativa(pasta: Path) -> dict[str, str]:
    documento = ler_documento(str(pasta / "ficha_iniciativa.docx"))
    return Extrator(INICIATIVA).extrair(documento).valores


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
    assert eh_rotulo("Proposta(s) de Escuta Social Associada(s) ao Compromisso")
    assert not eh_rotulo("Semestral")
    assert not eh_rotulo("40")


def test_marcacao_de_caixa_de_selecao():
    assert esta_marcado("x") and esta_marcado("☒")
    assert not esta_marcado("") and not esta_marcado("☐")


# ---------------------------------------------------------------------------
# Classificação automática
# ---------------------------------------------------------------------------
def test_classificacao_por_secao_de_vinculo(pasta: Path):
    def tipo(nome: str):
        return classificar(ler_documento(str(pasta / nome)))

    assert tipo("ficha_indicador.docx") is INDICADOR
    assert tipo("ficha_iniciativa.docx") is INICIATIVA
    assert tipo("ficha_paragrafos.docx") is INDICADOR  # vínculo em parágrafo
    assert tipo("outro_documento.docx") is None


# ---------------------------------------------------------------------------
# Campos da ficha de INDICADOR
# ---------------------------------------------------------------------------
@pytest.mark.parametrize(
    ("coluna", "esperado"),
    [
        ("Eixo", "EIXO DE TESTE"),
        ("Programa", "Programa de Teste"),
        ("Compromisso", "Compromisso de teste"),
        ("Problemas_Vinculados", "Problema de teste"),
        ("Atributos_Descricao", "Descrição do indicador"),
        ("Formula_Calculo", "Somatório de X"),
        ("Memoria_Calculo", "Memória detalhada"),
        ("Unidade_Medida", "Unidade"),
        ("Valor_Referencia", "6"),
        ("Ano_Referencia", "2023"),
        ("Valor_Meta", "40"),
        ("Periodicidade_Apuracao", "Semestral"),
        ("Polaridade", "Positiva"),
        ("Classificacao", "Produto"),
        ("Fonte", "Órgão XPTO"),
        ("Meios_Verificacao", "Relatórios administrativos"),
        ("Responsavel_Sigla_Orgao", "SJDH"),
        ("Responsavel_UO", "APG"),
        ("Responsavel_USP", "O16 GASEC"),
        ("Info_Complementares_Objetivo", "Mede a evolução de X"),
        ("Limitacoes_Indicador", "Não se aplica"),
        ("Fragilidades_Apuracao", "Nenhuma"),
        ("Limitacoes_Meta_Operacionais", "Limitação operacional"),
        ("Limitacoes_Meta_Orcamentarias", "Limitação orçamentária"),
        ("Limitacoes_Meta_Institucionais", "Limitação institucional"),
        ("Possibilidade_Desagregacao_Populacional", "Sim, por sexo"),
    ],
)
def test_indicador_campos_simples(indicador, coluna, esperado):
    assert indicador[coluna] == esperado


def test_indicador_lista_permanece_em_uma_unica_celula(indicador):
    assert indicador["Causas_Criticas"] == "Causa 1.\nCausa 2.\nCausa 3."


def test_indicador_desagregacao_territorial_consolidada(indicador):
    territorial = indicador["Desagregacao_Territorial"]
    assert "Estado: Sim" in territorial
    assert "Território de Identidade: Não" in territorial
    assert "Fórmula de cálculo Territorial: Fórmula territorial" in territorial
    assert "Unidade de Medida: Unidade" in territorial
    assert "Memória de Cálculo: Sem memória" in territorial
    assert "Territórios de Identidade: Metropolitano\nRecôncavo" in territorial
    assert "Meta Territorial: 10\n5" in territorial
    assert "Outras possibilidades de Regionalização: Não se aplica" in territorial


def test_indicador_programas_especiais(indicador):
    assert indicador["Programas_Especiais"] == (
        "Nome do Programa: Programa Especial A | Memória de Cálculo: Soma | Meta: 12"
    )


# ---------------------------------------------------------------------------
# Campos da ficha de INICIATIVA
# ---------------------------------------------------------------------------
@pytest.mark.parametrize(
    ("coluna", "esperado"),
    [
        ("Eixo", "EIXO DA INICIATIVA"),
        ("Programa", "Programa da Iniciativa"),
        ("Compromisso", "Compromisso da iniciativa"),
        ("Problemas_Vinculados", "Problema da iniciativa"),
        ("Causas_Criticas", "Causa A.\nCausa B."),
        ("Acoes_Criticas", "Ação 1.\nAção 2."),
        ("Atributos_Descricao", "Descrição da iniciativa"),
        ("Entregas_Vinculadas", "Entrega 1\nEntrega 2"),
        ("Responsavel_Sigla_Orgao", "SJDH"),
        ("Responsavel_UO", "APG"),
        ("Responsavel_USP", "GASEC"),
        ("Orgaos_Parceiros", "Parceiro A, Parceiro B"),
        ("Fatores_Criticos_Operacionais", "Fator operacional"),
        ("Fatores_Criticos_Orcamentarios", "Fator orçamentário"),
        ("Fatores_Criticos_Institucionais", "Fator institucional"),
        ("Indicador_Compromisso_Vinculado", "Número de campanhas realizadas"),
        ("Indicadores_Sensibilizados", "Indicador sensibilizado A"),
    ],
)
def test_iniciativa_campos_simples(iniciativa, coluna, esperado):
    assert iniciativa[coluna] == esperado


def test_iniciativa_propostas_de_escuta_social(iniciativa):
    linhas = iniciativa["Propostas_Escuta_Social"].split("\n")
    assert len(linhas) == 2
    assert linhas[0] == (
        "Proposta(s) de Escuta Social Associada(s) ao Compromisso: Proposta A | "
        "Status para atendimento: Atendida parcialmente | Justificativa(s): -"
    )
    assert "Justificativa(s): Sem orçamento" in linhas[1]


def test_iniciativa_recursos_orcamentarios_incluem_o_total(iniciativa):
    recursos = iniciativa["Recursos_Orcamentarios"].split("\n")
    assert recursos[0] == (
        "Código da Fonte: 100 | Nome da Fonte: Tesouro Estadual | "
        "Montante em R$: R$ 2.400.000,00"
    )
    assert recursos[-1] == "Total dos Recursos: R$ 2.400.000,00"


def test_iniciativa_acoes_e_produtos_orcamentarios(iniciativa):
    assert iniciativa["Acoes_Orcamentarias_Vinculadas"] == (
        "Órgão: SJDH | UO: APG | Código da Ação Orçamentária: 4170 | "
        "Nome da Ação Orçamentária: Realização de ações de promoção"
    )
    produtos = iniciativa["Produtos_Acoes_Orcamentarias"]
    assert "Código do Produto da Ação Orçamentária: P1" in produtos
    assert "Nome Produto da Ação Orçamentária: Ações de promoção realizadas" in produtos


def test_separador_configuravel(pasta: Path):
    documento = ler_documento(str(pasta / "ficha_iniciativa.docx"))
    valores = Extrator(INICIATIVA, separador="; ").extrair(documento).valores
    assert valores["Acoes_Criticas"] == "Ação 1.; Ação 2."


# ---------------------------------------------------------------------------
# Layouts fora do padrão
# ---------------------------------------------------------------------------
def test_layout_em_paragrafos(pasta: Path):
    documento = ler_documento(str(pasta / "ficha_paragrafos.docx"))
    valores = Extrator(INDICADOR).extrair(documento).valores
    assert valores["Eixo"] == "EIXO EM PARÁGRAFO"  # "Rótulo: valor" na mesma linha
    assert valores["Programa"] == "Programa em parágrafo"  # valor no parágrafo seguinte
    assert valores["Atributos_Descricao"] == "Indicador descrito em parágrafo"
    assert valores["Fonte"] == "Órgão do parágrafo"


def test_campos_ausentes_sao_reportados(pasta: Path):
    documento = ler_documento(str(pasta / "ficha_paragrafos.docx"))
    resultado = Extrator(INDICADOR).extrair(documento)
    assert "Polaridade" in resultado.nao_encontrados
    assert "Eixo" not in resultado.nao_encontrados


# ---------------------------------------------------------------------------
# Lote, resiliência e planilhas
# ---------------------------------------------------------------------------
def test_listagem_ignora_temporarios_e_outras_extensoes(pasta: Path):
    nomes = {caminho.name for caminho in listar_documentos(pasta)}
    assert nomes == {
        "ficha_indicador.docx",
        "ficha_iniciativa.docx",
        "ficha_paragrafos.docx",
        "outro_documento.docx",
        "corrompido.docx",
    }


def test_lote_separa_por_tipo_e_nao_para_em_erros(pasta: Path):
    arquivos = listar_documentos(pasta)
    registros, estatisticas = processar_lote(arquivos, pasta)

    assert estatisticas.total == len(arquivos)
    assert estatisticas.processados["INDICADOR"] == 2  # tabelas + parágrafos
    assert estatisticas.processados["INICIATIVA"] == 1
    assert len(registros["INDICADOR"]) == 2
    assert len(registros["INICIATIVA"]) == 1

    # Corrompido e "não é ficha do PPA" ficam de fora, mas são reportados.
    ignorados = {
        ocorrencia.arquivo: ocorrencia.motivo for ocorrencia in estatisticas.ignorados
    }
    assert set(ignorados) == {"corrompido.docx", "outro_documento.docx"}
    assert "erro de leitura" in ignorados["corrompido.docx"]
    assert "tipo não reconhecido" in ignorados["outro_documento.docx"]


def test_planilhas_geradas_tem_uma_linha_por_arquivo(pasta: Path, tmp_path: Path):
    import openpyxl

    arquivos = listar_documentos(pasta)
    registros, _ = processar_lote(arquivos, pasta)

    for modelo, esperado in ((INDICADOR, 2), (INICIATIVA, 1)):
        destino = tmp_path / modelo.arquivo_saida
        salvar_excel(montar_dataframe(registros[modelo.codigo], modelo), destino, modelo)

        planilha = openpyxl.load_workbook(destino)[destino.stem]
        assert planilha.max_row == esperado + 1  # + cabeçalho
        assert planilha.max_column == len(modelo.colunas())
        assert planilha.cell(row=1, column=1).value == "Nome_Arquivo"
        assert planilha.cell(row=1, column=1).font.bold
        assert planilha.cell(row=2, column=1).alignment.wrap_text
        assert planilha.cell(row=2, column=1).alignment.vertical == "top"
        assert planilha.freeze_panes == "A2"  # só o cabeçalho; nenhuma coluna fixa


def test_planilha_vazia_ainda_tem_cabecalho(tmp_path: Path):
    import openpyxl

    destino = tmp_path / INICIATIVA.arquivo_saida
    salvar_excel(montar_dataframe([], INICIATIVA), destino, INICIATIVA)

    planilha = openpyxl.load_workbook(destino)[destino.stem]
    assert planilha.max_row == 1
    assert [celula.value for celula in planilha[1]] == INICIATIVA.colunas()


def test_sanitizacao_para_o_excel():
    assert sanitizar(None) == ""
    assert sanitizar("=1+1").startswith("'")  # não vira fórmula
    assert sanitizar("a\x07b") == "ab"  # caractere de controle removido
    assert len(sanitizar("x" * (LIMITE_CELULA + 500))) <= LIMITE_CELULA


# ---------------------------------------------------------------------------
# Fluxo completo (compartilhado pela linha de comando e pela janela gráfica)
# ---------------------------------------------------------------------------
def test_executar_grava_as_duas_planilhas(pasta: Path, tmp_path: Path):
    from extrator.aplicacao import Opcoes, executar

    saida = tmp_path / "planilhas"
    execucao = executar(Opcoes(entrada=pasta, saida=saida))

    assert (saida / "Indicadores.xlsx").is_file()
    assert (saida / "Iniciativas.xlsx").is_file()
    assert execucao.estatisticas.processados == {"INDICADOR": 2, "INICIATIVA": 1}
    assert execucao.destinos["INDICADOR"] == (saida / "Indicadores.xlsx").resolve()


def test_executar_respeita_limite_e_progresso(pasta: Path, tmp_path: Path):
    from extrator.aplicacao import Opcoes, executar

    class Contador:
        def __init__(self, total):
            self.total, self.passos = total, 0

        def update(self, quantidade=1):
            self.passos += quantidade

    criados: list[Contador] = []

    def criar(total):
        criados.append(Contador(total))
        return criados[-1]

    execucao = executar(
        Opcoes(entrada=pasta, saida=tmp_path / "p", limite=2), criar_progresso=criar
    )
    assert len(execucao.arquivos) == 2
    assert criados[0].total == 2 and criados[0].passos == 2


def test_executar_sem_documentos_avisa(tmp_path: Path):
    from extrator.aplicacao import Opcoes, PastaSemDocumentos, executar

    vazia = tmp_path / "vazia"
    vazia.mkdir()
    with pytest.raises(PastaSemDocumentos):
        executar(Opcoes(entrada=vazia, saida=tmp_path / "p"))


def test_relatorio_traz_contagem_por_tipo_e_ignorados(pasta: Path, tmp_path: Path):
    from extrator.aplicacao import Opcoes, executar
    from extrator.relatorio import montar_relatorio

    execucao = executar(Opcoes(entrada=pasta, saida=tmp_path / "p"))
    texto = montar_relatorio(execucao.estatisticas, execucao.destinos, 1.23)

    assert "2 arquivos de Indicador de Compromisso processados." in texto
    assert "1 arquivos de Iniciativa processados." in texto
    assert "2 arquivos ignorados ou com erro." in texto
    assert "corrompido.docx" in texto and "outro_documento.docx" in texto
    assert "1.2s" in texto


def test_apelidos_de_separador():
    from extrator.aplicacao import resolver_separador

    assert resolver_separador("quebra") == "\n"
    assert resolver_separador("ponto-virgula") == "; "
    assert resolver_separador(" / ") == " / "  # texto literal passa direto


# ---------------------------------------------------------------------------
# Regra de negócio: UM problema por linha
# ---------------------------------------------------------------------------
def _ficha_com_dois_problemas(destino: Path) -> Path:
    documento = docx.Document()
    _tabela_vertical(documento, ["VÍNCULO DO INDICADOR DE COMPROMISSO"])
    _tabela_vertical(
        documento,
        [
            "Eixo",
            "EIXO COM DOIS PROBLEMAS",
            "Problema(s) vinculado(s) ao Compromisso",
            "Primeiro problema.\nSegundo problema.",
            "Causa(s) Crítica(s)",
            "Causa única.",
            "ATRIBUTOS DO INDICADOR DE COMPROMISSO",
        ],
    )
    _tabela_vertical(documento, ["Descrição", "Indicador de teste"])
    documento.save(destino)
    return destino


def test_ficha_com_dois_problemas_gera_duas_linhas(tmp_path: Path):
    pasta = tmp_path / "dois"
    pasta.mkdir()
    _ficha_com_dois_problemas(pasta / "ficha.docx")

    registros, estatisticas = processar_lote(listar_documentos(pasta), pasta)
    linhas = registros["INDICADOR"]

    assert [linha["Problemas_Vinculados"] for linha in linhas] == [
        "Primeiro problema.",
        "Segundo problema.",
    ]
    # As demais colunas se repetem em todas as linhas da mesma ficha.
    for coluna in ("Nome_Arquivo", "Eixo", "Causas_Criticas", "Atributos_Descricao"):
        assert linhas[0][coluna] == linhas[1][coluna]
    assert linhas[0]["Eixo"] == "EIXO COM DOIS PROBLEMAS"

    # O arquivo continua contando como UM arquivo processado, com duas linhas.
    assert estatisticas.processados["INDICADOR"] == 1
    assert estatisticas.linhas["INDICADOR"] == 2


def test_ficha_com_um_problema_gera_uma_linha(pasta: Path):
    registros, estatisticas = processar_lote(listar_documentos(pasta), pasta)
    assert estatisticas.processados["INDICADOR"] == estatisticas.linhas["INDICADOR"]
    assert len(registros["INICIATIVA"]) == 1


def test_expansao_respeita_o_separador_escolhido():
    from extrator.pipeline import expandir_por_problema

    registro = {"Nome_Arquivo": "f.docx", "Problemas_Vinculados": "A; B; C"}
    linhas = expandir_por_problema(registro, "; ")
    assert [linha["Problemas_Vinculados"] for linha in linhas] == ["A", "B", "C"]
    assert all(linha["Nome_Arquivo"] == "f.docx" for linha in linhas)


def test_ficha_sem_problema_nao_some_da_planilha():
    from extrator.pipeline import expandir_por_problema

    registro = {"Nome_Arquivo": "f.docx", "Problemas_Vinculados": ""}
    assert expandir_por_problema(registro, "\n") == [registro]
