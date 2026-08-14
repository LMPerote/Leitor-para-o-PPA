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
from .texto import casa, chave, esta_marcado, juntar, limpar

logger = logging.getLogger(__name__)

SEPARADOR_PADRAO = "\n"
#: Separador entre as colunas de uma mesma linha de tabela consolidada.
SEPARADOR_DE_COLUNAS = " | "

_REGEX_ROTULOS = re.compile("|".join(f"(?:{p})" for p in ROTULOS_CONHECIDOS))

#: Formas de um texto anunciar um rótulo, da mais literal à mais tolerante.
EXATO, EMBUTIDO, PRIMEIRA_LINHA = 0, 1, 2

#: Linha que fecha a tabela de recursos, com os nomes que ela recebe nas
#: diferentes versões do modelo.
_REGEX_TOTAL = re.compile(r"total(dos)?recursos|totaldo?teto")

#: Respostas da pergunta única de desagregação territorial, por chave canônica.
_RESPOSTAS_SIM_NAO = {"sim": "Sim", "nao": "Não"}

# Divisor "rótulo: valor" dentro de uma mesma célula/parágrafo (uma linha).
_DIVISOR_INLINE = re.compile(r"^([^:\n]{2,80}?)\s*[:：]\s*(.+?)\s*(?:\n|$)")


def eh_rotulo(texto: str) -> bool:
    """True se o texto é (apenas) um rótulo conhecido, e não um valor."""
    return casa(_REGEX_ROTULOS, texto)


def anuncia_rotulo(texto: str) -> bool:
    """True se o texto **anuncia** um rótulo conhecido, e não só o é.

    Diferente de :func:`eh_rotulo`, reconhece as três formas do documento:
    a célula com apenas o rótulo, o "Rótulo: valor" na mesma linha e o rótulo
    na primeira linha com o valor abaixo.
    """
    if eh_rotulo(texto):
        return True
    partes = _DIVISOR_INLINE.match(texto)
    if partes and eh_rotulo(partes.group(1)):
        return True
    return "\n" in texto and eh_rotulo(texto.split("\n", 1)[0])


def tem_conteudo(texto: str, minimo: int = 1) -> bool:
    """True se o texto é resposta de verdade, e não sujeira de formulário.

    Descarta o que só tem pontuação (".", "-", "–") e, quando ``minimo`` é 2,
    também o caractere solto. Esse caractere aparece muito nas fichas: caixas
    de seleção desenhadas em fonte de símbolos guardam no texto uma letra
    qualquer — o "e" que vinha parar na coluna de problemas.
    """
    return len(chave(texto)) >= minimo


def sem_ligacoes_soltas(texto: str) -> str:
    """Devolve o texto sem as linhas que separam itens em vez de descrever um.

    Nas fichas, a lista costuma vir com o "e" de ligação sozinho na linha
    antes do último item, e às vezes sobra um "." de digitação. Uma letra
    solta ou só pontuação não é problema, causa, ação nem entrega: é separação
    entre itens, e não pode virar uma linha da planilha.

    O que tem conteúdo é preservado como está na ficha — inclusive as
    anotações de trabalho de quem preencheu ("AC 4,5,7,810,12,13,14"): a
    limpeza desse tipo de texto é feita no documento, não aqui.
    """
    linhas = [linha for linha in texto.split("\n") if tem_conteudo(linha, 2)]
    return "\n".join(linhas).strip()


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
    #: Subconjunto de ``nao_encontrados`` cujo **rótulo** sequer existe no
    #: documento. Os demais têm o rótulo, mas a ficha ficou sem resposta —
    #: distinção que separa "o modelo da ficha é outro" de "não preencheram".
    rotulos_ausentes: list[str] = field(default_factory=list)


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
                    devolvido = getattr(self, f"_extrair_{campo.extrator}")(campo)
                    if len(devolvido) == 3:
                        valor, encontrado, tinha_rotulo = devolvido
                    else:
                        # Extrator de tabela: não achar a tabela é não achar o rótulo.
                        valor, encontrado = devolvido
                        tinha_rotulo = encontrado
                else:
                    valor, encontrado, tinha_rotulo = self._extrair_por_rotulo(campo)
            except Exception:  # pragma: no cover - blindagem por campo
                logger.exception(
                    "Falha ao extrair '%s' de %s", campo.coluna, documento.caminho
                )
                valor, encontrado, tinha_rotulo = "", False, False

            resultado.valores[campo.coluna] = limpar(valor)
            if not encontrado:
                resultado.nao_encontrados.append(campo.coluna)
                if not tinha_rotulo:
                    resultado.rotulos_ausentes.append(campo.coluna)

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

    def _extrair_por_rotulo(self, campo: Campo) -> tuple[str, bool, bool]:
        """Usa o primeiro rótulo que realmente **produz** um valor.

        O mesmo rótulo costuma aparecer mais de uma vez na ficha: no quadro de
        orientações de preenchimento (onde as células ao lado estão vazias ou
        trazem só a instrução) e, adiante, na tabela preenchida. Parar no
        primeiro deixava a coluna vazia justamente nas fichas que trazem o
        quadro de orientações; por isso a busca continua pelas demais
        ocorrências até encontrar conteúdo.

        Campo cujo rótulo existe mas está **sem resposta** é reportado como não
        encontrado, para aparecer em ``Campos_Nao_Encontrados`` em vez de sair
        calado como uma célula em branco.

        Devolve ``(valor, encontrado, tinha_rotulo)``. O terceiro item separa
        "esta ficha não tem esse rótulo" de "tem o rótulo e ninguém preencheu",
        que é o que distingue modelo de ficha diferente de ficha incompleta.
        """
        candidatos = self._candidatos(campo)
        if not candidatos:
            return "", False, False

        inicio = min(campo.ocorrencia, len(candidatos) - 1)
        for no in candidatos[inicio:]:
            valor = self._valor_do_rotulo(no, campo)
            if valor:
                return valor, True, True
        return "", False, True

    def _valor_do_rotulo(self, no: No, campo: Campo) -> str:
        """Valor associado a uma ocorrência do rótulo (pode não haver)."""
        forma = self._forma_do_rotulo(self._regex(campo), no.texto)

        if forma == EMBUTIDO:
            valor = _DIVISOR_INLINE.match(no.texto).group(2)
            if self._util(valor, campo) and not eh_rotulo(valor):
                return valor
        elif forma == PRIMEIRA_LINHA:
            resto = no.texto.split("\n", 1)[1].strip()
            if not campo.multiplo:
                # Campo de valor único: só a primeira linha preenchida. Nas
                # fichas de controle, a contagem vem seguida da lista de
                # códigos, que não faz parte do valor.
                resto = next((l.strip() for l in resto.split("\n") if l.strip()), "")
            if self._util(resto, campo) and not eh_rotulo(resto):
                return resto

        if no.tipo == "celula":
            tabela = self.documento.tabelas[no.tabela]
            return self._valor_na_tabela(tabela, no, campo)
        return self._valor_apos_paragrafo(no, campo)

    def _util(self, texto: str, campo: Campo) -> bool:
        """True se o texto serve como valor deste campo.

        Em campo de lista (``multiplo``) exige pelo menos dois caracteres
        úteis: são as colunas onde a sujeira das caixas de seleção aparecia.

        Em campo de valor único, qualquer texto serve. Símbolo é resposta
        legítima aqui — "%" é unidade de medida, "-" é o "não se aplica" de
        quem preencheu — e recusá-lo era pior do que aceitá-lo: a busca seguia
        descendo a coluna e trazia o valor do bloco de baixo (a periodicidade
        virava unidade de medida, a polaridade virava ano de referência).
        """
        return tem_conteudo(texto, 2) if campo.multiplo else bool(limpar(texto))

    def _valor_na_tabela(self, tabela: Tabela, no: No, campo: Campo) -> str:
        """Procura o valor à direita e, em seguida, abaixo do rótulo."""
        for texto in tabela.celulas_a_direita(no.linha, no.coluna):
            if self._util(texto, campo) and not eh_rotulo(texto):
                return texto

        abaixo = tabela.celulas_abaixo(no.linha, no.coluna)
        if not campo.multiplo:
            for texto in abaixo:
                # Parar em qualquer célula que *anuncie* outro rótulo, e não só
                # nas que são o rótulo puro: no layout de uma coluna só, o campo
                # seguinte vem como "RÓTULO: valor" na mesma célula, e descer
                # por cima dele trazia a resposta do campo errado.
                if anuncia_rotulo(texto):
                    return ""
                if self._util(texto, campo):
                    return texto
            return ""

        # Campo múltiplo: consome a coluna até o próximo rótulo conhecido ou
        # até a primeira célula vazia depois do valor. A célula vazia importa:
        # há fichas em que o rótulo do campo seguinte foi apagado, restando a
        # célula em branco, e sem essa parada a lista engolia o bloco de baixo
        # (as causas críticas iam parar na coluna de problemas).
        itens: list[str] = []
        for texto in abaixo:
            if eh_rotulo(texto):
                break
            if not self._util(texto, campo):
                if itens:
                    break
                continue  # antes do valor, célula vazia é só espaçamento
            itens.append(texto)
        return self._unir_itens(itens, campo)

    def _valor_apos_paragrafo(self, no: No, campo: Campo) -> str:
        """Fallback para fichas escritas em parágrafos, sem tabelas."""
        itens: list[str] = []
        for seguinte in self.documento.nos[no.ordem + 1 :]:
            if seguinte.tipo != "paragrafo":
                break
            if eh_rotulo(seguinte.texto):
                break
            if self._util(seguinte.texto, campo):
                itens.append(seguinte.texto)
                if not campo.multiplo:
                    break
        return self._unir_itens(itens, campo)

    def _unir_itens(self, itens: list[str], campo: Campo | None = None) -> str:
        """Une itens de uma lista usando o separador configurado.

        Com um separador diferente de ``\\n``, as quebras de linha internas de
        uma célula também são convertidas: na ficha, várias causas/ações
        críticas costumam vir em uma única célula separadas por quebra manual,
        e quem escolhe ``--separador ponto-virgula`` espera "a; b; c".

        Em campo de lista, as linhas que só ligam ou separam itens saem fora
        (ver :func:`sem_ligacoes_soltas`).
        """
        if self.separador != "\n":
            itens = [linha for item in itens for linha in item.split("\n")]
        if campo is not None and campo.multiplo:
            itens = [texto for texto in map(sem_ligacoes_soltas, itens) if texto]
        return juntar(itens, self.separador)

    def _valor_por_padrao(self, padrao: str, secao: str | None) -> str:
        """Atalho para extrair um rótulo avulso (usado pelos extratores)."""
        valor, _, _ = self._extrair_por_rotulo(
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

    def _extrair_compromisso_do_controle(self, campo: Campo) -> tuple[str, bool, bool]:
        """Compromisso da ficha de controle, com apoio na posição.

        Em parte das fichas o rótulo "COMPROMISSO" foi apagado e o texto do
        compromisso escrito por cima da célula. O bloco tem ordem fixa —
        diretório, eixo, programa, compromisso —, então, quando o rótulo não
        existe, o compromisso é a primeira célula preenchida abaixo da do
        "Programa" que não anuncie outro rótulo.
        """
        valor, encontrado, tinha_rotulo = self._extrair_por_rotulo(campo)
        if encontrado or tinha_rotulo:
            # Rótulo presente: sem resposta é ficha em branco, não posição.
            return valor, encontrado, tinha_rotulo

        programa = Campo(coluna="", padroes=(r"programa",), secao=campo.secao)
        for no in self._candidatos(programa):
            if no.tipo != "celula":
                continue
            tabela = self.documento.tabelas[no.tabela]
            for texto in tabela.celulas_abaixo(no.linha, no.coluna):
                if not limpar(texto):
                    continue
                if anuncia_rotulo(texto):
                    break  # chegou ao campo seguinte sem passar pelo compromisso
                return texto, True, True
            break
        return "", False, False

    def _extrair_recursos_orcamentarios(self, campo: Campo) -> tuple[str, bool]:
        """Fontes de recurso mais a linha de "Total dos Recursos" que as fecha."""
        localizacao = self._linha_com_rotulo(campo.secao, r"codigoda?fonte")
        if localizacao is None:
            return "", False
        tabela, linha_cabecalho = localizacao

        registros = self._consolidar_linhas(self._linhas_da_tabela(tabela, linha_cabecalho))

        # O total fica em uma linha rotulada, onde o laço acima parou. O rótulo
        # é preservado como está no documento ("Total dos Recursos" ou
        # "Total do Teto", conforme a versão do modelo).
        for numero in range(linha_cabecalho + 1, tabela.total_linhas):
            celulas = [texto for _, texto in tabela.linha_de_celulas(numero)]
            rotulo = next((limpar(t) for t in celulas if casa(_REGEX_TOTAL, t)), "")
            if not rotulo:
                continue
            valores = [
                limpar(texto)
                for texto in celulas
                if limpar(texto) and not casa(_REGEX_TOTAL, texto)
            ]
            if valores:
                registros.append(f"{rotulo}: {valores[-1]}")
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

    def _resposta_sim_ou_nao(self) -> tuple[str, bool]:
        """Resposta da variante "Desagregação territorial/regional (sim ou não?)".

        Parte das fichas troca o par de caixas "Estado / Território de
        Identidade" por uma pergunta única, respondida com uma marca ao lado de
        "Sim" ou de "Não". Sem ler essa resposta, a seção inteira saía vazia —
        e ainda era reportada como rótulo ausente, quando na verdade a ficha
        respondeu que não há desagregação (por isso o resto do bloco está em
        branco).
        """
        achados = [
            (no, _RESPOSTAS_SIM_NAO[chave(no.texto)])
            for no in self.documento.nos_da_secao("TERRITORIAL")
            if no.tipo == "celula" and chave(no.texto) in _RESPOSTAS_SIM_NAO
        ]
        if not achados:
            return "", False
        if len(achados) == 1:
            # Só a opção escolhida ficou no documento.
            return achados[0][1], True

        # As duas opções aparecem: vale a que está marcada, na própria célula
        # ou na vizinha (a marca fica ora antes, ora depois do texto).
        for no, resposta in achados:
            celulas = self.documento.tabelas[no.tabela].linha_de_celulas(no.linha)
            colunas = [coluna for coluna, _ in celulas]
            if no.coluna not in colunas:  # pragma: no cover - grade irregular
                continue
            posicao = colunas.index(no.coluna)
            vizinhas = [
                celulas[j][1] for j in (posicao - 1, posicao + 1) if 0 <= j < len(celulas)
            ]
            if esta_marcado(no.texto) or any(esta_marcado(t) for t in vizinhas):
                return resposta, True
        return achados[0][1], True

    def _extrair_desagregacao_territorial(self, campo: Campo) -> tuple[str, bool]:
        """Consolida toda a seção de desagregação territorial em uma célula."""
        estado, encontrada = self._marcacao(r"estado")
        territorio, _ = self._marcacao(r"territoriode?identidade")

        resposta = ""
        if not encontrada:
            # Modelo sem o par de caixas: a pergunta é uma só.
            resposta, encontrada = self._resposta_sim_ou_nao()

        itens = [
            ("Desagregação territorial/regional", resposta),
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
