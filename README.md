# Leitor para o PPA — fichas .docx → planilhas Excel

Aplicação (com **janela gráfica** e também linha de comando) que varre uma pasta com **centenas/milhares** de
fichas de planejamento governamental em Word (`.docx`), **identifica
automaticamente o tipo de cada documento** e consolida tudo em **três planilhas
Excel**, com **uma linha por problema vinculado** (na prática, uma linha por
ficha: quase toda ficha tem um único problema):

| Tipo do documento | Reconhecido por | Planilha gerada |
|---|---|---|
| Indicador de Compromisso | seção `VÍNCULO DO INDICADOR DE COMPROMISSO` | `Indicadores.xlsx` |
| Iniciativa | seção `VÍNCULO DA INICIATIVA` | `Iniciativas.xlsx` |
| Ficha de Controle | `NOME DO DIRETÓRIO DO COMPROMISSO` | `Controles.xlsx` |

Os dois tipos podem estar misturados na mesma pasta. Campos com vários itens
(ex.: *Causa(s) Crítica(s)*, *Ação(ões) Crítica(s)*, propostas de escuta social,
ações orçamentárias) são unidos em um único texto, ficando confinados na célula
da respectiva coluna.

---

## 1. O que o aplicativo faz

Veja a tabela acima: aponte a pasta das fichas, receba as duas planilhas.
Escolha abaixo como quer usar.

## 2. Baixar o executável pronto (não precisa de Python)

Para quem só quer usar o aplicativo, há um executável para Windows gerado
automaticamente a cada atualização do código:

**[Baixar LeitorPPA.exe](https://github.com/LMPerote/Leitor-para-o-PPA/releases/latest/download/LeitorPPA.exe)**

Baixe, dê dois cliques e use — sem instalar Python, sem linha de comando. Esse
link é fixo: aponta sempre para a versão mais recente. É o jeito mais simples de
distribuir o aplicativo para outras pessoas.

> A primeira abertura demora alguns segundos (o programa se descompacta em uma
> pasta temporária); as seguintes são mais rápidas. O Windows pode exibir um
> aviso do SmartScreen por ser um executável sem assinatura digital: clique em
> *Mais informações* → *Executar assim mesmo*.

Para gerar o executável você mesmo, dê dois cliques em `Gerar executavel.bat`
(ele instala o PyInstaller e grava o resultado em `dist\LeitorPPA.exe`).

## 3. Preparando a pasta e instalando (a partir do código)

Requer **Python 3.10 ou superior** ([python.org/downloads](https://www.python.org/downloads/)
— no instalador, marque *Add python.exe to PATH*).

No Windows, **não é preciso instalar mais nada**: o `Extrair PPA.bat` instala as
dependências na primeira execução. Para instalar manualmente (ou fora do
Windows):

```bash
pip install -r requirements.txt
```

Coloque **todos os `.docx`** (Indicadores e Iniciativas, misturados) em uma
pasta — subpastas são varridas por padrão. Não é preciso separar, renomear nem
ordenar nada: arquivos temporários do Word (`~$ficha.docx`) e arquivos de outras
extensões são ignorados automaticamente.

```
documentos/
├── E1_..._Indicador_1.docx
├── E1_..._Iniciativa_1.docx
└── eixo-2/
    └── E2_..._Indicador_7.docx
```

> Apenas `.docx` é suportado (Word 2007+). Arquivos `.doc` antigos precisam ser
> convertidos antes — no Word: *Salvar como → .docx*; em lote no Linux:
> `libreoffice --headless --convert-to docx *.doc`.

## 4. Rodando com dois cliques (janela gráfica)

Dê **dois cliques em `Extrair PPA.bat`**. Na primeira vez ele instala as
dependências sozinho; depois abre a janela do aplicativo. Se algo impedir a
abertura, ele mostra o motivo em uma janela de texto e espera — nunca falha em
silêncio.

A janela tem:

* **Documentos (.docx)** → *Procurar...* e selecione a pasta das fichas;
* **Salvar planilhas em** → *Procurar...* e escolha onde gravar;
* clique em **Extrair** e acompanhe a barra de progresso;
* ao final, o relatório aparece na janela e o botão **Abrir pasta das
  planilhas** leva direto ao resultado.

As pastas escolhidas ficam memorizadas (em `preferencias.json`) para a próxima
vez. A opção **Testar com os 10 primeiros arquivos** é o jeito rápido de
conferir o resultado antes de rodar o acervo inteiro.

> Para criar um atalho na Área de Trabalho: clique com o botão direito em
> `Extrair PPA.bat` → *Enviar para* → *Área de trabalho (criar atalho)*.

No macOS/Linux, a mesma janela abre com `python interface.py`.

## 5. Rodando pelo terminal

```bash
python extrair_indicadores.py -e ./documentos -s ./saida
```

Isso grava `saida/Indicadores.xlsx` e `saida/Iniciativas.xlsx` e imprime o
resumo do lote:

```
==================================================================
RELATÓRIO DE EXTRAÇÃO
==================================================================
Arquivos encontrados .......: 152
75 arquivos de Indicador de Compromisso processados.
  - planilha: /.../saida/Indicadores.xlsx
75 arquivos de Iniciativa processados.
  - planilha: /.../saida/Iniciativas.xlsx
2 arquivos ignorados ou com erro.
  - corrompido.docx: erro de leitura (PackageNotFoundError: ...)
  - memorando.docx: tipo não reconhecido (sem seção de vínculo conhecida)
Tempo total ................: 4.6s
==================================================================
```

### Opções

| Opção | Descrição |
|---|---|
| `-e`, `--entrada` | Pasta com os arquivos `.docx` (obrigatória). |
| `-s`, `--saida` | Pasta onde gravar as planilhas (padrão: pasta atual). |
| `--sem-recursao` | Não varre subpastas. |
| `--separador` | Como unir itens de listas na mesma célula: `quebra` (padrão), `ponto-virgula`, `barra` ou um texto literal. |
| `--csv` | Gera também um `.csv` de cada planilha (UTF-8 com BOM, separador `;`). |
| `--limite N` | Processa apenas os N primeiros arquivos (teste rápido). |
| `--log ARQUIVO` | Grava o log detalhado em arquivo. |
| `-v`, `--verboso` | Mostra mensagens de depuração. |
| `--inspecionar FICHA.docx` | Diagnostica um documento (tipo detectado, seções, rótulos) e encerra. |

Exemplos:

```bash
# Teste rápido com 10 arquivos, log em disco e CSV extra
python extrair_indicadores.py -e ./documentos -s ./saida --limite 10 --csv --log ./extracao.log

# Listas separadas por ponto e vírgula em vez de quebra de linha
python extrair_indicadores.py -e ./documentos -s ./saida --separador ponto-virgula
```

## 6. As planilhas geradas

Cada planilha tem a aba de dados e a aba **Dicionário**,
que descreve cada coluna. As duas são sempre geradas, mesmo que um dos tipos não
apareça na pasta. Formatação aplicada: cabeçalho em **negrito** com fundo azul
escuro e texto branco, linha de cabeçalho congelada, filtro automático, largura de colunas
ajustada e células com **quebra automática de texto** e alinhamento superior.

### Uma linha por problema

A granularidade da planilha é o **problema vinculado ao compromisso**: se uma
ficha trouxer mais de um problema, ela gera uma linha para cada um, repetindo
todas as demais colunas. Ficha com um problema — o caso normal — continua
gerando uma linha só, e nenhum arquivo deixa de aparecer na planilha, mesmo sem
problema preenchido.

O relatório final avisa quando isso acontece (`N linhas na planilha`), e a
coluna `Nome_Arquivo` repetida deixa a duplicação visível. Atenção ao contar:
para saber quantas **fichas** existem, conte os valores distintos de
`Nome_Arquivo`, não o número de linhas.

### `Indicadores.xlsx`

`Nome_Arquivo`, `Caminho_Relativo`, `Eixo`, `Programa`, `Compromisso`,
`Problemas_Vinculados`, `Causas_Criticas`, `Atributos_Descricao`,
`Formula_Calculo`, `Memoria_Calculo`, `Unidade_Medida`, `Valor_Referencia`,
`Ano_Referencia`, `Valor_Meta`, `Periodicidade_Apuracao`, `Polaridade`,
`Classificacao`, `Fonte`, `Meios_Verificacao`, `Responsavel_Sigla_Orgao`,
`Responsavel_UO`, `Responsavel_USP`, `Desagregacao_Territorial`,
`Info_Complementares_Objetivo`, `Limitacoes_Indicador`, `Fragilidades_Apuracao`,
`Limitacoes_Meta_Operacionais`, `Limitacoes_Meta_Orcamentarias`,
`Limitacoes_Meta_Institucionais`, `Possibilidade_Desagregacao_Populacional`,
`Programas_Especiais`, `Status`, `Campos_Nao_Encontrados`, `Observacoes`.

`Desagregacao_Territorial` consolida toda a seção em uma célula, item por linha:

```
Estado: Sim
Território de Identidade: Não
Fórmula de cálculo Territorial: Não se aplica
Unidade de Medida: Não se aplica
Memória de Cálculo: Não se aplica
Outras possibilidades de Regionalização: Não se aplica
```

Quando a ficha lista territórios, entram também `Territórios de Identidade`,
`Memória de Cálculo Territorial` e `Meta Territorial`.

### `Controles.xlsx`

A ficha de controle é a capa da pasta do compromisso: registra quem digitou no
Fiplan, quando, quantas fichas de cada tipo existem e o que ficou pendente.

`Nome_Arquivo`, `Caminho_Relativo`, `Nome_Diretorio_Compromisso`, `Eixo`,
`Programa`, `Compromisso`, `Nome_Digitador_Fiplan`, `Data_Insercao_Fiplan`,
`Qtd_Indicadores_Compromisso`, `Qtd_Fichas_Iniciativas`,
`Pendencias_Observacoes`, `Status`, `Campos_Nao_Encontrados`, `Observacoes`.

### `Iniciativas.xlsx`

`Nome_Arquivo`, `Caminho_Relativo`, `Eixo`, `Programa`, `Compromisso`,
`Problemas_Vinculados`, `Causas_Criticas`, `Acoes_Criticas`,
`Propostas_Escuta_Social`, `Atributos_Descricao`, `Entregas_Vinculadas`,
`Responsavel_Sigla_Orgao`, `Responsavel_UO`, `Responsavel_USP`,
`Orgaos_Parceiros`, `Fatores_Criticos_Operacionais`,
`Fatores_Criticos_Orcamentarios`, `Fatores_Criticos_Institucionais`,
`Recursos_Orcamentarios`, `Indicador_Compromisso_Vinculado`,
`Indicadores_Sensibilizados`, `Acoes_Orcamentarias_Vinculadas`,
`Produtos_Acoes_Orcamentarias`, `Status`, `Campos_Nao_Encontrados`,
`Observacoes`.

As colunas vindas de tabelas trazem uma linha por registro, com os cabeçalhos
originais preservados. Por exemplo, `Recursos_Orcamentarios`:

```
Código da Fonte: 100 | Nome da Fonte: Tesouro Estadual | Montante em R$: R$ 2.400.000,00
Total dos Recursos: R$ 2.400.000,00
```

### Colunas de auditoria

* `Nome_Arquivo` / `Caminho_Relativo` — rastreabilidade até o documento de origem.
* `Status` — `OK` (todos os rótulos localizados) ou `OK_COM_PENDENCIAS`.
* `Campos_Nao_Encontrados` — colunas cujo rótulo não existe naquele documento.
* `Observacoes` — avisos do processamento.

Arquivos corrompidos ou de tipo não reconhecido **não entram nas planilhas**:
são listados no relatório final e no log, e **nunca interrompem o lote**.

## 7. Como o parser funciona

1. **Leitura estrutural** (`extrator/documento.py`): parágrafos e células são
   lidos na ordem do documento; tabelas viram uma grade que trata **células
   mescladas**, e o texto preserva quebras manuais (`<w:br/>`).
2. **Detecção de seção**: os cabeçalhos (`VÍNCULO...`, `ATRIBUTOS...`,
   `DESAGREGAÇÃO TERRITORIAL`, `Fatores Críticos...`, `Etapa II...`) delimitam o
   escopo da busca, evitando confundir rótulos repetidos — por exemplo, as três
   ocorrências de *Memória de Cálculo* na ficha de Indicador, ou o *UO* que
   aparece tanto no responsável quanto nas ações orçamentárias.
3. **Classificação** (`extrator/parser.py::classificar`): vence o marcador de
   vínculo que aparecer mais no topo do documento.
4. **Casamento de rótulos por chave canônica**: o texto é reduzido a
   minúsculas, sem acentos, espaços ou pontuação. Assim
   `Indicador(es) doPrograma Sensibilizado(s)` (erro de digitação presente no
   modelo real) casa com o rótulo esperado. Uma segunda chave descarta
   ornamentos das variantes do modelo — a numeração das seções
   (`Bloco 1: VÍNCULO...`) e as explicações entre parênteses
   (`Compromisso (Objetivo do Compromisso)`).
5. **Resolução do valor**, nesta ordem: célula com apenas o rótulo (valor
   **à direita** ou **abaixo**) → `Rótulo: valor` na mesma célula → rótulo na
   primeira linha com o valor nas linhas seguintes. Isso cobre os layouts vertical e horizontal sem depender da
   posição fixa das tabelas.
6. **Extratores dedicados** para o que foge desse padrão: caixas de seleção
   (`Estado [x] / Território de Identidade [ ]` → `Sim`/`Não`) e tabelas com
   cabeçalho (territórios, programas especiais, escuta social, recursos, ações
   e produtos orçamentários).

### Adaptando a novos modelos de ficha

Na maioria dos casos basta editar **`extrator/modelos/indicador.py`** ou
**`extrator/modelos/iniciativa.py`** — acrescentar um `Campo` com o nome da
coluna e o padrão do rótulo. Para um terceiro tipo de documento, crie um módulo
novo com seu `Modelo`, registre-o em `extrator/modelos/__init__.py` (lista
`MODELOS`) e acrescente ali as seções dele.

Para descobrir os rótulos exatos de um documento novo:

```bash
python extrair_indicadores.py --inspecionar ./documentos/ficha.docx
```

A saída mostra o tipo detectado e, para cada célula, a seção, se ela foi
reconhecida como `RÓTULO` ou `valor`, e sua posição na tabela.

## 8. Testes

```bash
pip install pytest
pytest -q          # 75 testes: classificação, os dois modelos, erros e Excel
```

## 9. Observações operacionais

* Desempenho de referência: ~35 documentos/segundo (≈2.000/minuto).
* Textos acima do limite do Excel (32.767 caracteres por célula) são truncados
  com a marca `... [texto truncado]`.
* Feche as planilhas de saída antes de reexecutar: com o arquivo aberto, o Excel
  bloqueia a gravação no Windows.
* `--separador` afeta a união de **itens de lista**. Com um separador diferente
  de quebra de linha, as quebras internas de uma célula do Word (várias causas
  na mesma célula) também são convertidas.

## 10. Estrutura do projeto

```
Leitor-para-o-PPA/
├── .github/workflows/         # build automático do .exe a cada push
├── Extrair PPA.bat            # atalho de dois cliques (Windows)
├── Gerar executavel.bat       # gera o LeitorPPA.exe nesta máquina
├── interface.py               # janela gráfica (Tkinter)
├── extrair_indicadores.py     # CLI: argumentos, progresso, log e relatório
├── extrator/
│   ├── aplicacao.py           # fluxo completo, usado pela janela e pelo CLI
│   ├── relatorio.py           # texto do relatório final
│   ├── modelos/
│   │   ├── base.py            # Campo, Modelo e colunas de auditoria
│   │   ├── indicador.py       # mapa de campos do Indicador de Compromisso
│   │   ├── iniciativa.py      # mapa de campos da Iniciativa
│   │   ├── controle.py        # mapa de campos da Ficha de Controle
│   │   └── __init__.py        # registro dos modelos, seções e rótulos
│   ├── documento.py           # leitura do .docx (parágrafos, tabelas, merges)
│   ├── parser.py              # classificação e motor rótulo -> valor
│   ├── pipeline.py            # processamento em lote e estatísticas
│   ├── planilha.py            # geração e formatação dos .xlsx/.csv
│   └── texto.py               # normalização de texto e chaves canônicas
├── testes/test_extrator.py
└── requirements.txt
```
