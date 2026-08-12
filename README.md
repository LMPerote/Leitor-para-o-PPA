# Extrator de Indicadores de Compromisso (.docx → .xlsx)

Aplicação de linha de comando que varre uma pasta com **centenas/milhares** de
fichas de Indicador de Compromisso em Word (`.docx`) e consolida tudo em **uma
única planilha Excel**, onde **cada arquivo Word vira exatamente uma linha**.

Campos com vários itens (ex.: *Causa(s) Crítica(s)*, *Territórios de
Identidade*) são unidos em um único texto, ficando confinados na célula da
respectiva coluna.

---

## 1. Instalação

Requer **Python 3.10 ou superior**.

```bash
# 1) (recomendado) ambiente virtual
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate

# 2) dependências
pip install -r requirements.txt
```

## 2. Uso

```bash
# Processa todos os .docx da pasta (inclusive subpastas)
python extrair_indicadores.py -e ./documentos -s ./indicadores.xlsx
```

Ao final é exibido o relatório do lote:

```
==============================================================
RELATÓRIO DE EXTRAÇÃO
==============================================================
Arquivos encontrados .......: 500
Processados com sucesso ....: 500
  - completos ..............: 498
  - com campos pendentes ...: 2
Falhas de leitura ..........: 0
Tempo total ................: 16.5s
Planilha gerada ............: /caminho/indicadores.xlsx
==============================================================
```

### Opções

| Opção | Descrição |
|---|---|
| `-e`, `--entrada` | Pasta com os arquivos `.docx` (obrigatória). |
| `-s`, `--saida` | Arquivo `.xlsx` de saída (padrão: `indicadores_consolidados.xlsx`). |
| `--sem-recursao` | Não varre subpastas. |
| `--separador` | Como unir listas na mesma célula: `quebra` (padrão), `ponto-virgula`, `barra` ou um texto literal. |
| `--csv` | Gera também um `.csv` (UTF-8 com BOM, separador `;`). |
| `--limite N` | Processa apenas os N primeiros arquivos (teste rápido). |
| `--log ARQUIVO` | Grava o log detalhado em arquivo. |
| `-v`, `--verboso` | Mostra mensagens de depuração. |
| `--inspecionar FICHA.docx` | Diagnostica a estrutura de um documento e encerra. |

Exemplos:

```bash
# Teste rápido com 10 arquivos, log em disco e CSV extra
python extrair_indicadores.py -e ./documentos --limite 10 --csv --log ./extracao.log

# Listas separadas por ponto e vírgula em vez de quebra de linha
python extrair_indicadores.py -e ./documentos --separador ponto-virgula
```

## 3. A planilha gerada

**Aba `Indicadores`** — uma linha por arquivo, com cabeçalho em negrito,
congelamento de painéis, filtro automático, largura de colunas ajustada e
células com alinhamento superior e quebra automática de texto (`wrap_text`).

Colunas, na ordem:

| Grupo | Colunas |
|---|---|
| Auditoria | `Nome_do_Arquivo`, `Caminho_Relativo` |
| Vínculo | `Eixo`, `Programa`, `Compromisso`, `Problemas_Vinculados`, `Causas_Criticas` |
| Atributos | `Descricao`, `Formula_de_Calculo`, `Memoria_de_Calculo`, `Unidade_de_Medida`, `Valor_de_Referencia`, `Ano_de_Referencia`, `Valor_da_Meta`, `Periodicidade_da_Apuracao`, `Polaridade`, `Classificacao`, `Fonte`, `Meios_de_Verificacao`, `Responsavel_Sigla_Orgao`, `Responsavel_UO`, `Responsavel_USP` |
| Desagregação territorial | `Desagregacao_Estado`, `Desagregacao_Territorio_Identidade`, `Formula_Calculo_Territorial`, `Unidade_Medida_Territorial`, `Memoria_Calculo_Territorial`, `Territorios_de_Identidade`, `Memoria_Calculo_Por_Territorio`, `Metas_Territoriais`, `Outras_Possibilidades_Regionalizacao` |
| Informações complementares | `Objetivo_Interpretacao_Uso`, `Limitacoes_do_Indicador`, `Fragilidades_para_Apuracao`, `Limitacoes_Meta_Operacionais`, `Limitacoes_Meta_Orcamentarias`, `Limitacoes_Meta_Institucionais`, `Possibilidade_Desagregacao_Populacional`, `Programas_Especiais`, `Indicadores_Programa_Sensibilizado` |
| Auditoria | `Status`, `Campos_Nao_Encontrados`, `Observacoes` |

Valores de `Status`:

* `OK` — todos os rótulos foram localizados;
* `OK_COM_PENDENCIAS` — o arquivo foi lido, mas algum rótulo não existe nele
  (as colunas afetadas ficam em `Campos_Nao_Encontrados`);
* `ERRO_DE_LEITURA` — arquivo corrompido ou ilegível; a linha é mantida com os
  campos vazios e a causa em `Observacoes`. **O lote nunca é interrompido.**

**Aba `Dicionário`** — descrição de cada coluna e a seção da ficha de origem.

## 4. Como o parser funciona

1. **Leitura estrutural** (`extrator/documento.py`): parágrafos e células são
   lidos na ordem do documento; tabelas viram uma grade que trata **células
   mescladas**, e o texto preserva quebras manuais (`<w:br/>`).
2. **Detecção de seção**: os cabeçalhos (`VÍNCULO...`, `ATRIBUTOS...`,
   `DESAGREGAÇÃO TERRITORIAL`, `INFORMAÇÕES COMPLEMENTARES`) delimitam o escopo
   da busca, evitando confundir rótulos repetidos — por exemplo, as três
   ocorrências de *Memória de Cálculo* na ficha.
3. **Casamento de rótulos por chave canônica**: o texto é reduzido a
   minúsculas, sem acentos, espaços ou pontuação. Assim
   `Indicador(es) doPrograma Sensibilizado(s)` (erro de digitação presente no
   modelo real) casa com o rótulo esperado.
4. **Resolução do valor**, nesta ordem: valor após `:` na própria célula →
   primeira célula **à direita** que não seja outro rótulo → primeira célula
   **abaixo**. Isso cobre os layouts vertical e horizontal sem depender da
   posição fixa das tabelas.
5. **Extratores dedicados** para o que foge desse padrão: caixas de seleção
   (`Estado [x] / Território de Identidade [ ]` → `Sim`/`Não`), tabela de
   territórios e tabela de Programas Especiais.

### Adaptando a novos modelos de ficha

Na maioria dos casos basta editar **`extrator/campos.py`** — adicionar um
`Campo` com o nome da coluna e o padrão do rótulo. Para descobrir os rótulos
exatos de um documento novo:

```bash
python extrair_indicadores.py --inspecionar ./documentos/ficha.docx
```

A saída mostra, para cada célula, a seção, se ela foi reconhecida como
`RÓTULO` ou `valor`, e sua posição na tabela.

## 5. Testes

```bash
pip install pytest
pytest -q          # 45 testes: layouts, listas, erros e formatação do Excel
```

## 6. Observações operacionais

* Apenas `.docx` é suportado (formato do Word 2007+). Arquivos `.doc` antigos
  precisam ser convertidos antes — no Word: *Salvar como → .docx*; em lote no
  Linux: `libreoffice --headless --convert-to docx *.doc`.
* Arquivos temporários do Word (`~$ficha.docx`) são ignorados automaticamente.
* Desempenho de referência: ~30 documentos/segundo (≈1.800/minuto).
* Textos acima do limite do Excel (32.767 caracteres por célula) são truncados
  com a marca `... [texto truncado]`.
* Feche a planilha de saída antes de reexecutar: com o arquivo aberto, o Excel
  bloqueia a gravação no Windows.

## 7. Estrutura do projeto

```
Leitor-para-o-PPA/
├── extrair_indicadores.py     # CLI: argumentos, progresso, log e relatório
├── extrator/
│   ├── campos.py              # mapa dos campos -> colunas do Excel
│   ├── documento.py           # leitura do .docx (parágrafos, tabelas, merges)
│   ├── parser.py              # motor de extração rótulo -> valor
│   ├── pipeline.py            # processamento em lote e estatísticas
│   ├── planilha.py            # geração e formatação do .xlsx/.csv
│   └── texto.py               # normalização de texto e chaves canônicas
├── testes/test_extrator.py
└── requirements.txt
```
