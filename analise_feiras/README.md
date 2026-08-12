# Análise de Feiras / Campanhas — pipeline recorrente

Automação em Python para identificar pedidos faturados fora da vigência de
feiras/campanhas comerciais ("erro operacional") e calcular qual seria o
desconto e o preço líquido corretos.

## O que o script faz

Para cada **matriz** ativa no `config.yaml` (Feira, Campanha, ou outra que você
cadastrar):

1. **Filtra** a base grande de pedidos pela palavra-chave da matriz (ex.:
   `"FEIRA"`) dentro da coluna `Tabela de negociação`.
2. **Verifica (tipo PROCX)** se o CNPJ do pedido existe na coluna de CNPJ
   ajustado do arquivo de controle da matriz (ex.: `Controle_Feiras.xlsx`,
   coluna `CNPJ Ajustado` — a lista de quem realmente participou da feira).
3. Preenche a coluna `Check`:
   - `OK` — o CNPJ do pedido está na lista de participantes; as colunas
     `inicio_real` / `termino_real` são trazidas do controle com as datas
     reais em que a feira aconteceu.
   - `Erro Operacional` — o CNPJ do pedido **não** está na lista de
     participantes (comprou com preço de feira sem ter ido/estar cadastrado).
4. Para as linhas marcadas como `Erro Operacional`, **cruza com
   `Condicao_comercial.xlsx`** (por CNPJ + EAN) para achar o desconto que
   deveria ter sido aplicado, e calcula:
   - `preco_sem_desconto` = Faturado líquido / (1 − desconto aplicado)
   - `preco_liquido_desconto_correto` = preço sem desconto × (1 − desconto correto)
   - `diferenca_faturamento` = faturado líquido − preço líquido correto

O resultado de cada matriz sai num **único arquivo**:
`saida/<Matriz>_analise.xlsx`. O console mostra um resumo (contagem de
`Check` e impacto financeiro total).

## Como rodar

O `config.yaml` já vem apontando para caminhos absolutos dentro de:

```
C:\Users\I0507867.FARMA\OneDrive - Sanofi\Desktop\Analise Painel
```

Não importa de onde você execute o `pipeline.py` — os caminhos são fixos
nessa pasta. Estrutura esperada:

```
Analise Painel\
├── dados\
│   ├── planilha_base_vendas_*.csv  (um ou vários arquivos com esse prefixo —
│   │                                 o script lê e junta todos automaticamente)
│   ├── Controle_Feiras.xlsx        (aba "dados")
│   ├── Controle_Campanhas.xlsx     (aba "dados", se for usar a matriz Campanha)
│   └── Condicao_comercial.xlsx     (aba "dados")
└── saida\                          (criada automaticamente pelo script)
```

A base pode estar dividida em vários arquivos (ex.: um por mês/exportação),
desde que todos comecem com `planilha_base_vendas_`. O `pipeline.py` lê e
empilha automaticamente todos os arquivos que baterem com esse padrão — não
precisa juntar manualmente antes. Cada CSV tem separador (`;` ou `,`) e
encoding (utf-8/latin1/cp1252) detectados automaticamente, arquivo por
arquivo — não precisa que todos estejam salvos do mesmo jeito.

```bash
pip install -r requirements.txt

python pipeline.py
```

Se a pasta `Analise Painel` mudar de lugar, atualize o trecho
`C:/Users/I0507867.FARMA/OneDrive - Sanofi/Desktop/Analise Painel` em todos os
caminhos do `config.yaml` (find & replace no editor de texto).

## Selecionar quais matrizes analisar (recorrência)

Não precisa mexer no código para escolher o que analisar em cada rodada:

- **Editando o `config.yaml`**: marque `ativo: true`/`false` em cada item de
  `matrizes`. O script roda só as marcadas como `true`.
- **Por linha de comando**, sem editar o arquivo:
  ```bash
  python pipeline.py --matrizes Feira            # roda só "Feira"
  python pipeline.py --matrizes Feira,Campanha   # roda as duas
  python pipeline.py --listar                    # lista as matrizes cadastradas
  ```

## Adicionar uma nova matriz (ex.: "Campanha", "Convênio", etc.)

Copie um bloco em `matrizes:` no `config.yaml` e ajuste:

```yaml
- nome: "Campanha"
  ativo: true
  palavra_chave: "CAMPANHA"
  arquivo_controle: "dados/Controle_Campanhas.xlsx"
  aba_controle: "dados"
  chave_controle: "CNPJ Ajustado"
  colunas_trazidas:
    inicio_real: "Início Real"
    termino_real: "Término Real"
  colunas_data: ["inicio_real", "termino_real"]
```

O arquivo de controle precisa ter, na aba indicada, uma coluna com o CNPJ
ajustado (usada para o PROCX contra o CNPJ da base) mais as colunas listadas
em `colunas_trazidas`. Você pode trazer quantas colunas extras quiser — não
precisa ser só datas; qualquer coluna adicionada em `colunas_trazidas`
aparece no arquivo de saída.

## Adaptando nomes de coluna

Se os nomes de coluna dos seus arquivos reais forem diferentes dos usados
aqui, **não edite `pipeline.py`** — ajuste apenas os nomes à direita em
`config.yaml` (em `base.colunas`, `condicao_comercial.colunas` e
`matrizes[].colunas_trazidas`). O script avisa com uma mensagem clara se
alguma coluna configurada não existir na planilha.

## Testando sem os arquivos reais

```bash
pip install -r requirements.txt
pytest tests/ -v
```

O teste `tests/test_pipeline.py` gera planilhas de exemplo (mesma estrutura
da base real) num diretório temporário e roda o pipeline de ponta a ponta,
validando o filtro, o PROCX por CNPJ, o `Check` e o cálculo de desconto.

## Observações sobre performance em bases grandes

- O script usa `pandas` com leitura de Excel (`openpyxl`), que já é
  suficiente para a maioria das bases de algumas centenas de milhares de
  linhas.
- Se a base ultrapassar milhões de linhas e a leitura ficar lenta, salve a
  base como `.csv` (o script lê `.csv`/`.tsv` automaticamente pela extensão)
  ou considere trocar `pandas` por `polars` — a estrutura do pipeline
  (filtrar → merge → calcular) é a mesma, só a biblioteca muda.
