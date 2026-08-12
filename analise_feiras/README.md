# Análise de Feiras / Campanhas — pipeline recorrente

Automação em Python para identificar pedidos faturados fora da vigência de
feiras/campanhas comerciais ("erro operacional") e calcular qual seria o
desconto e o preço líquido corretos.

## O que o script faz

Para cada **matriz** ativa no `config.yaml` (Feira, Campanha, ou outra que você
cadastrar):

1. **Filtra** a base grande de pedidos pela palavra-chave da matriz (ex.:
   `"FEIRA"`) dentro da coluna `Tabela de negociação`, e salva esse recorte
   limpo em `saida/<Matriz>_filtrado.xlsx`.
2. **Traz (tipo PROCX)** do arquivo de controle da matriz (ex.:
   `Controle_Feiras.xlsx`, aba `dados`) os campos `Feira`, `Data inicial` e
   `Vigência`.
3. **Compara** a data do pedido com o intervalo `[Data inicial, Vigência]` e
   preenche a coluna `Check`:
   - `OK` — pedido dentro da vigência
   - `ERRO OPERACIONAL - pedido antes da vigência`
   - `ERRO OPERACIONAL - pedido após a vigência`
   - `SEM CADASTRO NA MATRIZ` — a tabela de negociação do pedido não foi
     encontrada no arquivo de controle
4. Para as linhas marcadas como erro operacional, **cruza com
   `Condicao_comercial.xlsx`** (por CNPJ + EAN) para achar o desconto que
   deveria ter sido aplicado, e calcula:
   - `preco_sem_desconto` = Faturado líquido / (1 − desconto aplicado)
   - `preco_liquido_desconto_correto` = preço sem desconto × (1 − desconto correto)
   - `diferenca_faturamento` = faturado líquido − preço líquido correto

O resultado final de cada matriz é salvo em `saida/<Matriz>_analise.xlsx`, e o
console mostra um resumo (contagem de `Check` e impacto financeiro total).

## Como rodar

```bash
cd analise_feiras
python -m venv .venv && source .venv/bin/activate   # opcional
pip install -r requirements.txt

# 1. Coloque os arquivos em dados/:
#    - dados/base_pedidos.xlsx      (a base grande de pedidos)
#    - dados/Controle_Feiras.xlsx   (aba "dados")
#    - dados/Condicao_comercial.xlsx (aba "dados")

# 2. Rode:
python pipeline.py
```

Saída vai para `analise_feiras/saida/`.

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
  chave_controle: "Tabela de negociação"
  colunas_trazidas:
    feira: "Campanha"
    data_inicial: "Data inicial"
    vigencia: "Vigência"
```

O arquivo de controle precisa ter, na aba indicada, uma coluna que bata com
`Tabela de negociação` da base (o "PROCX"), mais as colunas de nome, data
inicial e vigência.

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
validando o filtro, o PROCX de vigência, o `Check` e o cálculo de desconto.

## Observações sobre performance em bases grandes

- O script usa `pandas` com leitura de Excel (`openpyxl`), que já é
  suficiente para a maioria das bases de algumas centenas de milhares de
  linhas.
- Se a base ultrapassar milhões de linhas e a leitura ficar lenta, salve a
  base como `.csv` (o script lê `.csv`/`.tsv` automaticamente pela extensão)
  ou considere trocar `pandas` por `polars` — a estrutura do pipeline
  (filtrar → merge → calcular) é a mesma, só a biblioteca muda.
