# Movimentação consolidada (Python + Excel)

Esqueleto do pipeline que troca "800k+ linhas dentro do Excel com SUMIFS" por
"Python processa uma vez, Excel só lê uma tabela pequena e pronta".

## Como funciona

```
data/raw/*.xlsx (export bruto do SAP, 800k+ linhas)
        |
        v
   etl.py + classify.py   -> aplica category_rules.csv, agrega por
   (pandas)                  Cliente x Mês x Categoria
        |
        v
data/output/*.parquet   (cache rápido, poucos milhares de linhas)
        |
        v
   build_excel.py  ->  data/output/consolidado.xlsx
        - "Painel Cliente": dropdown de cliente + waterfall mensal
          (fórmulas SUMIFS em cima da tabela pequena, instantâneo)
        - "Geral": waterfall consolidado de todos os clientes
        - "Por Cliente": tabela fato (fonte do Painel / de Tabela Dinâmica)
        - "Fato (base Python)": granularidade cliente x mês x categoria
```

O Excel nunca mais abre as 800k linhas brutas -- só a tabela já agregada.

## Passo a passo

1. **Instalar dependências** (uma vez):
   ```
   pip install -r requirements.txt
   ```

2. **Exportar o Movimentação do SAP** (xlsx ou csv, o que o seu SAP permitir)
   e colocar o(s) arquivo(s) em `data/raw/`. Pode ser um arquivo por mês, o
   ETL lê todos juntos. Para `.xlsx` grande (800k+ linhas), o ETL usa o
   engine `calamine` (instalado via `requirements.txt`), que é bem mais
   rápido que o padrão do Excel/openpyxl para leitura.

3. **Calibrar `category_rules.csv`** (só na primeira vez / quando aparecer
   Document Type novo):
   ```
   python etl.py --raw-dir "<sua pasta>" --audit
   ```
   Isso lista cada `Document Type` do export, quantas linhas, a soma em
   `Value Confirmed` **e** em `Value Reserved` lado a lado, e mostra se já
   tem categoria mapeada. Edite o CSV até não sobrar nenhum "SEM REGRA".

   Cada linha do `category_rules.csv` tem 3 colunas:
   ```
   tipo_documento,categoria,coluna_valor
   ZTO,Ressarcimento SAP,valor_confirmado
   ZOR,Reserva,valor_reservado
   ```
   `coluna_valor` diz de onde tirar o valor daquele Document Type --
   `valor_confirmado` (= `Value Confirmed`, padrão se deixar em branco) ou
   `valor_reservado` (= `Value Reserved`). Precisa disso porque alguns tipos
   (ex.: `ZOR`, `ZREA`) vêm com `Value Confirmed` sempre zerado -- o audit
   deixa isso óbvio comparando as duas somas.

   > O `category_rules.csv` já vem calibrado com os Document Type
   > confirmados em produção (`ZTO`, `ZOR`, `ZF2`, `ZG2`, `ZREA`). Se
   > aparecer um tipo novo no seu export, o audit vai mostrar "SEM REGRA" e
   > você só adiciona a linha.

4. **Rodar o pipeline**:
   ```
   python etl.py
   ```
   Gera `data/output/consolidado.xlsx` e os `.parquet` intermediários.

   Se preferir apontar direto pra outra pasta (ex.: uma pasta do OneDrive
   onde você deposita o export), sem mexer em `config.py` nem mover
   arquivo nenhum:
   ```
   python etl.py --raw-dir "C:\Users\I0507867.FARMA\OneDrive - Sanofi\Desktop\OL Robo" --audit
   python etl.py --raw-dir "C:\Users\I0507867.FARMA\OneDrive - Sanofi\Desktop\OL Robo"
   ```
   Por padrão o `.xlsx` sai em `data/output/`; use `--output-dir` pra mudar
   isso também.

5. Abrir `consolidado.xlsx`, escolher o cliente no dropdown da aba
   **Painel Cliente** e pronto.

## Saldo Inicial (âncora de um mês específico)

Por padrão, o Saldo Inicial do primeiro mês de cada cliente começa em 0,00
(o ETL só conhece o que está no export de Movimentação). Pra ancorar isso
num saldo real, coloque na mesma pasta um arquivo cujo nome contenha "Saldo
Inicial" (ex.: `Saldo Inicial 01-05-2026.xlsx`) -- o ETL detecta esse nome
automaticamente e não mistura com os arquivos de Movimentação.

Esse arquivo usa **o mesmo formato de colunas do export de Movimentação**
(`Root CNPJ`, `Client Name`, `Reimbursement Load`, etc.), mas o ETL soma só
a coluna `Reimbursement Load` por cliente -- não classifica por Document
Type. Rode informando a que mês esse saldo se refere:
```
python etl.py --raw-dir "<sua pasta>" --saldo-inicial-data 2026-05-01
```
Isso substitui o Saldo Inicial calculado de maio/26 (de todo cliente que
aparecer no arquivo) pelo valor do arquivo, e os meses seguintes continuam
encadeando normalmente a partir daí. Meses anteriores (ex.: abr/26) não são
afetados. Se um cliente do arquivo de Saldo Inicial não tiver nenhum
movimento em maio/26, o ETL ainda cria a linha dele nesse mês (senão o saldo
"sumiria" até ele voltar a ter movimentação).

Sem `--saldo-inicial-data`, se houver um arquivo de Saldo Inicial na pasta o
ETL para com um erro claro pedindo a data -- pra nunca aplicar a âncora no
mês errado por engano.

## Testar com dados de exemplo (sem precisar do export real)

```
python sample_data/gerar_exemplo.py
cp sample_data/exemplo_movimentacao.xlsx data/raw/    # ou .csv, tanto faz
python etl.py --audit
python etl.py
```

## Próximos passos possíveis (não incluídos neste esqueleto)

- **Botão de atualização dentro da planilha**: dá pra plugar `xlwings` e
  criar um botão "Atualizar" que roda `etl.py` e regrava as abas sem sair do
  Excel. Só vale a pena se for você (ou poucas pessoas) abrindo o arquivo,
  porque exige Python instalado em quem clica o botão.
- **Agendamento**: rodar `etl.py` automaticamente (cron, Agendador de
  Tarefas) toda vez que um novo export cair em `data/raw/`.
- **Histórico**: hoje `etl.py` reprocessa tudo que está em `data/raw/` do
  zero a cada rodada. Se o volume mensal crescer muito, dá pra particionar o
  parquet por mês e só reprocessar o mês novo.
- **Multiusuário sem Python**: se o time inteiro precisa abrir o arquivo sem
  ter Python instalado, gerar o `.xlsx` num lugar central (seu PC / servidor
  agendado) e distribuir o arquivo pronto, em vez de xlwings.
