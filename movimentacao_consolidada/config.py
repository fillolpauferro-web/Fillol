from pathlib import Path

BASE_DIR = Path(__file__).parent

# Pasta onde ficam os exports brutos do SAP (.xlsx ou .csv). Um ou vários
# arquivos, inclusive um por mês -- o ETL lê todos juntos.
RAW_DIR = BASE_DIR / "data" / "raw"

# Onde o resultado processado é gravado.
OUTPUT_DIR = BASE_DIR / "data" / "output"
EXCEL_OUTPUT = OUTPUT_DIR / "consolidado.xlsx"

# Nome das colunas no export bruto do SAP -> nome interno usado no ETL.
# Ajuste aqui se o cabeçalho do seu export vier diferente.
COLUMN_MAP = {
    "Client Name": "cliente",
    "Root CNPJ": "cnpj_raiz",
    "Transaction Date": "data",
    "Document Type": "tipo_documento",
    "Value Confirmed": "valor_confirmado",
    "Value Reserved": "valor_reservado",
}

# Colunas que precisam ser lidas como texto (não número) pra não perder zero
# à esquerda -- CNPJ raiz sempre tem 8 dígitos, e vira número errado se o
# pandas/Excel inferir tipo numérico.
DTYPE_OVERRIDES = {"Root CNPJ": str}

# Regras de categoria (Document Type -> categoria do waterfall) ficam em
# category_rules.csv, e não aqui, porque é o dado que você (que conhece a
# regra de negócio) precisa calibrar e revisar com frequência.
CATEGORY_RULES_FILE = BASE_DIR / "category_rules.csv"
