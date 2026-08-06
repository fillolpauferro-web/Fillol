from pathlib import Path

BASE_DIR = Path(__file__).parent

# Pasta onde ficam os exports brutos do SAP (CSV). Um ou vários arquivos,
# inclusive um por mês -- o ETL le todos juntos.
RAW_DIR = BASE_DIR / "data" / "raw"

# Onde o resultado processado é gravado.
OUTPUT_DIR = BASE_DIR / "data" / "output"
DUCKDB_FILE = OUTPUT_DIR / "movimentacao.duckdb"
EXCEL_OUTPUT = OUTPUT_DIR / "consolidado.xlsx"

# Nome das colunas no export bruto do SAP -> nome interno usado no ETL.
# Ajuste aqui se o cabeçalho do seu export vier diferente.
COLUMN_MAP = {
    "Client Name": "cliente",
    "Root CNPJ": "cnpj_raiz",
    "Transaction Date": "data",
    "Document Type": "tipo_documento",
    "Value Confirmed": "valor",
}

# Regras de categoria (Document Type -> categoria do waterfall) ficam em
# category_rules.csv, e não aqui, porque é o dado que você (que conhece a
# regra de negócio) precisa calibrar e revisar com frequência.
CATEGORY_RULES_FILE = BASE_DIR / "category_rules.csv"
