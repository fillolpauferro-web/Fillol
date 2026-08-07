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
    "NFE Number": "nfe_number",
    "NFE Item": "nfe_item",
    "PO Number": "po_number",
}

# Colunas que precisam ser lidas como texto (não número) pra não perder zero
# à esquerda -- CNPJ raiz sempre tem 8 dígitos, e vira número errado se o
# pandas/Excel inferir tipo numérico. PO Number também precisa, porque a
# regra do ZTO compara PO Number com CNPJ raiz caractere a caractere.
DTYPE_OVERRIDES = {"Root CNPJ": str, "PO Number": str}

# Regras de categoria (Document Type -> categoria do waterfall) ficam em
# category_rules.csv, e não aqui, porque é o dado que você (que conhece a
# regra de negócio) precisa calibrar e revisar com frequência.
CATEGORY_RULES_FILE = BASE_DIR / "category_rules.csv"

# Mesmo CNPJ pode aparecer no export com grafias diferentes de Client Name
# (ex.: "PANPHARMA DIST MEDIC LTDA" com espaçamento/acento diferente). O ETL
# já consolida automaticamente pelo nome mais frequente; esse arquivo força
# um nome específico quando o automático não for o certo.
CLIENT_NAME_OVERRIDES_FILE = BASE_DIR / "client_name_overrides.csv"

# Categoria especial: Document Type marcado assim em category_rules.csv
# tem regra válida (não dá erro de "sem categoria"), mas é ignorado na hora
# de montar os painéis/waterfall -- usado pra Document Type que só interessa
# na aba "Notas Fiscais" (ex.: ZS1/ZS2).
CATEGORIA_FORA_DO_WATERFALL = "(fora do waterfall)"

# Cada painel é uma aba com seu próprio waterfall independente (Saldo
# Inicial/Final próprios), somando a lista de categorias definida em
# "categorias". "ancora_saldo_externo": True == recebe o valor do arquivo
# de Saldo Inicial (Reimbursement Load) quando ele existir.
PAINEIS = [
    {
        "chave": "faturado",
        "aba": "Painel Cliente Faturado",
        "categorias": ["Ressarcimento SAP", "Faturado em Nota", "Canc. / Devolução", "Recálculo", "Off Invoice"],
        "ancora_saldo_externo": True,
    },
    {
        "chave": "reservado",
        "aba": "Painel Cliente Reservado",
        "categorias": ["Ressarcimento SAP", "Reserva", "Canc. / Devolução", "Recálculo", "Off Invoice"],
        "ancora_saldo_externo": True,
    },
]

# As abas "Geral" e "Por Cliente" mostram a visão do painel "faturado".
PAINEL_PRINCIPAL = "faturado"

# Aba "Notas Fiscais": detalhe (não agregado) por nota, só desses Document
# Type -- independe de category_rules.csv, é um recorte direto do bruto.
NOTAS_FISCAIS_TIPOS_DOCUMENTO = ["ZS1", "ZS2", "ZF1", "ZF2"]
