"""Aplica a regra de negócio (category_rules.csv) sobre as linhas brutas."""
import pandas as pd

from config import CATEGORY_RULES_FILE

# Colunas do export que podem ser usadas como "valor" de um Document Type.
# Alguns tipos (ex.: ZOR) não têm o valor em Value Confirmed -- o valor real
# está em Value Reserved. category_rules.csv escolhe qual usar por linha.
VALUE_COLUMNS = {
    "valor_confirmado": "valor_confirmado",
    "valor_reservado": "valor_reservado",
}
DEFAULT_VALUE_COLUMN = "valor_confirmado"


def load_rules() -> pd.DataFrame:
    # utf-8 primeiro; se o CSV foi salvo pelo Bloco de Notas em "ANSI"
    # (comum no Windows/pt-BR ao colar acento), cai pro cp1252.
    try:
        rules = pd.read_csv(CATEGORY_RULES_FILE, comment="#", encoding="utf-8")
    except UnicodeDecodeError:
        rules = pd.read_csv(CATEGORY_RULES_FILE, comment="#", encoding="cp1252")
    if "coluna_valor" not in rules.columns:
        rules["coluna_valor"] = pd.NA
    rules["coluna_valor"] = rules["coluna_valor"].fillna(DEFAULT_VALUE_COLUMN)
    return rules


def linhas_sem_tipo(df: pd.DataFrame) -> pd.Series:
    """Linhas com Document Type vazio (linha em branco/rodapé do export)."""
    return df["tipo_documento"].isna() | (df["tipo_documento"].astype(str).str.strip() == "")


# ZTO não tem categoria fixa: depende do texto do PO Number.
#   PO Number contém "off invoice"/"off_invoice" (qualquer caixa) -> Off Invoice
#   caso contrário                                                -> Recálculo
# Isso substitui totalmente o antigo mapeamento fixo ZTO -> Ressarcimento SAP.
ZTO_OFF_INVOICE_MARCADORES = ("off invoice", "off_invoice")


def _dividir_zto_por_po_number(out: pd.DataFrame) -> pd.DataFrame:
    mask_zto = out["tipo_documento"] == "ZTO"
    if not mask_zto.any():
        return out

    po = out.loc[mask_zto, "po_number"].astype(str).str.lower()
    eh_off = pd.Series(False, index=po.index)
    for marcador in ZTO_OFF_INVOICE_MARCADORES:
        eh_off |= po.str.contains(marcador, regex=False)

    categoria_zto = pd.Series("Recálculo", index=out.loc[mask_zto].index)
    categoria_zto[eh_off] = "Off Invoice"

    out.loc[mask_zto, "categoria"] = categoria_zto
    out.loc[mask_zto, "coluna_valor"] = "valor_confirmado"
    return out


def classify(df: pd.DataFrame) -> pd.DataFrame:
    vazios = linhas_sem_tipo(df)
    if vazios.any():
        print(
            f"Aviso: {vazios.sum():,} linha(s) sem Document Type foram ignoradas "
            "(provavelmente linha em branco/rodapé do export)."
        )
        df = df.loc[~vazios].copy()

    rules = load_rules()
    out = df.merge(rules, on="tipo_documento", how="left")
    out = _dividir_zto_por_po_number(out)

    sem_regra = out["categoria"].isna()
    if sem_regra.any():
        tipos = sorted(out.loc[sem_regra, "tipo_documento"].dropna().unique().tolist())
        raise ValueError(
            "Document Type sem categoria em category_rules.csv: "
            f"{tipos}. Rode `python etl.py --audit` para ver o resumo por tipo "
            "e adicione as linhas que faltam no CSV."
        )

    invalidas = ~out["coluna_valor"].isin(VALUE_COLUMNS)
    if invalidas.any():
        ruins = sorted(out.loc[invalidas, "coluna_valor"].dropna().unique().tolist())
        raise ValueError(
            f"coluna_valor inválida em category_rules.csv: {ruins}. "
            f"Use um destes: {list(VALUE_COLUMNS)}."
        )

    out["valor"] = out["valor_confirmado"]
    usa_reservado = out["coluna_valor"] == "valor_reservado"
    out.loc[usa_reservado, "valor"] = out.loc[usa_reservado, "valor_reservado"]
    return out
