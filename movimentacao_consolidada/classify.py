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


# ZTO não tem categoria fixa: depende do PO Number, em ordem de prioridade:
#   1) 8 primeiros caracteres do PO Number == CNPJ raiz do cliente -> Ressarcimento SAP
#   2) contém "off invoice"/"off_invoice" (qualquer caixa)         -> Off Invoice
#   3) nenhuma das anteriores                                      -> Recálculo
#   4) em qualquer um dos casos acima, se contém "trans dívida ytd"
#      (ignorando o mês/ano do final, ex.: "Trans Dívida YTD Abr26")
#      o valor do lançamento é considerado 0 (não conta em nenhuma soma).
ZTO_OFF_INVOICE_MARCADORES = ("off invoice", "off_invoice")
ZTO_ZERAR_MARCADORES = ("trans dívida ytd", "trans divida ytd")


def _dividir_zto_por_po_number(out: pd.DataFrame) -> pd.DataFrame:
    mask_zto = out["tipo_documento"] == "ZTO"
    if not mask_zto.any():
        return out

    idx = out.loc[mask_zto].index
    po = out.loc[idx, "po_number"].astype(str).str.strip()
    po_lower = po.str.lower()
    cnpj = out.loc[idx, "cnpj_raiz"].astype(str)

    eh_cnpj = po.str.slice(0, 8) == cnpj

    eh_off = pd.Series(False, index=idx)
    for marcador in ZTO_OFF_INVOICE_MARCADORES:
        eh_off |= po_lower.str.contains(marcador, regex=False)

    eh_zerar = pd.Series(False, index=idx)
    for marcador in ZTO_ZERAR_MARCADORES:
        eh_zerar |= po_lower.str.contains(marcador, regex=False)

    categoria_zto = pd.Series("Recálculo", index=idx)
    categoria_zto[eh_off] = "Off Invoice"
    categoria_zto[eh_cnpj] = "Ressarcimento SAP"  # tem prioridade sobre Off Invoice

    out.loc[idx, "categoria"] = categoria_zto
    out.loc[idx, "coluna_valor"] = "valor_confirmado"

    idx_zerar = idx[eh_zerar]
    out.loc[idx_zerar, "valor_confirmado"] = 0.0
    out.loc[idx_zerar, "valor_reservado"] = 0.0
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
