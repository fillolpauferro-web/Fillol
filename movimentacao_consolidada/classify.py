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
    rules = pd.read_csv(CATEGORY_RULES_FILE, comment="#")
    if "coluna_valor" not in rules.columns:
        rules["coluna_valor"] = pd.NA
    rules["coluna_valor"] = rules["coluna_valor"].fillna(DEFAULT_VALUE_COLUMN)
    return rules


def classify(df: pd.DataFrame) -> pd.DataFrame:
    rules = load_rules()
    out = df.merge(rules, on="tipo_documento", how="left")

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
