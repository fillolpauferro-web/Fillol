"""Aplica a regra de negócio (category_rules.csv) sobre as linhas brutas."""
import pandas as pd

from config import CATEGORY_RULES_FILE


def load_rules() -> pd.DataFrame:
    return pd.read_csv(CATEGORY_RULES_FILE, comment="#")


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
    return out
