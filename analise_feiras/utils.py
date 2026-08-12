"""Funções utilitárias compartilhadas pelo pipeline de análise de feiras."""

from __future__ import annotations

import re
import unicodedata
from pathlib import Path

import pandas as pd


def read_table(caminho: str | Path, aba: str | None = None) -> pd.DataFrame:
    """Lê .xlsx/.xls/.csv de forma transparente.

    `aba=None` usa a primeira aba (comportamento padrão do pandas para Excel).
    """
    caminho = Path(caminho)
    if not caminho.exists():
        raise FileNotFoundError(
            f"Arquivo não encontrado: {caminho}. Verifique o caminho no config.yaml."
        )

    if caminho.suffix.lower() in (".csv", ".tsv"):
        sep = "\t" if caminho.suffix.lower() == ".tsv" else ","
        return pd.read_csv(caminho, sep=sep, dtype=str, keep_default_na=False, na_values=[""])

    return pd.read_excel(caminho, sheet_name=aba or 0, dtype=str)


def normalize_text(valor) -> str:
    """Maiúsculas, sem acento, sem espaço nas pontas — para comparações robustas."""
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ""
    texto = str(valor).strip().upper()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = re.sub(r"\s+", " ", texto)
    return texto


def normalize_cnpj(valor) -> str:
    """Mantém só os dígitos do CNPJ, para casar bases com/sem máscara."""
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ""
    return re.sub(r"\D", "", str(valor))


def to_datetime(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce", dayfirst=True)


def to_numeric(series: pd.Series) -> pd.Series:
    """Converte texto numérico pt-BR (1.234,56 ou 1234,56) para float."""
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")

    limpo = (
        series.astype(str)
        .str.strip()
        .str.replace(r"^'", "", regex=True)  # aspas de "número como texto" do Excel
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
    )
    return pd.to_numeric(limpo, errors="coerce")
