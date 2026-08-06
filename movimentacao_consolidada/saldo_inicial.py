"""Lê um arquivo de Saldo Inicial: mesmo formato do export de Movimentação,
mas em vez de classificar por Document Type, soma a coluna "Reimbursement
Load" por cliente. Usado como "âncora" do Saldo Inicial de um mês específico
do waterfall (ex.: o saldo em 01/05/2026), em vez de deixar o primeiro mês
do export começar do zero.
"""
from pathlib import Path

import pandas as pd

COLUNAS_NECESSARIAS = ("Root CNPJ", "Client Name", "Reimbursement Load")


def eh_arquivo_saldo_inicial(nome_arquivo: str) -> bool:
    n = nome_arquivo.lower()
    return "saldo" in n and "inicial" in n


def read_saldo_inicial(path: Path, ler_arquivo) -> pd.DataFrame:
    """ler_arquivo: função tipo etl._read_one, reaproveitada pra não duplicar
    a lógica de csv/xlsx/xlsm + engine calamine."""
    df = ler_arquivo(Path(path))
    faltando = [c for c in COLUNAS_NECESSARIAS if c not in df.columns]
    if faltando:
        raise KeyError(
            f"Coluna(s) esperada(s) não encontrada(s) no arquivo de Saldo "
            f"Inicial ({Path(path).name}): {faltando}."
        )
    df = df.rename(columns={
        "Root CNPJ": "cnpj_raiz",
        "Client Name": "cliente",
        "Reimbursement Load": "valor_reembolso",
    })
    df["cnpj_raiz"] = df["cnpj_raiz"].astype(str).str.strip().str.zfill(8)

    from nomes import aplicar_overrides
    df = aplicar_overrides(df)

    resumo = (
        df.groupby("cnpj_raiz")
        .agg(cliente=("cliente", "first"), saldo_inicial_externo=("valor_reembolso", "sum"))
        .reset_index()
    )
    return resumo
