"""Gera um CSV de exemplo (pequeno) só pra testar o pipeline ponta a ponta.
Não é dado real -- é só pra você rodar `python etl.py` uma vez e ver o
resultado antes de apontar RAW_DIR para o export de verdade do SAP.
"""
import random
from datetime import date
from pathlib import Path

import pandas as pd

random.seed(7)

clientes = [
    ("07224991", "NAZARIA DIST PROD FCEUTICO LTDA"),
    ("45453214", "PROFARMA DIST PROD FCEUTICO SA"),
    ("61940292", "DIST MEDIC SANTA CRUZ LTDA"),
    ("46054219", "SOLFARMA COM PROD FCEUTICO S A"),
    ("39455032", "MEDNORTE DIST MEDIC LTDA"),
]
meses = [date(2026, 4, 1), date(2026, 5, 1), date(2026, 6, 1), date(2026, 7, 1)]
tipos = ["ZTO", "ZRE", "ZCA", "ZRC", "ZOF", "ZFT"]

linhas = []
for cnpj, nome in clientes:
    for mes in meses:
        for _ in range(random.randint(3, 8)):
            tipo = random.choice(tipos)
            valor = round(random.uniform(-50_000, 80_000), 2)
            linhas.append(
                {
                    "Sales Organization": "X401",
                    "Root CNPJ": cnpj,
                    "Client Name": nome,
                    "Transaction Date": mes,
                    "Document Type": tipo,
                    "Value Confirmed": valor,
                }
            )

df = pd.DataFrame(linhas)
out = Path(__file__).parent / "exemplo_movimentacao.csv"
df.to_csv(out, index=False)
print(f"Gerado {out} com {len(df)} linhas.")
