"""Gera um CSV/XLSX de exemplo (pequeno) só pra testar o pipeline ponta a
ponta. Não é dado real -- é só pra você rodar `python etl.py` uma vez e ver
o resultado antes de apontar --raw-dir para o export de verdade do SAP.

Os Document Type usados aqui (ZTO, ZOR, ZF2, ZG2, ZREA) e o comportamento de
ZOR/ZREA sem valor em Value Confirmed espelham o que apareceu no export real,
pra exercitar o coluna_valor=valor_reservado do category_rules.csv.
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
# ZOR e ZREA só têm valor em "Value Reserved" (Value Confirmed fica 0), igual
# ao export real do usuário.
tipos_confirmado = ["ZTO", "ZF2", "ZG2"]
tipos_reservado = ["ZOR", "ZREA"]

linhas = []
for cnpj, nome in clientes:
    for mes in meses:
        for _ in range(random.randint(4, 10)):
            tipo = random.choice(tipos_confirmado + tipos_reservado)
            valor = round(random.uniform(-50_000, 80_000), 2)
            linhas.append(
                {
                    "Sales Organization": "X401",
                    "Root CNPJ": cnpj,
                    "Client Name": nome,
                    "Transaction Date": mes,
                    "Document Type": tipo,
                    "Value Confirmed": 0.0 if tipo in tipos_reservado else valor,
                    "Value Reserved": valor if tipo in tipos_reservado else 0.0,
                }
            )

df = pd.DataFrame(linhas)
out_csv = Path(__file__).parent / "exemplo_movimentacao.csv"
out_xlsx = Path(__file__).parent / "exemplo_movimentacao.xlsx"
df.to_csv(out_csv, index=False)
df.to_excel(out_xlsx, index=False)
print(f"Gerado {out_csv} e {out_xlsx} com {len(df)} linhas cada.")
