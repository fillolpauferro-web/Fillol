"""Gera um CSV/XLSX de exemplo (pequeno) só pra testar o pipeline ponta a
ponta. Não é dado real -- é só pra você rodar `python etl.py` uma vez e ver
o resultado antes de apontar --raw-dir para o export de verdade do SAP.

Os Document Type usados aqui (ZTO, ZOR, ZF1, ZF2, ZG1, ZG2, ZREA, ZREB, ZS1,
ZS2) e o comportamento de ZOR/ZREA/ZREB sem valor em Value Confirmed
espelham o que apareceu no export real, pra exercitar o
coluna_valor=valor_reservado do category_rules.csv e a aba Notas Fiscais.
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
# ZOR/ZREA/ZREB só têm valor em "Value Reserved" (Value Confirmed fica 0),
# igual ao export real do usuário. ZS1/ZS2 não entram em nenhum waterfall
# (só aparecem na aba Notas Fiscais).
tipos_confirmado = ["ZTO", "ZF1", "ZF2", "ZG1", "ZG2"]
tipos_reservado = ["ZOR", "ZREA", "ZREB"]
tipos_notas_extra = ["ZS1", "ZS2"]  # também vão pra Notas Fiscais, sem entrar no waterfall

nfe_seq = 1000
linhas = []
for cnpj, nome in clientes:
    for mes in meses:
        for _ in range(random.randint(4, 10)):
            tipo = random.choice(tipos_confirmado + tipos_reservado + tipos_notas_extra)
            valor = round(random.uniform(-50_000, 80_000), 2)
            nfe_seq += 1
            linhas.append(
                {
                    "Sales Organization": "X401",
                    "Root CNPJ": cnpj,
                    "Client Name": nome,
                    "Transaction Date": mes,
                    "Document Type": tipo,
                    "Value Confirmed": 0.0 if tipo in tipos_reservado else valor,
                    "Value Reserved": valor if tipo in tipos_reservado else 0.0,
                    "NFE Number": f"NF{nfe_seq}",
                    "NFE Item": "000010",
                    "PO Number": f"PO{nfe_seq}",
                }
            )

df = pd.DataFrame(linhas)
out_csv = Path(__file__).parent / "exemplo_movimentacao.csv"
out_xlsx = Path(__file__).parent / "exemplo_movimentacao.xlsx"
df.to_csv(out_csv, index=False)
df.to_excel(out_xlsx, index=False)
print(f"Gerado {out_csv} e {out_xlsx} com {len(df)} linhas cada.")
