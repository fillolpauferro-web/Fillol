"""Aba "Off Invoice a Pagar": quanto pagar de Off Invoice por cliente.

Compara o Saldo Final (painel Ressarcimento SAP/Faturado) do último mês
FECHADO contra um limite de saldo baseado na geração dos últimos 3 meses
fechados, prorateada por 75/60/45 dias conforme o cliente (config.py):

    Limite Saldo = (Ressarcimento SAP de M-3 + M-2 + M-1)
                   / dias corridos nos 3 meses
                   * dias do cliente (75 padrão, 60 ou 45 nas exceções)

    Off Invoice a Pagar = Saldo Final de M-1, SE Saldo Final de M-1 > Limite
                           Saldo, senão 0.

M-1/M-2/M-3 são sempre os últimos 3 meses FECHADOS em relação à data de
referência (hoje, por padrão) -- não o mês corrente. Em agosto, M-1 é
julho; em setembro, vira agosto, e assim por diante.
"""
import pandas as pd

import config


def meses_fechados(data_referencia: pd.Timestamp) -> tuple:
    mes_atual = pd.Timestamp(data_referencia.year, data_referencia.month, 1)
    m1 = mes_atual - pd.DateOffset(months=1)
    m2 = mes_atual - pd.DateOffset(months=2)
    m3 = mes_atual - pd.DateOffset(months=3)
    return m3, m2, m1


def _dias_por_cliente(nome_cliente: str) -> int:
    nome = (nome_cliente or "").upper()
    for chave, dias in config.OFF_INVOICE_DIAS_EXCECOES.items():
        if chave.upper() in nome:
            return dias
    return config.OFF_INVOICE_DIAS_PADRAO


def build_off_invoice(por_cliente_faturado: pd.DataFrame, data_referencia: pd.Timestamp) -> pd.DataFrame:
    m3, m2, m1 = meses_fechados(data_referencia)
    dias_totais = ((m1 + pd.DateOffset(months=1)) - m3).days

    base = por_cliente_faturado[["cnpj_raiz", "cliente", "mes", "Ressarcimento SAP", "Saldo Final"]]
    clientes = base[["cnpj_raiz", "cliente"]].drop_duplicates().sort_values("cliente").reset_index(drop=True)

    def _valor(mes: pd.Timestamp, coluna: str) -> pd.Series:
        sub = base.loc[base["mes"] == mes].set_index("cnpj_raiz")[coluna]
        return clientes["cnpj_raiz"].map(sub).fillna(0.0)

    col_m3 = f"Ressarcimento SAP {m3.strftime('%b/%y')}"
    col_m2 = f"Ressarcimento SAP {m2.strftime('%b/%y')}"
    col_m1 = f"Ressarcimento SAP {m1.strftime('%b/%y')}"
    col_saldo_m1 = f"Saldo Final {m1.strftime('%b/%y')}"

    resultado = clientes.rename(columns={"cnpj_raiz": "CNPJ", "cliente": "Cliente"})
    resultado[col_m3] = _valor(m3, "Ressarcimento SAP")
    resultado[col_m2] = _valor(m2, "Ressarcimento SAP")
    resultado[col_m1] = _valor(m1, "Ressarcimento SAP")

    soma_3_meses = resultado[col_m3] + resultado[col_m2] + resultado[col_m1]
    resultado["Dias (limite)"] = resultado["Cliente"].map(_dias_por_cliente)
    resultado["Limite Saldo"] = soma_3_meses / dias_totais * resultado["Dias (limite)"]

    resultado[col_saldo_m1] = _valor(m1, "Saldo Final")
    resultado["Off Invoice a Pagar"] = resultado[col_saldo_m1].where(
        resultado[col_saldo_m1] > resultado["Limite Saldo"], 0.0
    )

    return resultado
