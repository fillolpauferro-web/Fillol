"""Grava o resultado do ETL num .xlsx com abas prontas para usar no Excel.

- "Painel Cliente": escolhe o cliente num dropdown e vê o waterfall mensal
  dele (fórmulas SUMIFS em cima da tabela pequena -- instantâneo).
- "Geral": waterfall consolidado (todos os clientes), já calculado em Python.
- "Por Cliente": tabela fato (uma linha por cliente x mês) que alimenta o
  Painel e também serve de fonte para Tabela Dinâmica / Power Query.
- "Fato (base Python)": granularidade cliente x mês x categoria, caso precise
  auditar ou construir outra visão direto no Excel.
"""
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

from config import EXCEL_OUTPUT

NUM_FMT = "#,##0"
VALUE_COLS_HINT = ("Saldo Inicial", "Saldo Final", "Valor")


def _fmt_mes(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series).dt.strftime("%b/%y")


def _write_table(ws, df: pd.DataFrame, table_name: str, categorias: list) -> None:
    ws.append(list(df.columns))
    for row in df.itertuples(index=False):
        ws.append(list(row))

    n_rows, n_cols = df.shape
    last_col = get_column_letter(n_cols)
    table = Table(displayName=table_name, ref=f"A1:{last_col}{n_rows + 1}")
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2", showRowStripes=True, showFirstColumn=False
    )
    ws.add_table(table)

    valor_cols = set(VALUE_COLS_HINT) | set(categorias)
    for col_idx, col_name in enumerate(df.columns, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = max(14, len(str(col_name)) + 2)
        if col_name in valor_cols:
            for r in range(2, n_rows + 2):
                ws.cell(row=r, column=col_idx).number_format = NUM_FMT
    ws.freeze_panes = "A2"


def _sref(table_name: str, col_name: str) -> str:
    # Nome de coluna com caracteres especiais (ex.: "Canc./Devol") precisa
    # ficar entre colchetes na referência estruturada do Excel.
    return f"{table_name}[{col_name}]"


def write_workbook(
    fato: pd.DataFrame,
    geral: pd.DataFrame,
    por_cliente: pd.DataFrame,
    categorias: list,
) -> None:
    geral = geral.copy()
    por_cliente = por_cliente.copy()
    fato = fato.copy()

    geral["mes"] = _fmt_mes(geral["mes"])
    por_cliente["mes"] = _fmt_mes(por_cliente["mes"])
    fato["mes"] = _fmt_mes(fato["mes"])

    geral = geral.rename(columns={"mes": "Data"})
    por_cliente = por_cliente.rename(
        columns={"cnpj_raiz": "CNPJ", "cliente": "Cliente", "mes": "Data"}
    )
    fato = fato.rename(
        columns={
            "cnpj_raiz": "CNPJ",
            "cliente": "Cliente",
            "mes": "Data",
            "categoria": "Categoria",
            "valor": "Valor",
        }
    )

    wb = Workbook()
    wb.remove(wb.active)

    ws_geral = wb.create_sheet("Geral")
    _write_table(ws_geral, geral, "TabelaGeral", categorias)

    ws_cli = wb.create_sheet("Por Cliente")
    _write_table(ws_cli, por_cliente, "TabelaPorCliente", categorias)

    ws_fato = wb.create_sheet("Fato (base Python)")
    _write_table(ws_fato, fato, "TabelaFato", categorias)

    # --- Painel Cliente: dropdown + SUMIFS sobre TabelaPorCliente ---
    ws_painel = wb.create_sheet("Painel Cliente", 0)
    ws_painel["A1"] = "Cliente:"
    clientes_ordenados = sorted(por_cliente["Cliente"].dropna().unique().tolist())
    ws_painel["B1"] = clientes_ordenados[0] if clientes_ordenados else ""
    ws_painel["A1"].font = ws_painel["B1"].font.copy(bold=True)

    n_linhas_cli = len(por_cliente) + 1  # +1 pela linha de cabeçalho da tabela
    dv = DataValidation(
        type="list",
        formula1=f"='Por Cliente'!$B$2:$B${n_linhas_cli}",
        allow_blank=True,
    )
    ws_painel.add_data_validation(dv)
    dv.add(ws_painel["B1"])

    header_row = 3
    headers = ["Data", "Saldo Inicial", *categorias, "Saldo Final"]
    for col_idx, h in enumerate(headers, start=1):
        ws_painel.cell(row=header_row, column=col_idx, value=h)

    meses = por_cliente[["Data"]].drop_duplicates()
    meses["_ord"] = pd.to_datetime(meses["Data"], format="%b/%y")
    meses = meses.sort_values("_ord")["Data"].tolist()

    valor_cols_painel = ["Saldo Inicial", *categorias, "Saldo Final"]
    for i, mes in enumerate(meses):
        r = header_row + 1 + i
        ws_painel.cell(row=r, column=1, value=mes)
        for j, col_name in enumerate(valor_cols_painel, start=2):
            formula = (
                f"=SUMIFS({_sref('TabelaPorCliente', col_name)},"
                f"{_sref('TabelaPorCliente', 'Cliente')},$B$1,"
                f"{_sref('TabelaPorCliente', 'Data')},$A{r})"
            )
            cell = ws_painel.cell(row=r, column=j, value=formula)
            cell.number_format = NUM_FMT

    for col_idx in range(1, len(headers) + 1):
        ws_painel.column_dimensions[get_column_letter(col_idx)].width = 16
    ws_painel.freeze_panes = "A4"

    wb.save(EXCEL_OUTPUT)
