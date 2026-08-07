"""Grava o resultado do ETL num .xlsx com abas prontas para usar no Excel.

Cada cliente tem DOIS saldos rastreados em paralelo -- Ressarcimento SAP e
Reserva (config.SALDO_GROUPS) -- cada um com seu próprio Saldo Inicial/Final,
mas compartilhando as mesmas 4 colunas de ajuste (Faturado em Nota, Canc. /
Devolução, Recálculo, Off Invoice).

- "Painel Cliente": escolhe o cliente num dropdown e vê os dois waterfalls
  lado a lado (colunas A-H = Ressarcimento SAP, K-R = Reserva), via SUMIFS
  em cima da tabela pequena -- instantâneo.
- "Geral": os dois waterfalls consolidados (todos os clientes), lado a lado.
- "Por Cliente": tabela fato (uma linha por cliente x mês, colunas dos dois
  grupos lado a lado) que alimenta o Painel e também serve de fonte pra
  Tabela Dinâmica / Power Query.
- "Fato (base Python)": granularidade cliente x mês x categoria (todas as
  categorias que existirem em category_rules.csv), caso precise auditar.
"""
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

import config

NUM_FMT = "#,##0"
GAP_COLS = 2  # colunas em branco entre um bloco de saldo e o próximo (I, J)


def _fmt_mes(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series).dt.strftime("%b/%y")


def _nome_grupo(grupo: dict) -> str:
    return grupo["categoria_propria"]


def _colunas_do_grupo(grupo: dict) -> list:
    return ["Saldo Inicial", grupo["categoria_propria"], *config.CATEGORIAS_COMPARTILHADAS, "Saldo Final"]


def _merge_grupos(dfs_por_grupo: dict, grupos: list, chave_cols: list) -> pd.DataFrame:
    """Junta as tabelas dos dois grupos lado a lado (mesma linha por
    chave_cols), sufixando as colunas de valor com "(nome do grupo)" pra não
    colidir (ex.: "Saldo Final (Ressarcimento SAP)" x "Saldo Final (Reserva)")."""
    base = None
    for grupo in grupos:
        df = dfs_por_grupo[grupo["chave"]].copy()
        nome = _nome_grupo(grupo)
        renomeio = {c: f"{c} ({nome})" for c in df.columns if c not in chave_cols}
        df = df.rename(columns=renomeio)
        base = df if base is None else base.merge(df, on=chave_cols, how="outer")
    return base.sort_values(chave_cols).reset_index(drop=True)


def _write_table(ws, df: pd.DataFrame, table_name: str, valor_cols: set) -> None:
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

    for col_idx, col_name in enumerate(df.columns, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = max(14, len(str(col_name)) + 2)
        if col_name in valor_cols:
            for r in range(2, n_rows + 2):
                ws.cell(row=r, column=col_idx).number_format = NUM_FMT
    ws.freeze_panes = "A2"


def _sref(table_name: str, col_name: str) -> str:
    # Nome de coluna com caracteres especiais (ex.: "Canc. / Devolução (Reserva)")
    # precisa ficar entre colchetes na referência estruturada do Excel.
    return f"{table_name}[{col_name}]"


def _write_painel_cliente(ws, por_cliente: pd.DataFrame, grupos: list) -> None:
    ws["A1"] = "Cliente:"
    clientes_ordenados = sorted(por_cliente["Cliente"].dropna().unique().tolist())
    ws["B1"] = clientes_ordenados[0] if clientes_ordenados else ""
    ws["A1"].font = ws["B1"].font.copy(bold=True)

    n_linhas_cli = len(por_cliente) + 1
    dv = DataValidation(
        type="list", formula1=f"='Por Cliente'!$B$2:$B${n_linhas_cli}", allow_blank=True
    )
    ws.add_data_validation(dv)
    dv.add(ws["B1"])

    meses = por_cliente[["Data"]].drop_duplicates()
    meses["_ord"] = pd.to_datetime(meses["Data"], format="%b/%y")
    meses = meses.sort_values("_ord")["Data"].tolist()

    header_row = 3
    col = 1
    for grupo in grupos:
        nome = _nome_grupo(grupo)
        headers = ["Data", *_colunas_do_grupo(grupo)]
        for j, h in enumerate(headers):
            ws.cell(row=header_row, column=col + j, value=h)

        for k, mes in enumerate(meses):
            r = header_row + 1 + k
            ws.cell(row=r, column=col, value=mes)
            for j, col_name in enumerate(headers[1:], start=1):
                origem = f"{col_name} ({nome})"
                formula = (
                    f"=SUMIFS({_sref('TabelaPorCliente', origem)},"
                    f"{_sref('TabelaPorCliente', 'Cliente')},$B$1,"
                    f"{_sref('TabelaPorCliente', 'Data')},{get_column_letter(col)}{r})"
                )
                cell = ws.cell(row=r, column=col + j, value=formula)
                cell.number_format = NUM_FMT

        for c in range(col, col + len(headers)):
            ws.column_dimensions[get_column_letter(c)].width = 16

        col += len(headers) + GAP_COLS

    ws.freeze_panes = "A4"


def write_workbook(
    fato: pd.DataFrame,
    geral_por_grupo: dict,
    por_cliente_por_grupo: dict,
    grupos: list,
    categorias_fato: list,
    excel_path: Path = None,
) -> None:
    excel_path = Path(excel_path) if excel_path else config.EXCEL_OUTPUT
    fato = fato.copy()

    geral = _merge_grupos(geral_por_grupo, grupos, ["mes"])
    por_cliente = _merge_grupos(por_cliente_por_grupo, grupos, ["cnpj_raiz", "cliente", "mes"])

    geral["mes"] = _fmt_mes(geral["mes"])
    por_cliente["mes"] = _fmt_mes(por_cliente["mes"])
    fato["mes"] = _fmt_mes(fato["mes"])

    geral = geral.rename(columns={"mes": "Data"})
    por_cliente = por_cliente.rename(columns={"cnpj_raiz": "CNPJ", "cliente": "Cliente", "mes": "Data"})
    fato = fato.rename(columns={
        "cnpj_raiz": "CNPJ", "cliente": "Cliente", "mes": "Data",
        "categoria": "Categoria", "valor": "Valor",
    })

    valor_cols_grupos = {f"{c} ({_nome_grupo(g)})" for g in grupos for c in _colunas_do_grupo(g)}
    valor_cols_fato = {"Valor", *categorias_fato}

    wb = Workbook()
    wb.remove(wb.active)

    ws_geral = wb.create_sheet("Geral")
    _write_table(ws_geral, geral, "TabelaGeral", valor_cols_grupos)

    ws_cli = wb.create_sheet("Por Cliente")
    _write_table(ws_cli, por_cliente, "TabelaPorCliente", valor_cols_grupos)

    ws_fato = wb.create_sheet("Fato (base Python)")
    _write_table(ws_fato, fato, "TabelaFato", valor_cols_fato)

    ws_painel = wb.create_sheet("Painel Cliente", 0)
    _write_painel_cliente(ws_painel, por_cliente, grupos)

    wb.save(excel_path)
