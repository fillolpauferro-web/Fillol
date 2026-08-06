"""ETL da Movimentação: 800k+ linhas brutas -> tabela fato compacta (Cliente x
Mês x Categoria) -> waterfall Geral e Por Cliente -> Parquet + Excel.

Uso:
    python etl.py            # roda o pipeline completo
    python etl.py --audit    # só lista Document Type x soma, pra calibrar
                              # category_rules.csv (não escreve nada)
"""
import argparse

import duckdb
import pandas as pd

from config import RAW_DIR, OUTPUT_DIR, DUCKDB_FILE, COLUMN_MAP
from classify import classify, load_rules


def read_raw(con: duckdb.DuckDBPyConnection) -> pd.DataFrame:
    files = sorted(RAW_DIR.glob("*.csv"))
    if not files:
        raise FileNotFoundError(
            f"Nenhum .csv em {RAW_DIR}. Exporte o Movimentação do SAP como CSV "
            "e coloque o(s) arquivo(s) nessa pasta (pode ser um por mês)."
        )
    pattern = str(RAW_DIR / "*.csv")
    # DuckDB lê todos os CSV da pasta em uma única passada, sem estourar
    # memória -- é isso que substitui abrir 800k linhas dentro do Excel.
    df = con.execute(
        f"SELECT * FROM read_csv_auto('{pattern}', union_by_name=True)"
    ).df()
    faltando = [c for c in COLUMN_MAP if c not in df.columns]
    if faltando:
        raise KeyError(
            f"Coluna(s) esperada(s) não encontrada(s) no export: {faltando}. "
            "Confira o cabeçalho do CSV ou ajuste COLUMN_MAP em config.py."
        )
    df = df.rename(columns=COLUMN_MAP)
    df["data"] = pd.to_datetime(df["data"])
    df["mes"] = df["data"].values.astype("datetime64[M]")
    return df


def audit(df: pd.DataFrame) -> None:
    rules = load_rules().set_index("tipo_documento")["categoria"]
    resumo = (
        df.groupby("tipo_documento")["valor"]
        .agg(linhas="count", soma="sum")
        .sort_values("linhas", ascending=False)
    )
    resumo["categoria_atual"] = [
        rules.get(t, "<<< SEM REGRA -- adicionar no category_rules.csv")
        for t in resumo.index
    ]
    with pd.option_context("display.float_format", "{:,.2f}".format):
        print(resumo.to_string())


def build_fact_table(df: pd.DataFrame) -> pd.DataFrame:
    df = classify(df)
    fato = (
        df.groupby(["cnpj_raiz", "cliente", "mes", "categoria"])["valor"]
        .sum()
        .reset_index()
    )
    return fato


def build_waterfall(fato: pd.DataFrame, por_cliente: bool, categorias: list) -> pd.DataFrame:
    if por_cliente:
        grupo = ["cnpj_raiz", "cliente", "mes"]
        fato_grupo = fato
    else:
        grupo = ["mes"]
        fato_grupo = fato.groupby(["mes", "categoria"], as_index=False)["valor"].sum()

    tabela = fato_grupo.pivot_table(
        index=grupo, columns="categoria", values="valor", aggfunc="sum", fill_value=0
    ).reset_index()
    for c in categorias:
        if c not in tabela.columns:
            tabela[c] = 0.0
    tabela = tabela.sort_values(grupo).reset_index(drop=True)

    chave_cols = ["cnpj_raiz"] if por_cliente else []
    saldo_inicial, saldo_final = [], []
    saldo_anterior: dict = {}
    for _, row in tabela.iterrows():
        chave = tuple(row[c] for c in chave_cols)
        inicial = saldo_anterior.get(chave, 0.0)
        final = inicial + float(sum(row[c] for c in categorias))
        saldo_inicial.append(inicial)
        saldo_final.append(final)
        saldo_anterior[chave] = final

    tabela["Saldo Inicial"] = saldo_inicial
    tabela["Saldo Final"] = saldo_final
    ordem = grupo + ["Saldo Inicial"] + categorias + ["Saldo Final"]
    return tabela[ordem]


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--audit", action="store_true",
        help="Lista Document Type x linhas x soma, sem gravar nada (use pra calibrar category_rules.csv)",
    )
    args = parser.parse_args()

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    con = duckdb.connect(str(DUCKDB_FILE))
    df = read_raw(con)

    if args.audit:
        audit(df)
        return

    fato = build_fact_table(df)
    categorias = sorted(fato["categoria"].unique().tolist())

    geral = build_waterfall(fato, por_cliente=False, categorias=categorias)
    por_cliente = build_waterfall(fato, por_cliente=True, categorias=categorias)

    fato.to_parquet(OUTPUT_DIR / "fato_movimentacao.parquet", index=False)
    geral.to_parquet(OUTPUT_DIR / "consolidado_geral.parquet", index=False)
    por_cliente.to_parquet(OUTPUT_DIR / "consolidado_por_cliente.parquet", index=False)

    from build_excel import write_workbook
    write_workbook(fato, geral, por_cliente, categorias)

    print(f"OK: {len(df):,} linhas brutas -> {len(por_cliente):,} linhas na tabela fato consolidada.")
    print(f"Excel gerado em: {(OUTPUT_DIR / 'consolidado.xlsx').resolve()}")


if __name__ == "__main__":
    main()
