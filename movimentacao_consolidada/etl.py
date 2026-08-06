"""ETL da Movimentação: 800k+ linhas brutas -> tabela fato compacta (Cliente x
Mês x Categoria) -> waterfall Geral e Por Cliente -> Parquet + Excel.

Uso:
    python etl.py            # roda o pipeline completo
    python etl.py --audit    # só lista Document Type x soma, pra calibrar
                              # category_rules.csv (não escreve nada)
"""
import argparse
from pathlib import Path

import pandas as pd

import config
from config import COLUMN_MAP, DTYPE_OVERRIDES
from classify import classify, load_rules

EXTENSOES_SUPORTADAS = (".csv", ".xlsx", ".xlsm")


def _read_one(path: Path) -> pd.DataFrame:
    suf = path.suffix.lower()
    if suf == ".csv":
        return pd.read_csv(path, dtype=DTYPE_OVERRIDES)
    if suf in (".xlsx", ".xlsm"):
        # 'calamine' (python-calamine) lê xlsx grande muito mais rápido que o
        # engine padrão do pandas/openpyxl -- importante pros 800k+ linhas do
        # export do SAP. Se não estiver instalado, cai pro openpyxl mesmo.
        try:
            return pd.read_excel(path, engine="calamine", dtype=DTYPE_OVERRIDES)
        except (ImportError, ValueError):
            return pd.read_excel(path, engine="openpyxl", dtype=DTYPE_OVERRIDES)
    raise ValueError(f"Formato não suportado: {path.name}")


def listar_arquivos(raw_dir: Path) -> list:
    return sorted(f for ext in EXTENSOES_SUPORTADAS for f in raw_dir.glob(f"*{ext}"))


def read_raw(files: list) -> pd.DataFrame:
    if not files:
        raise FileNotFoundError(
            f"Nenhum arquivo de Movimentação {EXTENSOES_SUPORTADAS} encontrado. "
            "Coloque o export do Movimentação (xlsx ou csv) na pasta -- pode ser "
            "um arquivo por mês."
        )
    partes = [_read_one(f) for f in files]
    df = pd.concat(partes, ignore_index=True, sort=False)

    faltando = [c for c in COLUMN_MAP if c not in df.columns]
    if faltando:
        raise KeyError(
            f"Coluna(s) esperada(s) não encontrada(s) no export: {faltando}. "
            "Confira o cabeçalho do arquivo ou ajuste COLUMN_MAP em config.py."
        )
    df = df.rename(columns=COLUMN_MAP)
    # Reforço além do dtype na leitura: CNPJ raiz é sempre 8 dígitos. Se em
    # algum arquivo ele ainda vier como número (perdendo zero à esquerda),
    # zfill restaura -- se já vier certo, zfill não muda nada.
    df["cnpj_raiz"] = df["cnpj_raiz"].astype(str).str.strip().str.zfill(8)
    df["data"] = pd.to_datetime(df["data"])
    df["mes"] = df["data"].values.astype("datetime64[M]")
    return df


def audit(df: pd.DataFrame) -> None:
    from classify import linhas_sem_tipo

    vazios = linhas_sem_tipo(df)
    if vazios.any():
        print(
            f"{vazios.sum():,} linha(s) com Document Type vazio -- serão "
            "ignoradas ao rodar sem --audit (provavelmente linha em "
            "branco/rodapé do export).\n"
        )
        df = df.loc[~vazios]

    rules = load_rules().set_index("tipo_documento")
    resumo = df.groupby("tipo_documento").agg(
        linhas=("valor_confirmado", "count"),
        soma_confirmado=("valor_confirmado", "sum"),
        soma_reservado=("valor_reservado", "sum"),
    ).sort_values("linhas", ascending=False)

    resumo["categoria_atual"] = [
        rules["categoria"].get(t, "<<< SEM REGRA -- adicionar no category_rules.csv")
        for t in resumo.index
    ]
    resumo["coluna_valor_atual"] = [
        rules["coluna_valor"].get(t) if t in rules.index and pd.notna(rules["coluna_valor"].get(t)) else "valor_confirmado (padrão)"
        for t in resumo.index
    ]
    with pd.option_context("display.float_format", "{:,.2f}".format):
        print(resumo.to_string())
    print(
        "\nDica: se soma_confirmado vier 0,00 mas soma_reservado não, esse "
        "Document Type provavelmente deve usar coluna_valor=valor_reservado "
        "no category_rules.csv (foi o caso do ZOR)."
    )


def build_fact_table(df: pd.DataFrame) -> pd.DataFrame:
    df = classify(df)
    fato = (
        df.groupby(["cnpj_raiz", "cliente", "mes", "categoria"])["valor"]
        .sum()
        .reset_index()
    )
    return fato


def _garantir_linhas_ancora(
    tabela: pd.DataFrame, grupo: list, por_cliente: bool, categorias: list,
    saldo_externo: dict, saldo_externo_mes, nomes_extra: dict,
) -> pd.DataFrame:
    """Garante que todo cliente do saldo_externo tenha uma linha no mês âncora,
    mesmo sem movimento nesse mês -- senão o saldo dele "some" até o próximo
    mês em que ele voltar a ter movimentação."""
    if not saldo_externo or saldo_externo_mes is None:
        return tabela

    if por_cliente:
        existentes = set(tabela.loc[tabela["mes"] == saldo_externo_mes, "cnpj_raiz"])
        faltando = [c for c in saldo_externo if c not in existentes]
        if not faltando:
            return tabela
        novas = pd.DataFrame({
            "cnpj_raiz": faltando,
            "cliente": [nomes_extra.get(c, "") for c in faltando],
            "mes": saldo_externo_mes,
            **{c: 0.0 for c in categorias},
        })
    else:
        if (tabela["mes"] == saldo_externo_mes).any():
            return tabela
        novas = pd.DataFrame({"mes": [saldo_externo_mes], **{c: [0.0] for c in categorias}})

    return pd.concat([tabela, novas], ignore_index=True).sort_values(grupo).reset_index(drop=True)


def build_waterfall(
    fato: pd.DataFrame, por_cliente: bool, categorias: list,
    saldo_externo: dict = None, saldo_externo_mes=None, nomes_extra: dict = None,
) -> pd.DataFrame:
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

    tabela = _garantir_linhas_ancora(
        tabela, grupo, por_cliente, categorias, saldo_externo, saldo_externo_mes, nomes_extra or {}
    )

    chave_cols = ["cnpj_raiz"] if por_cliente else []
    saldo_inicial, saldo_final = [], []
    saldo_anterior: dict = {}
    for _, row in tabela.iterrows():
        chave = tuple(row[c] for c in chave_cols)
        eh_mes_ancora = saldo_externo and saldo_externo_mes is not None and row["mes"] == saldo_externo_mes
        if eh_mes_ancora:
            if por_cliente:
                inicial = saldo_externo.get(row["cnpj_raiz"], saldo_anterior.get(chave, 0.0))
            else:
                inicial = float(sum(saldo_externo.values()))
        else:
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
    parser.add_argument(
        "--raw-dir", type=str, default=None,
        help=r'Pasta com o(s) export(s) do SAP. Ex: --raw-dir "C:\Users\...\OL Robo". '
             "Se omitido, usa RAW_DIR de config.py (data/raw/).",
    )
    parser.add_argument(
        "--output-dir", type=str, default=None,
        help="Pasta de saída (parquet + consolidado.xlsx). Se omitido, usa OUTPUT_DIR de config.py.",
    )
    parser.add_argument(
        "--saldo-inicial-data", type=str, default=None,
        help='Mês (YYYY-MM-DD) que o arquivo de Saldo Inicial representa -- o '
             '"dia zero" do waterfall (ex.: 2026-05-01 = saldo em 01/05/2026). '
             "Obrigatório se houver algum arquivo com 'Saldo Inicial' no nome "
             "dentro da pasta.",
    )
    args = parser.parse_args()

    raw_dir = Path(args.raw_dir) if args.raw_dir else config.RAW_DIR
    output_dir = Path(args.output_dir) if args.output_dir else config.OUTPUT_DIR
    output_dir.mkdir(parents=True, exist_ok=True)

    from saldo_inicial import eh_arquivo_saldo_inicial, read_saldo_inicial

    todos = listar_arquivos(raw_dir)
    arquivos_saldo = [f for f in todos if eh_arquivo_saldo_inicial(f.name)]
    arquivos_movimento = [f for f in todos if f not in arquivos_saldo]

    df = read_raw(arquivos_movimento)

    if args.audit:
        audit(df)
        return

    saldo_externo = None
    saldo_externo_mes = None
    nomes_extra = None
    if arquivos_saldo:
        if not args.saldo_inicial_data:
            raise ValueError(
                f"Encontrei arquivo(s) de Saldo Inicial ({[f.name for f in arquivos_saldo]}) "
                "mas falta informar --saldo-inicial-data (ex.: 2026-05-01) pra saber "
                "a que mês eles se referem."
            )
        partes_saldo = [read_saldo_inicial(f, _read_one) for f in arquivos_saldo]
        saldo_df = pd.concat(partes_saldo, ignore_index=True)
        saldo_df = saldo_df.groupby("cnpj_raiz", as_index=False).agg(
            cliente=("cliente", "first"), saldo_inicial_externo=("saldo_inicial_externo", "sum")
        )
        saldo_externo = dict(zip(saldo_df["cnpj_raiz"], saldo_df["saldo_inicial_externo"]))
        nomes_extra = dict(zip(saldo_df["cnpj_raiz"], saldo_df["cliente"]))
        saldo_externo_mes = pd.Timestamp(args.saldo_inicial_data).to_period("M").to_timestamp()
        print(
            f"Saldo Inicial externo: {len(saldo_externo)} cliente(s), "
            f"ancorado em {saldo_externo_mes.strftime('%Y-%m')}."
        )

    fato = build_fact_table(df)
    categorias = sorted(fato["categoria"].unique().tolist())

    geral = build_waterfall(
        fato, por_cliente=False, categorias=categorias,
        saldo_externo=saldo_externo, saldo_externo_mes=saldo_externo_mes,
    )
    por_cliente = build_waterfall(
        fato, por_cliente=True, categorias=categorias,
        saldo_externo=saldo_externo, saldo_externo_mes=saldo_externo_mes, nomes_extra=nomes_extra,
    )

    fato.to_parquet(output_dir / "fato_movimentacao.parquet", index=False)
    geral.to_parquet(output_dir / "consolidado_geral.parquet", index=False)
    por_cliente.to_parquet(output_dir / "consolidado_por_cliente.parquet", index=False)

    excel_path = output_dir / "consolidado.xlsx"
    from build_excel import write_workbook
    write_workbook(fato, geral, por_cliente, categorias, excel_path)

    print(f"OK: {len(df):,} linhas brutas -> {len(por_cliente):,} linhas na tabela fato consolidada.")
    print(f"Excel gerado em: {excel_path.resolve()}")


if __name__ == "__main__":
    main()
