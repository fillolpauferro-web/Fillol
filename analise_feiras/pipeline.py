"""Pipeline recorrente de análise de matrizes comerciais (Feira, Campanha, ...).

Passos (para cada matriz ativa no config.yaml):
  1. Filtra a base grande pela palavra-chave da matriz na coluna
     "Tabela de negociação" e exporta o recorte limpo.
  2. Faz um PROCX (merge) com o arquivo de controle da matriz (aba "dados"),
     trazendo feira/campanha, data inicial e vigência.
  3. Compara a vigência com a data do pedido e marca a coluna "Check" como
     OK ou ERRO OPERACIONAL (pedido antes/depois da vigência).
  4. Para os pedidos com erro operacional, cruza com Condicao_comercial para
     achar o desconto correto, calcula o preço sem desconto e o preço
     líquido que deveria ter sido faturado.

Uso:
    python pipeline.py                       # roda as matrizes marcadas ativo: true
    python pipeline.py --matrizes Feira       # roda só "Feira", ignora o config
    python pipeline.py --listar               # lista as matrizes configuradas e sai
    python pipeline.py --config outro.yaml
"""

from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd
import yaml

from utils import normalize_cnpj, normalize_text, read_table, read_table_or_glob, to_datetime, to_numeric

BASE_DIR = Path(__file__).resolve().parent

CHECK_OK = "OK"
CHECK_ANTES = "ERRO OPERACIONAL - pedido antes da vigência"
CHECK_DEPOIS = "ERRO OPERACIONAL - pedido após a vigência"
CHECK_SEM_CADASTRO = "SEM CADASTRO NA MATRIZ"


def carregar_config(caminho: str | Path) -> dict:
    with open(caminho, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def carregar_base(cfg: dict) -> pd.DataFrame:
    caminho = BASE_DIR / cfg["base"]["arquivo"]
    df = read_table_or_glob(caminho, cfg["base"].get("aba"))

    colunas = cfg["base"]["colunas"]
    faltando = [c for c in colunas.values() if c not in df.columns]
    if faltando:
        raise KeyError(
            f"Colunas ausentes na base ({caminho.name}): {faltando}. "
            "Ajuste base.colunas no config.yaml para os nomes reais da planilha."
        )

    df["_cnpj_norm"] = df[colunas["cnpj"]].map(normalize_cnpj)
    df["_ean_norm"] = df[colunas["ean"]].astype(str).str.strip()
    df["_tabela_norm"] = df[colunas["tabela_negociacao"]].map(normalize_text)
    df["_data_pedido"] = to_datetime(df[colunas["data_pedido"]])
    df["_faturado_liquido"] = to_numeric(df[colunas["faturado_liquido"]])
    df["_desconto_aplicado_pct"] = to_numeric(df[colunas["desconto_aplicado_pct"]])
    return df


def filtrar_matriz(df_base: pd.DataFrame, palavra_chave: str) -> pd.DataFrame:
    chave = normalize_text(palavra_chave)
    mask = df_base["_tabela_norm"].str.contains(chave, na=False)
    return df_base.loc[mask].copy()


def carregar_controle(matriz_cfg: dict) -> pd.DataFrame:
    caminho = BASE_DIR / matriz_cfg["arquivo_controle"]
    df = read_table(caminho, matriz_cfg.get("aba_controle"))

    chave_col = matriz_cfg["chave_controle"]
    colunas_trazidas = matriz_cfg["colunas_trazidas"]

    obrigatorias = [chave_col, *colunas_trazidas.values()]
    faltando = [c for c in obrigatorias if c not in df.columns]
    if faltando:
        raise KeyError(
            f"Colunas ausentes no arquivo de controle ({caminho.name}): {faltando}. "
            "Ajuste chave_controle/colunas_trazidas no config.yaml."
        )

    df = df.rename(columns={v: k for k, v in colunas_trazidas.items()})
    df["_chave_controle_norm"] = df[chave_col].map(normalize_text)
    df["data_inicial"] = to_datetime(df["data_inicial"])
    df["vigencia"] = to_datetime(df["vigencia"])

    colunas_saida = ["_chave_controle_norm", "feira", "data_inicial", "vigencia"]
    df_controle = df[colunas_saida].drop_duplicates(subset="_chave_controle_norm", keep="first")
    return df_controle


def aplicar_procx_vigencia(df_matriz: pd.DataFrame, df_controle: pd.DataFrame) -> pd.DataFrame:
    """Traz feira/data_inicial/vigencia da tabela de controle (tipo PROCX)."""
    return df_matriz.merge(
        df_controle,
        how="left",
        left_on="_tabela_norm",
        right_on="_chave_controle_norm",
    )


def calcular_check(df: pd.DataFrame) -> pd.Series:
    sem_cadastro = df["data_inicial"].isna() | df["vigencia"].isna()
    antes = df["_data_pedido"] < df["data_inicial"]
    depois = df["_data_pedido"] > df["vigencia"]

    check = pd.Series(CHECK_OK, index=df.index)
    check = check.mask(depois, CHECK_DEPOIS)
    check = check.mask(antes, CHECK_ANTES)
    check = check.mask(sem_cadastro, CHECK_SEM_CADASTRO)
    return check


def carregar_condicao_comercial(cfg: dict) -> pd.DataFrame:
    cc_cfg = cfg["condicao_comercial"]
    caminho = BASE_DIR / cc_cfg["arquivo"]
    df = read_table(caminho, cc_cfg.get("aba"))

    col_cnpj = cc_cfg["colunas"]["chave_cnpj"]
    col_ean = cc_cfg["colunas"]["chave_ean"]
    col_desconto = cc_cfg["colunas"]["desconto_correto_pct"]

    faltando = [c for c in (col_cnpj, col_ean, col_desconto) if c not in df.columns]
    if faltando:
        raise KeyError(
            f"Colunas ausentes em Condicao_comercial ({caminho.name}): {faltando}. "
            "Ajuste condicao_comercial.colunas no config.yaml."
        )

    df["_cnpj_norm"] = df[col_cnpj].map(normalize_cnpj)
    df["_ean_norm"] = df[col_ean].astype(str).str.strip()
    df["_desconto_correto_pct"] = to_numeric(df[col_desconto])

    return df[["_cnpj_norm", "_ean_norm", "_desconto_correto_pct"]].drop_duplicates(
        subset=["_cnpj_norm", "_ean_norm"], keep="first"
    )


def aplicar_condicao_correta(df: pd.DataFrame, df_condicao: pd.DataFrame) -> pd.DataFrame:
    df = df.merge(df_condicao, how="left", on=["_cnpj_norm", "_ean_norm"])

    # preço sem desconto = líquido faturado / (1 - desconto aplicado)
    fator_aplicado = 1 - (df["_desconto_aplicado_pct"] / 100)
    df["preco_sem_desconto"] = (df["_faturado_liquido"] / fator_aplicado).where(fator_aplicado != 0)

    # preço líquido que deveria ter sido faturado com o desconto correto
    fator_correto = 1 - (df["_desconto_correto_pct"] / 100)
    df["preco_liquido_desconto_correto"] = df["preco_sem_desconto"] * fator_correto

    df["diferenca_faturamento"] = df["_faturado_liquido"] - df["preco_liquido_desconto_correto"]

    # só faz sentido preencher esses cálculos para linhas de erro operacional
    erro_operacional = df["Check"].isin([CHECK_ANTES, CHECK_DEPOIS])
    for col in (
        "_desconto_correto_pct",
        "preco_sem_desconto",
        "preco_liquido_desconto_correto",
        "diferenca_faturamento",
    ):
        df[col] = df[col].where(erro_operacional)

    df = df.rename(columns={"_desconto_correto_pct": "desconto_correto_pct"})
    return df


def montar_saida(df: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    colunas_base_originais = list(cfg["base"]["colunas"].values())
    colunas_novas = [
        "feira",
        "data_inicial",
        "vigencia",
        "Check",
        "desconto_correto_pct",
        "preco_sem_desconto",
        "preco_liquido_desconto_correto",
        "diferenca_faturamento",
    ]
    return df[colunas_base_originais + colunas_novas].copy()


def rodar_matriz(nome_matriz: str, matriz_cfg: dict, df_base: pd.DataFrame, cfg: dict) -> pd.DataFrame | None:
    print(f"\n=== Matriz: {nome_matriz} ===")

    df_filtrado = filtrar_matriz(df_base, matriz_cfg["palavra_chave"])
    if df_filtrado.empty:
        print(f"Nenhuma linha encontrada para a palavra-chave '{matriz_cfg['palavra_chave']}'.")
        return None
    print(f"{len(df_filtrado)} linhas filtradas em 'Tabela de negociação'.")

    pasta_saida = BASE_DIR / cfg["saida"]["pasta"]
    pasta_saida.mkdir(parents=True, exist_ok=True)
    caminho_filtrado = pasta_saida / f"{nome_matriz}_filtrado.xlsx"
    df_filtrado.drop(columns=[c for c in df_filtrado.columns if c.startswith("_")]).to_excel(
        caminho_filtrado, index=False
    )
    print(f"Recorte limpo salvo em {caminho_filtrado}")

    df_controle = carregar_controle(matriz_cfg)
    df_merge = aplicar_procx_vigencia(df_filtrado, df_controle)

    df_merge["Check"] = calcular_check(df_merge)
    print(df_merge["Check"].value_counts(dropna=False).to_string())

    df_condicao = carregar_condicao_comercial(cfg)
    df_final = aplicar_condicao_correta(df_merge, df_condicao)

    df_saida = montar_saida(df_final, cfg)
    caminho_final = pasta_saida / f"{nome_matriz}_analise.xlsx"
    df_saida.to_excel(caminho_final, index=False)
    print(f"Análise completa salva em {caminho_final}")

    impacto = df_saida["diferenca_faturamento"].sum(skipna=True)
    print(f"Impacto financeiro total dos erros operacionais: R$ {impacto:,.2f}")

    return df_saida


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--config", default=str(BASE_DIR / "config.yaml"), help="Caminho do config.yaml")
    parser.add_argument(
        "--matrizes",
        help="Lista separada por vírgula das matrizes a rodar (ignora o 'ativo' do config). Ex: Feira,Campanha",
    )
    parser.add_argument("--listar", action="store_true", help="Lista as matrizes configuradas e sai")
    args = parser.parse_args()

    cfg = carregar_config(args.config)

    if args.listar:
        for m in cfg["matrizes"]:
            print(f"- {m['nome']} (ativo={m['ativo']}, palavra_chave='{m['palavra_chave']}')")
        return

    if args.matrizes:
        selecionadas = {nome.strip() for nome in args.matrizes.split(",")}
        matrizes_a_rodar = [m for m in cfg["matrizes"] if m["nome"] in selecionadas]
    else:
        matrizes_a_rodar = [m for m in cfg["matrizes"] if m.get("ativo")]

    if not matrizes_a_rodar:
        print("Nenhuma matriz ativa. Marque 'ativo: true' no config.yaml ou use --matrizes.")
        return

    df_base = carregar_base(cfg)
    print(f"Base carregada: {len(df_base)} linhas.")

    for matriz_cfg in matrizes_a_rodar:
        rodar_matriz(matriz_cfg["nome"], matriz_cfg, df_base, cfg)


if __name__ == "__main__":
    main()
