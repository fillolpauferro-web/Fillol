"""Análise Erro Bandeira — script independente e dedicado (sem menu, sem
outras análises): compila a base de vendas e compara a Tabela de
negociação de cada pedido de bandeira física (Raia, Carrefour, DPSP,
Panvel, ...) com o guia "Painel x Tabela".

Passo a passo:
  1. Compila toda a base de vendas (aceita * no caminho pra juntar vários
     arquivos de uma vez).
  2. Filtra só os pedidos dos CNPJs cadastrados no arquivo de controle da
     Bandeira (controle_bandeira, ex.: Painel_Bandeira), trazendo as
     colunas configuradas em controle_bandeira.colunas_trazidas (ex.:
     id_bandeira, bandeira, perfil_bandeira, razao_social, cidade, estado).
  3. Compara a Tabela de negociação de cada pedido com a Tabela 1
     (Genérico) e a Tabela 2 (CA) do Grupo de clientes daquele pedido, de
     acordo com o guia "Painel x Tabela" (painel_tabela no config) — cada
     palavra da Tabela esperada precisa aparecer, como prefixo, em alguma
     palavra da Tabela real (mesma lógica usada nas outras análises pra
     cobrir abreviação/reordenação). Batendo com qualquer uma das duas,
     o pedido está correto e NÃO aparece em lugar nenhum.
  4. Pedidos que não batem com nenhuma das duas viram uma linha no
     relatório "Análise Erro Bandeira.xlsx", com os campos da base + as
     colunas do controle da Bandeira + a condição correta: a Tabela que
     deveria ter sido usada (Tabela 2/CA se o CNPJ estiver cadastrado como
     CA em algum item de cnpjs_ca, senão Tabela 1/Genérico) e o desconto
     correspondente em Condicao_comercial. Quando o Grupo de clientes nem
     consta no Painel x Tabela, a linha também entra, com a condição
     correta marcada como desconhecida — pra não esconder um caso não
     mapeado.

Uso:
    python analise_erro_bandeira.py
    python analise_erro_bandeira.py --config outro_config.yaml
"""

from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd
import yaml

from utils import (
    normalize_cnpj,
    normalize_ean,
    normalize_text,
    read_table_mais_recente,
    read_table_or_glob,
    remover_prefixo_tabela_agregadora,
    to_datetime,
    to_numeric,
    tokenizar,
)

BASE_DIR = Path(__file__).resolve().parent

TABELA_NAO_CADASTRADA = "Grupo de clientes não cadastrado no Painel x Tabela"


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
            "Ajuste base.colunas no config_erro_bandeira.yaml para os nomes reais da planilha."
        )

    df["_cnpj_norm"] = df[colunas["cnpj"]].map(normalize_cnpj)
    df["_ean_norm"] = df[colunas["ean"]].map(normalize_ean)
    # "Tabela Agregadora - X" é a mesma Tabela de negociação que "X" (versão
    # agregada) — pedidos lançados em qualquer uma das duas são igualmente
    # válidos. O texto original da coluna continua intacto na saída.
    df["_tabela_norm"] = (
        df[colunas["tabela_negociacao"]].map(normalize_text).map(remover_prefixo_tabela_agregadora)
    )
    df["_tabela_tokens"] = df["_tabela_norm"].map(tokenizar)
    df["_grupo_norm"] = df[colunas["grupo_clientes"]].map(normalize_text)
    df["_data_pedido"] = to_datetime(df[colunas["data_pedido"]])
    return df


def carregar_controle_bandeira(cfg: dict) -> pd.DataFrame:
    """Lê o arquivo de controle da Bandeira (ex.: Painel_Bandeira — lista de
    CNPJs de clientes físicos, com a bandeira/rede de cada um). Aceita nome
    de arquivo com curinga (ex.: "Painel_Bandeira*.xlsx") e usa sempre o
    mais recente.
    """
    ctrl_cfg = cfg["controle_bandeira"]
    caminho = BASE_DIR / ctrl_cfg["arquivo"]
    df = read_table_mais_recente(caminho, ctrl_cfg.get("aba"))

    chave = ctrl_cfg["chave"]
    colunas_trazidas = ctrl_cfg.get("colunas_trazidas") or {}
    faltando = [c for c in (chave, *colunas_trazidas.values()) if c not in df.columns]
    if faltando:
        raise KeyError(
            f"Colunas ausentes no arquivo de controle da Bandeira ({caminho.name}): {faltando}. "
            "Ajuste controle_bandeira no config_erro_bandeira.yaml."
        )

    df = df.rename(columns={v: k for k, v in colunas_trazidas.items()})
    df["_chave_controle_norm"] = df[chave].map(normalize_cnpj)

    colunas_saida = ["_chave_controle_norm", *colunas_trazidas.keys()]
    return df[colunas_saida].drop_duplicates(subset="_chave_controle_norm", keep="first")


def _tokens_ou_none(valor) -> tuple[str, ...] | None:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    tokens = tokenizar(normalize_text(valor))
    return tokens or None


def carregar_painel_tabela(cfg: dict) -> pd.DataFrame:
    """Lê o guia "Painel x Tabela": Grupo de clientes -> Tabela 1
    (Genérico, sempre preenchida) e Tabela 2 (CA, só quando esse grupo
    também tem canal CA). É o único critério pra saber se um pedido está
    na Tabela certa, e pra achar o desconto correto em Condicao_comercial
    pela Tabela que DEVERIA estar sendo usada.
    """
    painel_cfg = cfg["painel_tabela"]
    caminho = BASE_DIR / painel_cfg["arquivo"]
    df = read_table_mais_recente(caminho, painel_cfg.get("aba"))

    colunas = painel_cfg["colunas"]
    col_grupo = colunas["grupo_clientes"]
    col_tabela1 = colunas["tabela_1"]
    col_tabela2 = colunas["tabela_2"]
    faltando = [c for c in (col_grupo, col_tabela1, col_tabela2) if c not in df.columns]
    if faltando:
        raise KeyError(
            f"Colunas ausentes no Painel x Tabela ({caminho.name}): {faltando}. "
            "Ajuste painel_tabela.colunas no config_erro_bandeira.yaml."
        )

    df["_grupo_norm"] = df[col_grupo].map(normalize_text)
    df["_tabela1_texto"] = df[col_tabela1]
    df["_tabela2_texto"] = df[col_tabela2]
    df["_tabela1_tokens"] = df[col_tabela1].map(_tokens_ou_none)
    df["_tabela2_tokens"] = df[col_tabela2].map(_tokens_ou_none)
    colunas_saida = ["_grupo_norm", "_tabela1_texto", "_tabela2_texto", "_tabela1_tokens", "_tabela2_tokens"]
    return df[colunas_saida].drop_duplicates(subset="_grupo_norm", keep="first")


def carregar_cnpjs_ca(cfg: dict) -> set[str]:
    """União dos CNPJs cadastrados como "CA" em qualquer uma das fontes
    configuradas em cnpjs_ca (ex.: Painel_NV com rótulo NAO_VISITADO,
    rotulos_lojas.csv com SELL_OUT_CA/ATACAREJO_CONECTA_CA). Usado só pra
    decidir, no Painel x Tabela, se um pedido deve olhar a Tabela 1
    (Genérico) ou a Tabela 2 (CA). Uma fonte com arquivo ausente é
    ignorada (com aviso), sem travar as outras nem o resto da análise.
    """
    cnpjs: set[str] = set()
    for fonte in cfg.get("cnpjs_ca") or []:
        caminho = BASE_DIR / fonte["arquivo"]
        try:
            df = read_table_mais_recente(caminho, fonte.get("aba"))
        except FileNotFoundError as erro:
            print(f"!!! Aviso: fonte de cnpjs_ca ignorada ({erro})")
            continue

        chave = fonte["chave"]
        coluna_rotulo = fonte["coluna_rotulo"]
        faltando = [c for c in (chave, coluna_rotulo) if c not in df.columns]
        if faltando:
            raise KeyError(
                f"Colunas ausentes na fonte de cnpjs_ca ({Path(fonte['arquivo']).name}): {faltando}. "
                "Ajuste cnpjs_ca no config_erro_bandeira.yaml."
            )

        alvo = normalize_text(fonte["rotulo_valido"])
        df = df[df[coluna_rotulo].map(normalize_text) == alvo]
        cnpjs |= set(df[chave].map(normalize_cnpj))
    return cnpjs


def carregar_condicao_comercial(cfg: dict) -> pd.DataFrame:
    """Condicao_comercial traz o desconto correto por produto (EAN) e por
    Tabela de negociação (chave_tabela) — o mesmo EAN pode ter descontos
    diferentes conforme a Tabela (ex.: RAIA CA vs RAIA_GENERICO).
    """
    cc_cfg = cfg["condicao_comercial"]
    caminho = BASE_DIR / cc_cfg["arquivo"]
    df = read_table_or_glob(caminho, cc_cfg.get("aba"))

    col_ean = cc_cfg["colunas"]["chave_ean"]
    col_desconto = cc_cfg["colunas"]["desconto_correto_pct"]
    col_tabela = cc_cfg["colunas"]["chave_tabela"]

    faltando = [c for c in (col_ean, col_desconto, col_tabela) if c not in df.columns]
    if faltando:
        raise KeyError(
            f"Colunas ausentes em Condicao_comercial ({caminho.name}): {faltando}. "
            "Ajuste condicao_comercial.colunas no config_erro_bandeira.yaml."
        )

    df["_ean_norm"] = df[col_ean].map(normalize_ean)
    df["_desconto_correto_pct"] = to_numeric(df[col_desconto])
    # a coluna vem formatada como % no Excel (célula guarda 0.2398, exibe
    # "23,98%") — converte pra escala 0-100
    if cc_cfg.get("desconto_em_fracao"):
        df["_desconto_correto_pct"] = df["_desconto_correto_pct"] * 100

    df["_tabela_condicao_norm"] = df[col_tabela].map(normalize_text)
    df["_tabela_condicao_tokens"] = df["_tabela_condicao_norm"].map(tokenizar)

    colunas_saida = ["_ean_norm", "_tabela_condicao_tokens", "_desconto_correto_pct"]
    subset_dedup = ["_ean_norm", "_tabela_condicao_norm"]
    return df[colunas_saida + ["_tabela_condicao_norm"]].drop_duplicates(subset=subset_dedup, keep="first")[
        colunas_saida
    ]


def _rotulo_bate_na_tabela(esperado_tokens: tuple[str, ...] | None, tabela_tokens: tuple[str, ...]) -> bool:
    """Cada palavra da Tabela esperada precisa ser prefixo de alguma
    palavra da Tabela de negociação real, em qualquer ordem — cobre tanto
    abreviação ("AUT" prefixo de "AUTORIZADOR") quanto reordenação
    ("GENERICO_D1000" vs "D1000_GENERICO").
    """
    if not esperado_tokens:
        return False
    return all(any(tab.startswith(rot) for tab in tabela_tokens) for rot in esperado_tokens)


def _selecionar_desconto_por_tokens(
    ean_norm: pd.Series, tokens_esperados: pd.Series, df_condicao: pd.DataFrame
) -> pd.Series:
    """Escolhe o desconto correto linha a linha, comparando os tokens da
    Tabela esperada de cada pedido com os tokens de cada Tabela cadastrada
    em Condicao_comercial pro mesmo EAN.
    """
    base = pd.DataFrame({"_ean_norm": ean_norm.values, "_tokens_esperados": tokens_esperados.values})
    base["_idx_original"] = ean_norm.index
    candidatos = base.merge(df_condicao, how="left", on="_ean_norm")

    bate = [
        _rotulo_bate_na_tabela(esperado, tab) if isinstance(esperado, tuple) and isinstance(tab, tuple) else False
        for esperado, tab in zip(candidatos["_tokens_esperados"], candidatos["_tabela_condicao_tokens"])
    ]
    candidatos = candidatos[pd.Series(bate, index=candidatos.index)]
    candidatos = candidatos.drop_duplicates(subset="_idx_original", keep="first")

    resultado = pd.Series(float("nan"), index=ean_norm.index)
    resultado.loc[candidatos["_idx_original"]] = candidatos["_desconto_correto_pct"].to_numpy()
    return resultado


def calcular_erro_bandeira(df_base: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    """Filtra a base pelos CNPJs do controle da Bandeira e devolve só os
    pedidos cuja Tabela de negociação NÃO bate com a Tabela 1 nem com a
    Tabela 2 do Grupo de clientes (Painel x Tabela) — já com
    "tabela_correta" e "desconto_correto_pct" preenchidos.
    """
    df_controle = carregar_controle_bandeira(cfg)
    cnpjs_validos = set(df_controle["_chave_controle_norm"])
    df = df_base[df_base["_cnpj_norm"].isin(cnpjs_validos)].copy()
    if df.empty:
        return df

    df = df.merge(df_controle, how="left", left_on="_cnpj_norm", right_on="_chave_controle_norm")

    df_painel = carregar_painel_tabela(cfg)
    df = df.merge(df_painel, how="left", on="_grupo_norm")

    bate = pd.Series(
        [
            (isinstance(t1, tuple) and _rotulo_bate_na_tabela(t1, tab))
            or (isinstance(t2, tuple) and _rotulo_bate_na_tabela(t2, tab))
            for t1, t2, tab in zip(df["_tabela1_tokens"], df["_tabela2_tokens"], df["_tabela_tokens"])
        ],
        index=df.index,
    )
    df = df[~bate].copy()
    if df.empty:
        return df

    cnpjs_ca = carregar_cnpjs_ca(cfg)
    is_ca = df["_cnpj_norm"].isin(cnpjs_ca)

    def _escolher(t1_tok, t2_tok, t1_txt, t2_txt, ca: bool):
        if ca and isinstance(t2_tok, tuple):
            return t2_tok, t2_txt
        if isinstance(t1_tok, tuple):
            return t1_tok, t1_txt
        return None, TABELA_NAO_CADASTRADA

    escolhidos = [
        _escolher(t1, t2, t1t, t2t, ca)
        for t1, t2, t1t, t2t, ca in zip(
            df["_tabela1_tokens"], df["_tabela2_tokens"], df["_tabela1_texto"], df["_tabela2_texto"], is_ca
        )
    ]
    tokens_esperados = pd.Series([e[0] for e in escolhidos], index=df.index)
    df["tabela_correta"] = [e[1] for e in escolhidos]

    df_condicao = carregar_condicao_comercial(cfg)
    df["desconto_correto_pct"] = _selecionar_desconto_por_tokens(df["_ean_norm"], tokens_esperados, df_condicao)

    return df


def montar_saida(df: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    """Monta a saída com os campos pedidos, nessa ordem exata: campos da
    base (sem o desconto aplicado), colunas do controle da Bandeira
    (id_bandeira, bandeira, perfil_bandeira, razao_social, cidade, estado)
    e a condição correta (tabela_correta, desconto_correto_pct).
    """
    colunas = cfg["base"]["colunas"]
    colunas_base = [
        colunas["tabela_negociacao"],
        colunas["cnpj"],
        colunas["ean"],
        colunas["id_pedido"],
        colunas["tipo_cliente"],
        colunas["data_pedido"],
        colunas["faturado_liquido"],
        colunas["numero_nota"],
        colunas["quantidade_faturada"],
        colunas["distribuidor"],
        colunas["grupo_clientes"],
    ]
    colunas_trazidas = list((cfg["controle_bandeira"].get("colunas_trazidas") or {}).keys())
    colunas_condicao = ["tabela_correta", "desconto_correto_pct"]
    return df[colunas_base + colunas_trazidas + colunas_condicao].copy()


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument(
        "--config", default=str(BASE_DIR / "config_erro_bandeira.yaml"), help="Caminho do config"
    )
    args = parser.parse_args()

    cfg = carregar_config(args.config)

    df_base = carregar_base(cfg)
    print(f"Base carregada: {len(df_base)} linhas.")

    df_erros = calcular_erro_bandeira(df_base, cfg)
    if df_erros.empty:
        print("Nenhum pedido em Tabela errada — todos batem com o Painel x Tabela.")
        return

    df_saida = montar_saida(df_erros, cfg)
    colunas = cfg["base"]["colunas"]
    df_saida = df_saida.sort_values([colunas["grupo_clientes"], colunas["data_pedido"]]).reset_index(drop=True)
    print(f"{len(df_saida)} pedidos em Tabela errada.")

    pasta_saida = BASE_DIR / cfg["saida"]["pasta"]
    pasta_saida.mkdir(parents=True, exist_ok=True)
    nome_arquivo = cfg["saida"].get("nome_arquivo") or "Análise Erro Bandeira.xlsx"
    caminho_final = pasta_saida / nome_arquivo
    df_saida.to_excel(caminho_final, index=False)
    print(f"Resultado salvo em {caminho_final}")


if __name__ == "__main__":
    main()
