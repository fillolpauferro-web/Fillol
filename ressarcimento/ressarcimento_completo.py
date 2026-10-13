# -*- coding: utf-8 -*-
r"""
RESSARCIMENTO - Santa Cruz e Panpharma (arquivo único)
======================================================
1) Converte os relatórios Excel/CSV de
   C:\Users\I0507867.FARMA\OneDrive - Sanofi\Desktop\Ressarcimento
   em Parquet na subpasta \Parquet
2) Analisa desconto líquido e Margem OL (6% s/ bruto -> % s/ líquido)
3) Gera Analise_Desconto_SantaCruz_Panpharma.xlsx na pasta Ressarcimento

Funciona em: Spyder, VS Code, PyCharm, Jupyter (cole numa célula) ou
             Prompt/Anaconda Prompt:  python ressarcimento_completo.py

Pré-requisito (uma vez só):  pip install pandas pyarrow openpyxl python-calamine
"""


import argparse
import re
import sys
import unicodedata
from pathlib import Path

try:
    import pandas as pd
    import pyarrow  # noqa: F401
    import openpyxl  # noqa: F401
except ImportError as _e:
    sys.exit(f"Falta instalar uma biblioteca ({_e.name}). Rode no Prompt:\n"
             "    pip install pandas pyarrow openpyxl python-calamine")

# % sobre o Faturado Líquido NF para calcular o valor mensal (1,8% = 1.8)
PCT_LIQUIDO = 1.8
PASTA_PADRAO = r"C:\Users\I0507867.FARMA\OneDrive - Sanofi\Desktop\Ressarcimento\Parquet"
DISTRIBUIDORES = {"SANTA CRUZ": "SANTA CRUZ", "PANPHARMA": "PANPHARMA"}
# False = todos os distribuidores/clientes; True = só Santa Cruz e Panpharma
APENAS_SC_PANPHARMA = False
# Período do quadro "Projeto" (Mês de emissão da NF, formato AAAA-MM)
MES_INICIO = "2026-05"
MES_FIM = "2026-08"
# Arquivo DE PARA Cluster (na pasta Ressarcimento). None = procura pelo nome
ARQUIVO_DE_PARA = None
COLUNA_DE_PARA = "L"   # coluna com o De-Para (cluster) no arquivo DE PARA
# Rep. OL por cluster (PROCX do novo cluster). OL Líquido = Faturado Líquido (PF) x Rep. OL
REP_OL = {
    "EXCLUSIVOS": 0.12,
    "F. POPULAR": 0.12,
    "OUTROS": 0.15,
    "TOP SKUS": 0.17,
}
REP_OL_SEM_CLUSTER = 0.0   # % para EAN sem De-Para / cluster fora da lista acima
PCT_OL_BRUTO = 0.06        # 6%OL = Preço Fábrica[A] x Quantidade Faturada[B] x 6%


# --------------------------------------------------------------------------
# Normalização de colunas
# --------------------------------------------------------------------------
def norm(texto) -> str:
    """minúsculas, sem acento, só letras/números."""
    t = unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]", "", t.lower())


# Nome interno -> (nome no relatório, ocorrência). O relatório tem colunas
# repetidas (Produto, Faturado Bruto, Valor desconto, Margem OL); a ocorrência
# diz qual delas usar (0 = primeira, 1 = segunda).
COLUNAS = {
    "distribuidor": ("Distribuidor", 0),
    "filial": ("Filial", 0),
    "cliente": ("Cliente", 0),
    "cnpj": ("CNPJ", 0),
    "grupo_clientes": ("Grupo de clientes", 0),
    "uf": ("Sigla", 0),
    "tipo": ("Tipo", 0),
    "produto": ("Produto", 1),          # descrição (a 1ª coluna "Produto" é o código TO/...)
    "familia": ("Familia", 0),
    "divisao": ("Divisao", 0),
    "ean": ("EAN do produto", 0),
    "canal": ("Canal de Vendas", 0),
    "origem": ("Origem do pedido", 0),
    "data_nf": ("Data de Emissão da NF", 0),
    "pf": ("Preço Fábrica[A]", 0),
    "qtd_faturada": ("Quantidade Faturada[B]", 0),
    "fat_bruto_pf": ("Faturado Bruto (PF)[C]", 0),
    "desc_pedido_pct": ("Desconto pedido[D]", 0),
    "desc_portal": ("Valor desconto[E]", 0),
    "fat_liq_pf": ("Faturado Líquido (PF)", 0),
    "fat_bruto_nf": ("Faturado Bruto", 0),
    "desc_com_nf_pct": ("Desconto Com. faturado", 0),
    "desc_fin_nf_pct": ("Desconto Fin. faturado", 0),
    "desc_nf": ("Valor desconto", 0),   # NF (o do Portal é "Valor desconto[E]")
    "fat_liq_nf": ("Faturado Líquido", 0),
    "desc_entrada_pct": ("% Desconto entrada", 0),
    "margem_ol_pct": ("% Margem OL", 0),
    "margem_ol_rs": ("Margem OL", 1),   # R$ (a 1ª "margemol" é "% Margem OL")
    "ressarcimento": ("Ressarcimento", 0),
    "valor_repasse": ("Valor Repasse", 0),
    "desc_adicional": ("Valor Desconto Adicional", 0),
    "valor_icms": ("Valor ICMS", 0),
    "valor_st": ("Valor ST", 0),
    "devolucao": ("Valor a pagar sobre as quantidades devolvidas", 0),
}
NUMERICAS = [
    "pf", "qtd_faturada", "fat_bruto_pf", "desc_pedido_pct", "desc_portal",
    "fat_liq_pf", "fat_bruto_nf", "desc_com_nf_pct", "desc_fin_nf_pct", "desc_nf",
    "fat_liq_nf", "desc_entrada_pct", "margem_ol_pct", "margem_ol_rs", "ressarcimento",
    "valor_repasse", "desc_adicional", "valor_icms", "valor_st", "devolucao",
]
OBRIGATORIAS = ["distribuidor", "fat_bruto_pf", "fat_liq_nf"]
MARGEM_OL_PADRAO = 6.0
VERBOSO = False  # True mostra qual coluna do arquivo foi usada para cada campo

# Outros nomes aceitos (já normalizados: minúsculo, sem acento/espaço)
SINONIMOS = {
    "distribuidor": {"distribuidora", "nomedistribuidor", "razaosocialdistribuidor", "ol", "operadorlogistico"},
    "fat_bruto_pf": {"faturadobrutopf", "valorbrutopf", "brutopf", "faturamentobrutopf"},
    "fat_liq_nf": {"faturadoliquidonf", "valorliquidonf", "liquidonf", "faturamentoliquido",
                   "valorliquido", "faturadoliquidonotafiscal"},
    "fat_bruto_nf": {"faturadobrutonf", "valorbrutonf", "brutonf", "valorbruto"},
    "qtd_faturada": {"quantidadefaturada", "qtdefaturada", "qtdfaturada"},
    "pf": {"precofabrica", "pf"},
}


def _base(col: str) -> str:
    """Remove sufixos de duplicidade (.1, _1, ' 1') que pandas/conversores criam."""
    return re.sub(r"([._ ]\d+)$", "", str(col))


def mapear_colunas(df: pd.DataFrame) -> dict:
    """Retorna {nome_interno: nome_real_no_dataframe}."""
    ocorrencias: dict[str, list[str]] = {}
    for col in df.columns:
        chave = norm(col)
        ocorrencias.setdefault(chave, []).append(col)
        b = norm(_base(col))
        if b != chave:  # "Produto.1" também conta como ocorrência de "Produto"
            ocorrencias.setdefault(b, []).append(col)
    for k in ocorrencias:  # mantém a ordem original das colunas
        ocorrencias[k] = sorted(set(ocorrencias[k]), key=list(df.columns).index)

    mapa = {}
    for interno, (nome, idx) in COLUNAS.items():
        cands = ocorrencias.get(norm(nome), [])
        if len(cands) > idx:
            mapa[interno] = cands[idx]
        elif interno == "produto" and cands:
            mapa[interno] = cands[-1]

    # Não achou pelo nome exato: aceita nomes "parecidos", ex.
    # "('Portal Fidelize', 'Faturado Bruto (PF)[C]')", "DISTRIBUIDOR_NOME",
    # "Faturado Bruto PF", "Faturado Liquido NF"...
    usados = set(mapa.values())
    for interno, (nome, _) in COLUNAS.items():
        if interno in mapa or interno == "margem_ol_rs":
            continue
        alvo = norm(re.sub(r"\[.*?\]", "", nome))  # sem [A], [B]...
        cands = []
        for col in df.columns:
            n = norm(re.sub(r"\[.*?\]", "", str(col)))
            if col in usados or not alvo:
                continue
            if n.endswith(alvo) or n.startswith(alvo):
                cands.append((abs(len(n) - len(alvo)), col))
        # "CNPJ da Filial do Distribuidor" não é o Distribuidor
        if interno == "distribuidor":  # tem que ter nomes de distribuidor dentro
            cands = [c for c in cands if "cnpj" not in norm(c[1])
                     and df[c[1]].head(20000).astype(str).map(classificar_distribuidor).notna().any()]
        if cands:
            mapa[interno] = min(cands)[1]
            usados.add(mapa[interno])
    for interno, alts in SINONIMOS.items():
        if interno in mapa:
            continue
        for col in df.columns:
            if col not in usados and norm(col) in alts:
                mapa[interno] = col
                usados.add(col)
                break
    return mapa


def achar_distribuidor_pelo_conteudo(df: pd.DataFrame):
    """Coluna de texto que mais contém SANTA CRUZ / PANPHARMA."""
    melhor, maior = None, 0
    amostra = df.head(20000)
    for col in amostra.columns:
        if "cnpj" in norm(col):
            continue
        try:
            v = amostra[col].dropna().astype(str)
        except Exception:  # noqa: BLE001
            continue
        if v.empty or pd.api.types.is_numeric_dtype(amostra[col]):
            continue
        qtd = v.map(classificar_distribuidor).notna().sum()
        bonus = 1.5 if "distrib" in norm(col) else 1.0
        if qtd * bonus > maior:
            melhor, maior = col, qtd * bonus
    return melhor


def completar_colunas(df: pd.DataFrame) -> list[str]:
    """Cria as colunas que faltam a partir das que existem. Retorna o que foi criado."""
    criadas = []
    tem = lambda c: c in df.columns and df[c].abs().sum() > 0  # noqa: E731

    def criar(col, valor, como):
        df[col] = valor
        criadas.append(f"{col} = {como}")

    if not tem("fat_bruto_pf"):
        if tem("pf") and tem("qtd_faturada"):
            criar("fat_bruto_pf", df["pf"] * df["qtd_faturada"], "PF x Qtde faturada")
        elif tem("fat_liq_pf") and tem("desc_portal"):
            criar("fat_bruto_pf", df["fat_liq_pf"] + df["desc_portal"], "Líquido PF + Desconto portal")
        elif tem("fat_bruto_nf"):
            criar("fat_bruto_pf", df["fat_bruto_nf"], "Faturado Bruto NF")
    if not tem("desc_portal") and tem("fat_bruto_pf") and tem("desc_pedido_pct"):
        criar("desc_portal", df["fat_bruto_pf"] * df["desc_pedido_pct"] / 100,
              "Bruto PF x Desconto pedido")
    if not tem("fat_liq_pf") and tem("fat_bruto_pf"):
        criar("fat_liq_pf", df["fat_bruto_pf"] - df.get("desc_portal", 0), "Bruto PF - Desconto portal")
    if not tem("fat_bruto_nf") and tem("fat_liq_nf") and tem("desc_nf"):
        criar("fat_bruto_nf", df["fat_liq_nf"] + df["desc_nf"], "Líquido NF + Desconto NF")
    if not tem("fat_liq_nf"):
        if tem("fat_bruto_nf") and tem("desc_nf"):
            criar("fat_liq_nf", df["fat_bruto_nf"] - df["desc_nf"], "Bruto NF - Desconto NF")
        elif tem("fat_liq_pf"):
            criar("fat_liq_nf", df["fat_liq_pf"], "Faturado Líquido PF")
    if not tem("fat_bruto_nf"):
        criar("fat_bruto_nf", df["fat_bruto_pf"], "Faturado Bruto PF")
    if not tem("desc_nf"):
        criar("desc_nf", df["fat_bruto_nf"] - df["fat_liq_nf"], "Bruto NF - Líquido NF")
    if not tem("margem_ol_pct"):
        criar("margem_ol_pct", MARGEM_OL_PADRAO, f"{MARGEM_OL_PADRAO}% (padrão)")
    return criadas


def ajustar_cabecalho(df: pd.DataFrame) -> pd.DataFrame:
    """Se o parquet foi gerado com as linhas de título do Excel, acha a linha
    que contém 'Distribuidor' e usa como cabeçalho."""
    # colunas vindas de cabeçalho em 2 linhas: "('Portal Fidelize', 'Distribuidor')"
    def _ultimo(c):
        if isinstance(c, tuple):
            partes = [str(x) for x in c]
        else:
            m = re.match(r"^\((.*)\)$", str(c).strip())
            partes = re.findall(r"'([^']*)'", m.group(1)) if m else []
        partes = [x for x in partes if x and not x.startswith("Unnamed")]
        return partes[-1] if partes else c
    if any(_ultimo(c) != c for c in df.columns):
        df = df.copy()
        df.columns = [_ultimo(c) for c in df.columns]
    if any(norm(c) == "distribuidor" for c in df.columns):
        return df
    for i in range(min(50, len(df))):
        if any(norm(v) == "distribuidor" for v in df.iloc[i].tolist()):
            novo = df.iloc[i + 1:].copy()
            nomes, vistos = [], {}
            for v in df.iloc[i].tolist():
                v = str(v)
                n = vistos.get(v, 0)
                nomes.append(v if n == 0 else f"{v}.{n}")
                vistos[v] = n + 1
            novo.columns = nomes
            return novo.reset_index(drop=True)
    return df


def para_numero(s: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(s):
        return s.fillna(0)
    txt = s.astype(str).str.strip().str.replace("R$", "", regex=False)
    # formato brasileiro "1.234,56" -> 1234.56
    br = txt.str.contains(",", regex=False)
    txt = txt.where(~br, txt.str.replace(".", "", regex=False).str.replace(",", ".", regex=False))
    return pd.to_numeric(txt, errors="coerce").fillna(0)


# --------------------------------------------------------------------------
# Leitura
# --------------------------------------------------------------------------
def classificar_distribuidor(valor) -> str | None:
    v = norm(valor)
    for chave, rotulo in DISTRIBUIDORES.items():
        if norm(chave) in v:
            return rotulo
    return None


def ler_parquets(pasta: Path) -> pd.DataFrame:
    arquivos = sorted(pasta.rglob("*.parquet"))
    if not arquivos:
        sys.exit(f"Nenhum arquivo .parquet encontrado em: {pasta}")

    partes = []
    for arq in arquivos:
        try:
            bruto = None
            try:  # lê só as colunas usadas (bem mais rápido/leve)
                import pyarrow.parquet as pq
                nomes = pq.read_schema(arq).names
                mapa0 = mapear_colunas(pd.DataFrame(columns=nomes))
                if all(c in mapa0 for c in OBRIGATORIAS):
                    bruto = pd.read_parquet(arq, columns=list(dict.fromkeys(mapa0.values())))
            except Exception:  # noqa: BLE001
                bruto = None
            if bruto is None:
                bruto = ajustar_cabecalho(pd.read_parquet(arq))
        except Exception as e:  # noqa: BLE001
            print(f"  [ERRO] {arq.name}: {e}")
            continue
        mapa = mapear_colunas(bruto)
        if "distribuidor" not in mapa:
            col = achar_distribuidor_pelo_conteudo(bruto)
            if col is not None:
                mapa["distribuidor"] = col
                print(f"  [AVISO] {arq.name}: Distribuidor lido da coluna '{col}'")
        if "distribuidor" not in mapa:
            print(f"  [PULADO] {arq.name}: coluna Distribuidor não encontrada.")
            print(f"           Colunas do arquivo: {list(bruto.columns)}")
            continue
        if VERBOSO:
            print(f"  Mapeamento {arq.name}: {mapa}")

        df = pd.DataFrame({interno: bruto[real] for interno, real in mapa.items()})
        df["grupo_distribuidor"] = df["distribuidor"].map(classificar_distribuidor)
        if APENAS_SC_PANPHARMA:
            df = df[df["grupo_distribuidor"].notna()].copy()
        else:
            outros = df["distribuidor"].astype(str).str.strip().str.upper()
            df["grupo_distribuidor"] = df["grupo_distribuidor"].fillna(outros)
            df = df[~df["grupo_distribuidor"].isin(["", "NAN", "NONE"])].copy()
        if df.empty:
            print(f"  [OK] {arq.name}: 0 linhas")
            continue

        for c in NUMERICAS:
            df[c] = para_numero(df[c]) if c in df else 0.0
        criadas = completar_colunas(df)
        for c in criadas:
            print(f"  [CRIADA] {arq.name}: {c}")
        faltando = [c for c in OBRIGATORIAS[1:] if c not in df or df[c].abs().sum() == 0]
        if faltando:
            print(f"  [AVISO] {arq.name}: sem dados para {faltando} (ficarão zerados).")
            print(f"           Colunas do arquivo: {list(bruto.columns)}")
        df["arquivo"] = arq.name
        partes.append(df)
        print(f"  [OK] {arq.name}: {len(df):,} linhas")

    if not partes:
        sys.exit("Nenhuma linha encontrada nos parquets.")
    df = pd.concat(partes, ignore_index=True)

    # ---- Margem OL ----------------------------------------------------------
    # Ressarcimento = (Menor Desc. - Desc. Entrada + %OL) x PF x Qtd, ou seja,
    # a OL (6%) é calculada sobre o Faturado Bruto PF.
    ol_calc = df["fat_bruto_pf"] * df["margem_ol_pct"] / 100
    df["margem_ol_rs"] = df["margem_ol_rs"].where(df["margem_ol_rs"] != 0, ol_calc)
    # Quanto o distribuidor realmente ficou: vendeu (líquido NF) + recebeu
    # ressarcimento - custo de compra (PF com desconto de entrada).
    df["custo_distrib"] = df["fat_bruto_pf"] * (1 - df["desc_entrada_pct"] / 100)
    df["ol_realizada"] = df["fat_liq_nf"] + df["ressarcimento"] - df["custo_distrib"]

    if "data_nf" in df:
        datas = pd.to_datetime(df["data_nf"], errors="coerce", format="mixed", dayfirst=False)
        br = df["data_nf"].astype(str).str.match(r"^\d{2}/\d{2}/\d{4}")
        if br.any():
            datas[br] = pd.to_datetime(df.loc[br, "data_nf"].astype(str).str[:10],
                                       format="%d/%m/%Y", errors="coerce")
        df["data_nf"] = datas
        df["mes"] = df["data_nf"].dt.to_period("M").astype(str)
    else:
        df["mes"] = "N/D"
    return df


# --------------------------------------------------------------------------
# Indicadores
# --------------------------------------------------------------------------
SOMAS = [
    "qtd_faturada", "fat_bruto_pf", "desc_portal", "fat_liq_pf", "fat_bruto_nf",
    "desc_nf", "fat_liq_nf", "ressarcimento", "valor_repasse", "desc_adicional",
    "margem_ol_rs", "custo_distrib", "ol_realizada",
]


def _div(a, b):
    return (a / b.where(b != 0)).fillna(0)


def resumir(df: pd.DataFrame, por: list[str]) -> pd.DataFrame:
    por = [p for p in por if p in df.columns]
    r = df.groupby(por, dropna=False)[SOMAS].sum().reset_index()
    r.insert(len(por), "linhas", df.groupby(por, dropna=False).size().values)

    # Desconto efetivo do cliente medido contra o Preço Fábrica
    r["desc_liquido_pf"] = r["fat_bruto_pf"] - r["fat_liq_nf"]
    r["pct_desc_liquido_pf"] = _div(r["desc_liquido_pf"], r["fat_bruto_pf"])
    r["pct_desc_nf"] = _div(r["desc_nf"], r["fat_bruto_nf"])
    r["pct_desc_portal"] = _div(r["desc_portal"], r["fat_bruto_pf"])
    r["pct_ressarcimento_pf"] = _div(r["ressarcimento"], r["fat_bruto_pf"])
    # Parte do desconto ao cliente que é bancada pelo ressarcimento
    r["pct_desc_coberto_ressarc"] = _div(r["ressarcimento"], r["desc_liquido_pf"])

    # OL: os 6% calculados no bruto, quanto representam no líquido
    r["pct_ol_bruto"] = _div(r["margem_ol_rs"], r["fat_bruto_pf"])
    r["pct_ol_liquido"] = _div(r["margem_ol_rs"], r["fat_liq_nf"])
    r["pct_ol_real_bruto"] = _div(r["ol_realizada"], r["fat_bruto_pf"])
    r["pct_ol_real_liquido"] = _div(r["ol_realizada"], r["fat_liq_nf"])
    r["ol_cedida"] = r["margem_ol_rs"] - r["ol_realizada"]
    r["pct_ol_retida"] = _div(r["ol_realizada"], r["margem_ol_rs"])

    # X% (padrão 1,8%) do Faturado Líquido NF
    r["valor_pct_liq"] = r["fat_liq_nf"] * PCT_LIQUIDO / 100
    r["pct_liq_s_bruto"] = _div(r["valor_pct_liq"], r["fat_bruto_pf"])
    r["pct_liq_s_ol"] = _div(r["valor_pct_liq"], r["margem_ol_rs"])

    nomes = {
        "linhas": "Linhas",
        "qtd_faturada": "Qtde Faturada",
        "fat_bruto_pf": "Faturado Bruto PF (R$)",
        "desc_portal": "Desconto Portal (R$)",
        "fat_liq_pf": "Faturado Líquido PF (R$)",
        "fat_bruto_nf": "Faturado Bruto NF (R$)",
        "desc_nf": "Desconto NF (R$)",
        "fat_liq_nf": "Faturado Líquido NF (R$)",
        "ressarcimento": "Ressarcimento (R$)",
        "valor_repasse": "Valor Repasse (R$)",
        "desc_adicional": "Desconto Adicional (R$)",
        "desc_liquido_pf": "Desconto Líquido s/ PF (R$)",
        "pct_desc_liquido_pf": "% Desc. Líquido s/ PF",
        "pct_desc_nf": "% Desc. NF",
        "pct_desc_portal": "% Desc. Portal",
        "pct_ressarcimento_pf": "% Ressarc. s/ PF",
        "pct_desc_coberto_ressarc": "% Desc. coberto p/ Ressarc.",
        "margem_ol_rs": "OL s/ Bruto (R$)",
        "custo_distrib": "Custo Distribuidor (R$)",
        "ol_realizada": "OL Realizada (R$)",
        "ol_cedida": "OL Repassada ao Cliente (R$)",
        "pct_ol_bruto": "% OL s/ Bruto PF",
        "pct_ol_liquido": "% OL s/ Líquido NF",
        "pct_ol_real_bruto": "% OL Realizada s/ Bruto PF",
        "pct_ol_real_liquido": "% OL Realizada s/ Líquido NF",
        "pct_ol_retida": "% da OL Retida",
        "valor_pct_liq": f"{PCT_LIQUIDO:g}% do Líquido NF (R$)".replace(".", ","),
        "pct_liq_s_bruto": f"% equiv. s/ Bruto PF ({PCT_LIQUIDO:g}% liq.)".replace(".", ","),
        "pct_liq_s_ol": f"% da OL ({PCT_LIQUIDO:g}% liq.)".replace(".", ","),
        "grupo_distribuidor": "Distribuidor",
        "distribuidor": "Distribuidor (original)",
        "filial": "Filial", "mes": "Mês", "produto": "Produto",
        "familia": "Família", "divisao": "Divisão", "uf": "UF",
        "grupo_clientes": "Grupo de clientes", "cliente": "Cliente", "cnpj": "CNPJ",
        "canal": "Canal de Vendas",
    }
    return r.sort_values("fat_bruto_pf", ascending=False).rename(columns=nomes)


def total_geral(df: pd.DataFrame) -> pd.DataFrame:
    df = df.assign(_t="TOTAL")
    t = resumir(df, ["_t"]).rename(columns={"_t": "Distribuidor"})
    return pd.concat([resumir(df, ["grupo_distribuidor"]), t], ignore_index=True)


def media_mensal(df: pd.DataFrame) -> pd.DataFrame:
    """Valor de PCT_LIQUIDO% do líquido em cada mês e a média mensal por distribuidor."""
    m = df.groupby(["grupo_distribuidor", "mes"])[["fat_bruto_pf", "fat_liq_nf",
                                                   "margem_ol_rs"]].sum().reset_index()
    m["valor"] = m["fat_liq_nf"] * PCT_LIQUIDO / 100
    tot = m.groupby("mes")[["fat_bruto_pf", "fat_liq_nf", "margem_ol_rs", "valor"]].sum()
    m = pd.concat([m, tot.reset_index().assign(grupo_distribuidor="TOTAL")], ignore_index=True)

    linhas = []
    for dist, g in m.groupby("grupo_distribuidor", sort=False):
        g = g.sort_values("mes")
        for _, r in g.iterrows():
            linhas.append({"Distribuidor": dist, "Mês": r["mes"],
                           "Faturado Bruto PF (R$)": r["fat_bruto_pf"],
                           "Faturado Líquido NF (R$)": r["fat_liq_nf"],
                           "VALOR (R$)": r["valor"],
                           "OL s/ Bruto (R$)": r["margem_ol_rs"]})
        n = g["mes"].nunique()
        linhas.append({"Distribuidor": dist, "Mês": f"MÉDIA MENSAL ({n} meses)",
                       "Faturado Bruto PF (R$)": g["fat_bruto_pf"].sum() / n,
                       "Faturado Líquido NF (R$)": g["fat_liq_nf"].sum() / n,
                       "VALOR (R$)": g["valor"].sum() / n,
                       "OL s/ Bruto (R$)": g["margem_ol_rs"].sum() / n})
        linhas.append({"Distribuidor": dist, "Mês": "TOTAL DO PERÍODO",
                       "Faturado Bruto PF (R$)": g["fat_bruto_pf"].sum(),
                       "Faturado Líquido NF (R$)": g["fat_liq_nf"].sum(),
                       "VALOR (R$)": g["valor"].sum(),
                       "OL s/ Bruto (R$)": g["margem_ol_rs"].sum()})
    out = pd.DataFrame(linhas)
    out["% equiv. s/ Bruto PF"] = _div(out["VALOR (R$)"], out["Faturado Bruto PF (R$)"])
    out["% da OL"] = _div(out["VALOR (R$)"], out["OL s/ Bruto (R$)"])
    rot = f"{PCT_LIQUIDO:g}% do Líquido NF (R$)".replace(".", ",")
    return out.rename(columns={"VALOR (R$)": rot})


def imprimir_media_mensal(t: pd.DataFrame) -> None:
    rot = [c for c in t.columns if "% do Líquido NF" in c][0]
    print(f"\n{rot.replace(' (R$)', '').upper()} - POR MÊS E MÉDIA MENSAL")
    print("=" * 78)
    for dist, g in t.groupby("Distribuidor", sort=False):
        print(f"\n>> {dist}")
        for _, r in g.iterrows():
            print(f"   {r['Mês']:<24} Líquido R$ {r['Faturado Líquido NF (R$)']:>16,.2f}"
                  f"   -> R$ {r[rot]:>14,.2f}"
                  f"   ({r['% equiv. s/ Bruto PF']:.2%} do bruto)")
    print("=" * 78)


# --------------------------------------------------------------------------
# Saída
# --------------------------------------------------------------------------
def salvar_excel(abas: dict, caminho: Path) -> None:
    with pd.ExcelWriter(caminho, engine="openpyxl") as xw:
        for nome, tabela in abas.items():
            tabela = tabela.head(1_048_000)  # limite de linhas do Excel
            tabela.to_excel(xw, sheet_name=nome[:31], index=False)
            ws = xw.sheets[nome[:31]]
            ws.freeze_panes = "A2"
            for i, col in enumerate(tabela.columns, start=1):
                letra = ws.cell(row=1, column=i).column_letter
                ws.column_dimensions[letra].width = min(max(len(str(col)) + 2, 12), 45)
                fmt = "0.00%" if str(col).startswith("%") else (
                    "#,##0.00" if "(R$)" in str(col) else (
                        "#,##0" if col in ("Linhas", "Qtde Faturada") else None))
                if fmt:
                    for cel in ws.iter_cols(min_col=i, max_col=i, min_row=2):
                        for c in cel:
                            c.number_format = fmt


def imprimir_resumo(tot: pd.DataFrame) -> None:
    print("\n" + "=" * 78)
    print("DESCONTO LÍQUIDO REPASSADO POR DISTRIBUIDOR")
    print("=" * 78)
    for _, r in tot.iterrows():
        print(f"\n>> {r['Distribuidor']}")
        print(f"   Faturado Bruto PF ........ R$ {r['Faturado Bruto PF (R$)']:>18,.2f}")
        print(f"   Faturado Líquido NF ...... R$ {r['Faturado Líquido NF (R$)']:>18,.2f}")
        print(f"   Desconto Líquido s/ PF ... R$ {r['Desconto Líquido s/ PF (R$)']:>18,.2f}"
              f"   ({r['% Desc. Líquido s/ PF']:.2%})")
        print(f"   Desconto na NF ........... R$ {r['Desconto NF (R$)']:>18,.2f}"
              f"   ({r['% Desc. NF']:.2%} s/ bruto NF)")
        print(f"   Desconto Portal .......... R$ {r['Desconto Portal (R$)']:>18,.2f}"
              f"   ({r['% Desc. Portal']:.2%})")
        print(f"   Ressarcimento ............ R$ {r['Ressarcimento (R$)']:>18,.2f}"
              f"   ({r['% Ressarc. s/ PF']:.2%} s/ PF)")
        print("   --- Margem OL ---")
        print(f"   OL s/ Bruto PF ........... R$ {r['OL s/ Bruto (R$)']:>18,.2f}"
              f"   ({r['% OL s/ Bruto PF']:.2%} do bruto)")
        print(f"   -> no Líquido NF equivale a ............................"
              f"   {r['% OL s/ Líquido NF']:.2%} do líquido")
        print(f"   OL Realizada ............. R$ {r['OL Realizada (R$)']:>18,.2f}"
              f"   ({r['% OL Realizada s/ Bruto PF']:.2%} bruto | "
              f"{r['% OL Realizada s/ Líquido NF']:.2%} líquido)")
        print(f"   OL Repassada ao cliente .. R$ {r['OL Repassada ao Cliente (R$)']:>18,.2f}"
              f"   ({r['% da OL Retida']:.2%} da OL ficou com o distribuidor)")
    print("\n" + "=" * 78)


# --------------------------------------------------------------------------
# DE PARA Cluster (EAN -> De-Para)
# --------------------------------------------------------------------------
def norm_ean(v) -> str:
    """'7896420000000.0', '7,89642E+12', ' 07896420000000' -> '7896420000000'."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    if isinstance(v, (int, float)):
        return str(int(v))
    t = str(v).strip().strip('"').replace(" ", "")
    if re.fullmatch(r"[\d.,]+[eE][+-]?\d+", t):
        try:
            return str(int(float(t.replace(",", "."))))
        except ValueError:
            pass
    t = re.sub(r"\.0+$", "", t)
    return re.sub(r"\D", "", t).lstrip("0")


def achar_arquivo_de_para(origem: Path) -> Path | None:
    if ARQUIVO_DE_PARA:
        return Path(ARQUIVO_DE_PARA)
    for p in sorted(origem.iterdir()):
        if p.is_file() and not p.name.startswith("~$") and p.suffix.lower() in EXTENSOES \
                and "depara" in norm(p.stem):
            return p
    return None


def _letra_para_indice(letra: str) -> int:
    n = 0
    for ch in letra.upper():
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def ler_de_para(arq: Path) -> pd.DataFrame:
    """Retorna DataFrame [ean, de_para] (EAN único)."""
    if arq.suffix.lower() == ".csv":
        bruto = pd.read_csv(arq, sep=None, engine="python", header=None, dtype=str,
                            encoding="latin-1")
    else:
        bruto = pd.read_excel(arq, header=None, dtype=str, engine=_motor_excel(arq))
    # linha de cabeçalho = primeira que tem uma célula com "EAN"
    h = 0
    for i in range(min(30, len(bruto))):
        if any("ean" in norm(v) for v in bruto.iloc[i].tolist()):
            h = i
            break
    cab = [norm(v) for v in bruto.iloc[h].tolist()]
    dados = bruto.iloc[h + 1:]

    i_dp = next((i for i, c in enumerate(cab) if c == "depara"), None)
    if i_dp is None:
        i_dp = _letra_para_indice(COLUNA_DE_PARA)
    i_ean = next((i for i, c in enumerate(cab) if c in ("ean", "eandoproduto", "codigoean",
                                                        "codean", "eanproduto")), None)
    if i_ean is None:
        i_ean = next((i for i, c in enumerate(cab) if "ean" in c), None)
    if i_ean is None:  # coluna com mais valores de 13 dígitos começando com 789
        conta = [dados.iloc[:, i].map(norm_ean).str.match(r"^\d{12,14}$").sum()
                 for i in range(dados.shape[1])]
        i_ean = max(range(len(conta)), key=conta.__getitem__)

    dp = pd.DataFrame({"ean": dados.iloc[:, i_ean].map(norm_ean),
                       "de_para": dados.iloc[:, i_dp].astype(str).str.strip()})
    dp = dp[(dp["ean"] != "") & ~dp["de_para"].isin(["", "nan", "None"])]
    print(f"DE PARA: {arq.name} | EAN = coluna {i_ean + 1}, De-Para = coluna {i_dp + 1} "
          f"| {dp['ean'].nunique():,} EANs")
    return dp.drop_duplicates("ean")


def aplicar_de_para(df: pd.DataFrame, dp: pd.DataFrame | None) -> pd.DataFrame:
    df["ean_norm"] = df["ean"].map(norm_ean) if "ean" in df else ""
    if dp is None:
        df["de_para"] = "SEM DE PARA"
        return df
    df = df.merge(dp.rename(columns={"ean": "ean_norm"}), on="ean_norm", how="left")
    df["de_para"] = df["de_para"].fillna("SEM DE PARA")
    return df


def calcular_campos_projeto(df: pd.DataFrame) -> pd.DataFrame:
    """6%OL, Rep. OL (PROCX do cluster) e OL Líquido."""
    tabela = {norm(k): v for k, v in REP_OL.items()}
    df["ol_6"] = df["pf"] * df["qtd_faturada"] * PCT_OL_BRUTO
    df["rep_ol"] = df["de_para"].map(lambda c: tabela.get(norm(c), REP_OL_SEM_CLUSTER))
    df["ol_liquido"] = df["fat_liq_pf"] * df["rep_ol"]
    fora = sorted(set(df.loc[~df["de_para"].map(lambda c: norm(c) in tabela), "de_para"]))
    if fora:
        print(f"[AVISO] Clusters sem Rep. OL na tabela (usando {REP_OL_SEM_CLUSTER:.0%}): {fora}")
    return df


# --------------------------------------------------------------------------
# Quadro "Projeto" (formato da tabela dinâmica)
# --------------------------------------------------------------------------
MESES_PT = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
MESES_EXT = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto",
             "Setembro", "Outubro", "Novembro", "Dezembro"]


def rotulo_mes(m: str, extenso=False) -> str:
    a, mm = m.split("-")
    return MESES_EXT[int(mm) - 1] if extenso else f"{MESES_PT[int(mm) - 1]}.{a[2:]}"


def quadro_projeto(df: pd.DataFrame, divisor: float = 1.0) -> pd.DataFrame:
    """Distribuidor x De-Para com OL Líquido, %OL, Faturado Líquido PF e DDC."""
    g = (df.groupby(["grupo_distribuidor", "de_para"])[["ol_liquido", "fat_bruto_pf", "fat_liq_pf"]]
         .sum() / divisor).reset_index()
    linhas = []

    def linha(dist, dp, r, tipo):
        return {"Distribuidor": dist, "DE PARA": dp, "tipo": tipo,
                ".OL Líquido": r["ol_liquido"],
                ".OL Liq Novo": r["ol_liquido"] / r["fat_bruto_pf"] if r["fat_bruto_pf"] else 0,
                ".OL Sob Liq": r["ol_liquido"] / r["fat_liq_pf"] if r["fat_liq_pf"] else 0,
                ".Faturado Líquido (PF)": r["fat_liq_pf"],
                ddc_col(): r["fat_liq_pf"] * PCT_LIQUIDO / 100}

    ordem = g.groupby("grupo_distribuidor")["fat_liq_pf"].sum().sort_values(ascending=False).index
    for dist in ordem:
        gd = g[g["grupo_distribuidor"] == dist].sort_values("de_para")
        for i, (_, r) in enumerate(gd.iterrows()):
            linhas.append(linha(dist if i == 0 else "", r["de_para"], r, "item"))
        linhas.append(linha(f"{dist} Total", "", gd.sum(numeric_only=True), "total"))
    linhas.append(linha("Total Geral", "", g.sum(numeric_only=True), "geral"))
    return pd.DataFrame(linhas)


def ddc_col() -> str:
    return f".DDC {PCT_LIQUIDO:g}%"


def escrever_quadro(ws, q: pd.DataFrame, titulo: str, rodape_esq: str, meses_ano: int = 12):
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    azul = PatternFill("solid", fgColor="BDD7EE")
    fino = Side(style="thin", color="000000")
    cols = ["Distribuidor", "DE PARA", ".OL Líquido", ".OL Liq Novo", ".OL Sob Liq",
            ".Faturado Líquido (PF)", ddc_col()]
    fmts = [None, None, "#,##0.00", "0%", "0%", "#,##0.00", "#,##0.00"]
    ws["A1"] = titulo
    for j, c in enumerate(cols, start=1):
        cel = ws.cell(row=2, column=j, value=c)
        cel.font, cel.fill = Font(bold=True), azul
        cel.alignment = Alignment(horizontal="center")
    lin = 3
    for _, r in q.iterrows():
        for j, c in enumerate(cols, start=1):
            cel = ws.cell(row=lin, column=j, value=r[c])
            if fmts[j - 1]:
                cel.number_format = fmts[j - 1]
            if j in (4, 5):
                cel.alignment = Alignment(horizontal="center")
            if r["tipo"] != "item":
                cel.font = Font(bold=True)
                cel.border = Border(top=fino)
            if r["tipo"] == "geral":
                cel.fill = azul
            elif r["tipo"] == "item" and j == 1 and r[c]:
                cel.font = Font(bold=True)
        lin += 1
    total_ddc = q.loc[q["tipo"] == "geral", ddc_col()].iloc[0]
    vermelho = Font(bold=True, color="FF0000", size=12)
    ws.cell(row=lin, column=1, value=rodape_esq).font = vermelho
    ws.cell(row=lin, column=6, value="Anual Projetado:").font = vermelho
    ws.cell(row=lin, column=6).alignment = Alignment(horizontal="right")
    c = ws.cell(row=lin, column=7, value=total_ddc * meses_ano)
    c.font, c.number_format = vermelho, "#,##0.00"
    for letra, w in zip("ABCDEFG", (34, 16, 18, 14, 13, 24, 16)):
        ws.column_dimensions[letra].width = w
    ws.freeze_panes = "A3"


def gerar_projeto(df: pd.DataFrame, origem: Path, saida: Path) -> None:
    arq_dp = achar_arquivo_de_para(origem)
    if arq_dp is None:
        print("[AVISO] Arquivo DE PARA Cluster não encontrado na pasta; tudo ficará 'SEM DE PARA'.")
        dp = None
    else:
        dp = ler_de_para(arq_dp)
    df = calcular_campos_projeto(aplicar_de_para(df, dp))

    per = df[(df["mes"] >= MES_INICIO) & (df["mes"] <= MES_FIM)].copy()
    meses = sorted(per["mes"].dropna().unique())
    if per.empty:
        print(f"[AVISO] Nenhuma linha entre {MES_INICIO} e {MES_FIM}. "
              f"Meses disponíveis: {sorted(df['mes'].dropna().unique())}")
        return

    sem = per[per["de_para"] == "SEM DE PARA"]
    pct_sem = sem["fat_liq_pf"].sum() / per["fat_liq_pf"].sum() if per["fat_liq_pf"].sum() else 0
    print(f"\nDE PARA: {per['ean_norm'].nunique() - sem['ean_norm'].nunique():,} EANs com cluster, "
          f"{sem['ean_norm'].nunique():,} sem cluster ({pct_sem:.2%} do Faturado Líquido PF)")

    from openpyxl import Workbook
    wb = Workbook()
    wb.remove(wb.active)
    for m in meses:
        q = quadro_projeto(per[per["mes"] == m])
        escrever_quadro(wb.create_sheet(f"Projeto {rotulo_mes(m)}"), q, "Projeto",
                        f"{rotulo_mes(m, True)} mês Referência")
        tg = q[q["tipo"] == "geral"].iloc[0]
        print(f"  {rotulo_mes(m)}: Fat. Líquido PF R$ {tg['.Faturado Líquido (PF)']:,.2f} | "
              f"{ddc_col()} R$ {tg[ddc_col()]:,.2f} | Anual proj. R$ {tg[ddc_col()] * 12:,.2f}")
    n = len(meses)
    periodo = f"{rotulo_mes(meses[0])} a {rotulo_mes(meses[-1])}"
    q = quadro_projeto(per, divisor=n)
    escrever_quadro(wb.create_sheet("Projeto Média"), q, f"Projeto - média mensal {periodo}",
                    f"Média mensal {periodo} ({n} meses)")
    tg = q[q["tipo"] == "geral"].iloc[0]
    print(f"  MÉDIA ({n} meses): {ddc_col()} R$ {tg[ddc_col()]:,.2f}/mês | "
          f"Anual proj. R$ {tg[ddc_col()] * 12:,.2f}")
    q = quadro_projeto(per)
    escrever_quadro(wb.create_sheet("Projeto Acumulado"), q, f"Projeto - acumulado {periodo}",
                    f"Acumulado {periodo}", meses_ano=12 / n)

    # base por EAN para conferência do De-Para
    if "produto" not in per:
        per["produto"] = ""
    base = (per.groupby(["ean_norm", "produto", "de_para", "rep_ol"], dropna=False)
            [["fat_bruto_pf", "fat_liq_pf", "ol_6", "ol_liquido"]].sum().reset_index()
            .sort_values("fat_liq_pf", ascending=False))
    ws = wb.create_sheet("EAN x De-Para")
    ws.append(["EAN", "Produto", "De-Para", "Rep. OL", "Faturado Bruto (PF)[C]",
               "Faturado Líquido (PF)", "6%OL", "OL Líquido"])
    for r in base.itertuples(index=False):
        ws.append(list(r))
    ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 16, 45
    wb.save(saida)
    print(f"Excel do Projeto gerado: {saida}")


def analisar() -> None:
    global PCT_LIQUIDO
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--pasta", default=PASTA_PADRAO, help="pasta com os .parquet")
    ap.add_argument("--saida", default=None, help="arquivo Excel de saída")
    ap.add_argument("--origem", default=None, help="pasta Ressarcimento (onde está o DE PARA)")
    ap.add_argument("--pct", type=float, default=PCT_LIQUIDO,
                    help="%% sobre o líquido para o cálculo mensal (padrão 1.8)")
    args, _ = ap.parse_known_args()
    PCT_LIQUIDO = args.pct

    pasta = Path(args.pasta)
    saida = Path(args.saida) if args.saida else pasta.parent / "Analise_Desconto_SantaCruz_Panpharma.xlsx"

    print(f"Lendo parquets em: {pasta}")
    df = ler_parquets(pasta)
    print(f"\nTotal de linhas analisadas: {len(df):,}")

    origem = Path(args.origem) if args.origem else pasta.parent
    print("\n" + "=" * 78 + f"\nPROJETO {MES_INICIO} a {MES_FIM} (todos os distribuidores)\n" + "=" * 78)
    gerar_projeto(df, origem, origem / f"Projeto_OL_DDC_{MES_INICIO}_a_{MES_FIM}.xlsx")

    tot = total_geral(df)
    imprimir_resumo(tot)
    mensal = media_mensal(df)
    imprimir_media_mensal(mensal)

    abas = {
        "Resumo": tot,
        "Média Mensal %Líq": mensal,
        "Por Mês": resumir(df, ["grupo_distribuidor", "mes"]).sort_values(["Distribuidor", "Mês"]),
        "Por Filial": resumir(df, ["grupo_distribuidor", "filial"]),
        "Por Produto": resumir(df, ["grupo_distribuidor", "produto", "familia", "divisao"]),
        "Por Canal": resumir(df, ["grupo_distribuidor", "canal"]),
        "Por UF": resumir(df, ["grupo_distribuidor", "uf"]),
        "Por Grupo Clientes": resumir(df, ["grupo_distribuidor", "grupo_clientes"]),
        "Por Cliente (PDV)": resumir(df, ["grupo_distribuidor", "cnpj", "cliente", "uf"]),
    }
    salvar_excel(abas, saida)
    print(f"Excel gerado: {saida}")


# ==========================================================================
# CONVERSÃO EXCEL/CSV -> PARQUET
# ==========================================================================
import time

ORIGEM_PADRAO = r"C:\Users\I0507867.FARMA\OneDrive - Sanofi\Desktop\Ressarcimento"
DESTINO_PADRAO = ORIGEM_PADRAO + r"\Parquet"
EXTENSOES = {".xlsx", ".xlsm", ".xls", ".xlsb", ".csv"}


def _motor_excel(arq: Path) -> str | None:
    """calamine é ~10x mais rápido que openpyxl em arquivos grandes."""
    try:
        import python_calamine  # noqa: F401
        return "calamine"
    except ImportError:
        return "pyxlsb" if arq.suffix.lower() == ".xlsb" else None


def _linha_cabecalho(previa: pd.DataFrame) -> int:
    """Linha onde aparece 'Distribuidor' (no layout padrão é a 3ª)."""
    for i in range(len(previa)):
        if any(norm(v) == "distribuidor" for v in previa.iloc[i].tolist()):
            return i
    return -1


def ler_relatorio(arq: Path) -> pd.DataFrame:
    if arq.suffix.lower() == ".csv":
        for sep, enc in ((";", "utf-8-sig"), (";", "latin-1"), (",", "utf-8-sig")):
            try:
                previa = pd.read_csv(arq, sep=sep, encoding=enc, header=None, nrows=50, dtype=str)
                if previa.shape[1] > 5:
                    break
            except Exception:  # noqa: BLE001
                continue
        h = max(_linha_cabecalho(previa), 0)
        df = pd.read_csv(arq, sep=sep, encoding=enc, skiprows=h, dtype=str)
    else:
        motor = _motor_excel(arq)
        abas = pd.ExcelFile(arq, engine=motor).sheet_names
        aba, h = abas[0], 0
        for nome in abas:  # usa a aba que tiver o cabeçalho "Distribuidor"
            previa = pd.read_excel(arq, sheet_name=nome, header=None, nrows=50, engine=motor)
            linha = _linha_cabecalho(previa)
            if linha >= 0:
                aba, h = nome, linha
                break
        df = pd.read_excel(arq, sheet_name=aba, header=h, engine=motor)

    # remove rodapé "Fim do arquivo" e linhas vazias
    primeira = df.columns[0]
    df = df[~df[primeira].astype(str).str.strip().str.lower().eq("fim do arquivo")]
    df = df.dropna(how="all")
    return df


def preparar_para_parquet(df: pd.DataFrame) -> pd.DataFrame:
    """Colunas de texto/mistas viram string (pyarrow não aceita tipos mistos)."""
    df = df.copy()
    df.columns = [str(c) for c in df.columns]
    for c in df.columns:
        if df[c].dtype == "object":
            df[c] = df[c].map(lambda v: None if pd.isna(v) else str(v))
    return df


def converter(origem: Path, destino: Path, refazer: bool) -> None:
    destino.mkdir(parents=True, exist_ok=True)
    arquivos = sorted(p for p in origem.iterdir()
                      if p.is_file() and p.suffix.lower() in EXTENSOES
                      and not p.name.startswith("~$")
                      and not p.name.startswith(("Analise_Desconto", "Projeto_OL_DDC"))
                      and "depara" not in norm(p.stem))
    if not arquivos:
        print(f"Nenhum Excel/CSV encontrado em {origem}")
        return

    print(f"Convertendo {len(arquivos)} arquivo(s) de {origem}")
    for arq in arquivos:
        saida = destino / (arq.stem + ".parquet")
        if saida.exists() and not refazer and saida.stat().st_mtime >= arq.stat().st_mtime:
            print(f"  [JÁ EXISTE] {saida.name}")
            continue
        t0 = time.time()
        try:
            df = preparar_para_parquet(ler_relatorio(arq))
            df.to_parquet(saida, index=False)
            print(f"  [OK] {arq.name} -> {saida.name}: {len(df):,} linhas "
                  f"({time.time() - t0:.0f}s)")
        except Exception as e:  # noqa: BLE001
            print(f"  [ERRO] {arq.name}: {e}")


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--origem", default=ORIGEM_PADRAO)
    ap.add_argument("--destino", default=DESTINO_PADRAO)
    ap.add_argument("--refazer", action="store_true")
    ap.add_argument("--so-converter", action="store_true")
    args, extras = ap.parse_known_args()

    origem, destino = Path(args.origem), Path(args.destino)
    converter(origem, destino, args.refazer)
    if args.so_converter:
        return

    print("\n" + "#" * 78 + "\nANÁLISE\n" + "#" * 78)
    sys.argv = [sys.argv[0], "--pasta", str(destino), "--origem", str(origem),
                "--saida", str(origem / "Analise_Desconto.xlsx")] + extras
    analisar()
    print("\nConcluído.")


if __name__ == "__main__":
    main()
