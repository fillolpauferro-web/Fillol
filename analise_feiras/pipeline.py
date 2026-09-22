"""Pipeline recorrente de análise de matrizes comerciais (Feira, Canal
Autorizador, Bandeira, ...). Cada matriz no config.yaml tem um "tipo" que
define como a base é filtrada e o que sai no resultado:

  tipo: "tabela" (ex.: Feira) — parte da coluna "Tabela de negociação":
    1. Filtra a base pela palavra-chave da matriz.
    2. PROCX (merge) com o arquivo de controle pelo CNPJ, trazendo as datas
       reais (Início Real / Término Real) em que a feira aconteceu.
    3. Check = OK só quando o CNPJ existe no controle E a data do pedido está
       dentro do período Início Real / Término Real.

  tipo: "cnpj" (ex.: Canal Autorizador, Sell Out, Atacarejo Conecta) — parte
  da relação de CNPJs:
    1. Filtra a base pelos CNPJs presentes no arquivo de controle (ex.:
       Painel_NV, rotulos_lojas — uma lista de clientes que deveriam comprar
       por um canal específico — não pela Tabela de negociação da base).
    2. Check = OK só quando a "Tabela de negociação" do pedido bate com a
       palavra_chave configurada (ex.: pedido de um CNPJ do Canal Autorizador
       feito em qualquer tabela diferente de "Canal Autorizador" = erro).
       Por padrão (correspondencia_tabela: "contem", ou omitido) a
       palavra_chave só precisa aparecer como substring na Tabela. Quando a
       palavra_chave é substring de outra Tabela válida (ex.: "Sell Out" é
       substring de "Sell Out CA") use correspondencia_tabela: "exata" — aí
       a comparação exige o mesmo conjunto de palavras (em qualquer ordem),
       não só a substring, pra não confundir as duas condições.

  Nos tipos "tabela" e "cnpj", pedidos com Check = Erro Operacional são
  cruzados com Condicao_comercial (por EAN, e por Tabela se a planilha tiver
  essa coluna) para achar o desconto correto, calcular o preço sem desconto
  e o preço líquido que deveria ter sido faturado — exceto quando o desconto
  comercial faturado do pedido é 0 (indício de dado ausente/errado): o Check
  continua Erro Operacional, mas o cálculo de preço/desconto fica em branco.

  tipo: "consolidacao" (ex.: Bandeira) — sem Check e sem cruzamento de
  desconto: só filtra a base pelos CNPJs do arquivo de controle (ex.:
  Painel_Bandeira) e consolida essas vendas num único arquivo, trazendo as
  colunas configuradas em colunas_trazidas (ex.: id_bandeira, bandeira,
  perfil_bandeira, razao_social, cidade, estado).

  tipo: "erro_bandeira" (ex.: Análise Erro Bandeira) — parte do mesmo
  arquivo de controle (ex.: Painel_Bandeira) e do guia "Painel x Tabela"
  (painel_tabela: Grupo de clientes -> Tabela 1 Genérico / Tabela 2 CA):
    1. Filtra a base pelos CNPJs do arquivo de controle, trazendo as
       colunas de colunas_trazidas (igual "consolidacao").
    2. Compara a Tabela de negociação de cada pedido com a Tabela 1 e a
       Tabela 2 do Grupo de clientes daquele pedido (Painel x Tabela) —
       cada palavra da Tabela esperada precisa aparecer, como prefixo, em
       alguma palavra da Tabela real (mesma lógica de abreviação/
       reordenação usada nas outras matrizes). Batendo com qualquer uma
       das duas, está correto e a linha NÃO entra no relatório.
    3. Pedidos que não batem com nenhuma das duas viram uma linha no
       relatório (nome_arquivo_saida — padrão "Análise Erro Bandeira.xlsx")
       com os campos da base + as colunas de colunas_trazidas + a condição
       correta: a Tabela que deveria ter sido usada (Tabela 2/CA se o CNPJ
       estiver cadastrado como CA em cnpjs_ca, senão Tabela 1/Genérico) e o
       desconto correspondente em Condicao_comercial. Exige
       base.colunas.grupo_clientes configurado. Quando o Grupo de clientes
       do pedido nem consta no Painel x Tabela, a linha também entra no
       relatório, com a condição correta marcada como desconhecida.

  tipo: "resumo_volume" (ex.: composição CA x WE) — sem Check, sem filtro
  por CNPJ, roda em cima da base inteira: classifica cada pedido em duas
  categorias a partir da Tabela de negociação (categoria_a se contém
  alguma das palavras_chave_categoria_a; categoria_b — "o resto" — senão) e
  soma quantidade de pedidos e faturado líquido de cada categoria, com o
  percentual de cada uma sobre o total. Também abre por mês: CA x WE mês a
  mês (com percentual dentro do mês) + média mensal por categoria; dentro
  de CA, o volume total e mês a mês de cada tabela específica isolada (ex.:
  só "Canal Autorizador") e qual delas puxou mais volume em cada mês. Saída
  é um arquivo com várias abas (Resumo, Mensal_CAxWE, Media_Mensal,
  CA_por_Tabela, CA_por_Tabela_Mes, Maior_Volume_por_Mes) — resumo, não um
  pedido por linha.

  tipo: "resumo_por_cnpj" (ex.: Raia e DPSP por CNPJ, Canal Autorizador por
  distribuidor) — sem Check, sem filtro por CNPJ de entrada, roda na base
  inteira: filtra só as tabelas de palavras_chave e acumula quantidade de
  pedidos, faturado líquido e quantidade faturada por CNPJ (ou por
  distribuidor, se coluna_agrupamento: "distribuidor"), dentro de cada
  tabela. Por padrão sai só o total acumulado; com abrir_por_mes: true, sai
  também a mesma quebra mês a mês, num arquivo com abas "Total" e "Mensal".

  tipo: "resumo_cnpj_grupo" — sem Check, sem filtro por CNPJ de entrada,
  filtra exclusivamente pela palavra_chave da matriz (ex.: "Canal
  Autorizador"): agrupa por CNPJ + Grupo de clientes + mês, com quantidade
  de pedidos, faturado líquido e quantidade faturada de cada mês. Exige
  base.colunas.grupo_clientes.

  Um arquivo de saída é salvo por matriz (nome_arquivo_saida no config, ou
  "{nome}_analise.xlsx" por padrão).

Uso:
    python pipeline.py                       # pergunta interativamente quais matrizes rodar
    python pipeline.py --todas               # roda todas as matrizes ativas sem perguntar
    python pipeline.py --matrizes Feira       # roda só "Feira", ignora o config
    python pipeline.py --listar               # lista as matrizes configuradas e sai
    python pipeline.py --config outro.yaml
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
    read_table,
    read_table_mais_recente,
    read_table_or_glob,
    remover_prefixo_tabela_agregadora,
    to_datetime,
    to_numeric,
    tokenizar,
)

BASE_DIR = Path(__file__).resolve().parent

CHECK_OK = "OK"
CHECK_ERRO = "Erro Operacional"


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
    df["_ean_norm"] = df[colunas["ean"]].map(normalize_ean)
    # "Tabela Agregadora - X" é a mesma Tabela de negociação que "X" (versão
    # agregada) — pedidos lançados em qualquer uma das duas são igualmente
    # válidos, então tiramos esse prefixo antes de qualquer comparação
    # (Feira, Canal Autorizador, regra da Bandeira). O texto original da
    # coluna "Tabela de negociação" continua intacto na saída.
    df["_tabela_norm"] = (
        df[colunas["tabela_negociacao"]].map(normalize_text).map(remover_prefixo_tabela_agregadora)
    )
    df["_data_pedido"] = to_datetime(df[colunas["data_pedido"]])
    df["_faturado_liquido"] = to_numeric(df[colunas["faturado_liquido"]])
    df["_desconto_aplicado_pct"] = to_numeric(df[colunas["desconto_aplicado_pct"]])
    if "quantidade_faturada" in colunas:
        df["_quantidade_faturada"] = to_numeric(df[colunas["quantidade_faturada"]])
    if "distribuidor" in colunas:
        df["_distribuidor_norm"] = df[colunas["distribuidor"]].map(normalize_text)
    return df


def filtrar_matriz(df_base: pd.DataFrame, palavra_chave: str) -> pd.DataFrame:
    """tipo: "tabela" — filtra pela palavra-chave na coluna Tabela de negociação."""
    chave = normalize_text(palavra_chave)
    mask = df_base["_tabela_norm"].str.contains(chave, na=False)
    return df_base.loc[mask].copy()


def filtrar_por_cnpj(df_base: pd.DataFrame, cnpjs_norm: set[str]) -> pd.DataFrame:
    """tipo: "cnpj" — filtra pelos CNPJs presentes no arquivo de controle,
    independente da Tabela de negociação do pedido.
    """
    mask = df_base["_cnpj_norm"].isin(cnpjs_norm)
    return df_base.loc[mask].copy()


def _mask_tabela_bate(tabela_norm: pd.Series, palavra_chave: str, correspondencia: str = "contem") -> pd.Series:
    """Compara a coluna "_tabela_norm" (já normalizada, sem prefixo "Tabela
    Agregadora -") com a palavra_chave configurada.

    "contem" (padrão): a palavra_chave aparece como substring — ok quando ela
    não é substring de nenhuma outra Tabela válida (ex.: "Canal Autorizador").

    "exata": exige o mesmo conjunto de palavras (tokenizado, em qualquer
    ordem) — necessário quando a palavra_chave É substring de outra Tabela
    válida (ex.: "Sell Out" é substring de "Sell Out CA"; "contem" misturaria
    as duas condições).
    """
    if correspondencia == "exata":
        tokens_chave = frozenset(tokenizar(normalize_text(palavra_chave)))
        return tabela_norm.map(lambda t: frozenset(tokenizar(t)) == tokens_chave)
    palavra = normalize_text(palavra_chave)
    return tabela_norm.str.contains(palavra, na=False)


def carregar_controle(matriz_cfg: dict) -> pd.DataFrame:
    """Lê o arquivo de controle da matriz (ex.: Controle_Feiras, Painel_NV) e
    prepara a coluna-chave (CNPJ ajustado) mais as colunas extras a trazer
    (ex.: Início Real / Término Real). Aceita nome de arquivo com curinga
    (ex.: "Painel_NV_*.xlsx") e usa sempre o mais recente.

    Se a matriz configurar coluna_rotulo_controle/rotulo_valido_controle
    (caso do Painel_NV, que traz vários rótulos no mesmo arquivo), filtra o
    controle só pelas linhas com esse rótulo antes de montar a lista de CNPJs.
    """
    caminho = BASE_DIR / matriz_cfg["arquivo_controle"]
    df = read_table_mais_recente(caminho, matriz_cfg.get("aba_controle"))

    chave_col = matriz_cfg["chave_controle"]
    colunas_trazidas = matriz_cfg.get("colunas_trazidas") or {}
    colunas_data = matriz_cfg.get("colunas_data", [])
    coluna_rotulo = matriz_cfg.get("coluna_rotulo_controle")
    rotulo_valido = matriz_cfg.get("rotulo_valido_controle")

    obrigatorias = [chave_col, *colunas_trazidas.values()]
    if coluna_rotulo:
        obrigatorias.append(coluna_rotulo)
    faltando = [c for c in obrigatorias if c not in df.columns]
    if faltando:
        raise KeyError(
            f"Colunas ausentes no arquivo de controle ({caminho.name}): {faltando}. "
            "Ajuste chave_controle/colunas_trazidas/coluna_rotulo_controle no config.yaml."
        )

    if coluna_rotulo:
        alvo = normalize_text(rotulo_valido)
        df = df[df[coluna_rotulo].map(normalize_text) == alvo]

    df = df.rename(columns={v: k for k, v in colunas_trazidas.items()})
    df["_chave_controle_norm"] = df[chave_col].map(normalize_cnpj)

    for col in colunas_data:
        df[col] = to_datetime(df[col])

    colunas_saida = ["_chave_controle_norm", *colunas_trazidas.keys()]
    df_controle = df[colunas_saida].drop_duplicates(subset="_chave_controle_norm", keep="first")
    return df_controle


def aplicar_procx_cnpj(df_matriz: pd.DataFrame, df_controle: pd.DataFrame) -> pd.DataFrame:
    """Verifica se o CNPJ do pedido existe no controle da feira (tipo PROCX)
    e traz as colunas extras configuradas quando existir.
    """
    return df_matriz.merge(
        df_controle,
        how="left",
        left_on="_cnpj_norm",
        right_on="_chave_controle_norm",
        indicator="_encontrado_no_controle",
    )


def calcular_check(df: pd.DataFrame, matriz_cfg: dict) -> pd.Series:
    """tipo: "tabela" (ex.: Feira) — Check = OK só quando o CNPJ existe no
    controle E o pedido foi feito dentro do período real (inicio_real a
    termino_real, comparando só a data — sem o horário).

    tipo: "cnpj" (ex.: Canal Autorizador) — a base já foi filtrada pelos
    CNPJs do controle, então Check = OK só quando a Tabela de negociação do
    pedido bate com a palavra_chave configurada; qualquer outra tabela
    (pedido de um CNPJ do canal, mas lançado em condição diferente) = erro.
    """
    encontrado = df["_encontrado_no_controle"] == "both"
    tipo = matriz_cfg.get("tipo", "tabela")

    if tipo == "cnpj":
        correspondencia = matriz_cfg.get("correspondencia_tabela", "contem")
        tabela_correta = _mask_tabela_bate(df["_tabela_norm"], matriz_cfg["palavra_chave"], correspondencia)
        ok = encontrado & tabela_correta
    else:
        colunas_trazidas = matriz_cfg["colunas_trazidas"]
        tem_periodo = "inicio_real" in colunas_trazidas and "termino_real" in colunas_trazidas
        if tem_periodo:
            data_pedido = df["_data_pedido"].dt.normalize()
            inicio = df["inicio_real"].dt.normalize()
            termino = df["termino_real"].dt.normalize()
            dentro_periodo = (data_pedido >= inicio) & (data_pedido <= termino)
            ok = encontrado & dentro_periodo
        else:
            ok = encontrado

    return pd.Series(CHECK_OK, index=df.index).mask(~ok, CHECK_ERRO)


def carregar_condicao_comercial(cfg: dict) -> pd.DataFrame:
    """Condicao_comercial traz o desconto correto por produto (EAN). Se a
    planilha tiver a coluna "Tabela" (condicao_comercial.colunas.chave_tabela
    configurada), o desconto correto passa a variar por Tabela de negociação
    também — nesse caso, aplicar_condicao_correta filtra pela tabela certa de
    cada matriz antes de casar por EAN.
    """
    cc_cfg = cfg["condicao_comercial"]
    caminho = BASE_DIR / cc_cfg["arquivo"]
    df = read_table(caminho, cc_cfg.get("aba"))

    col_ean = cc_cfg["colunas"]["chave_ean"]
    col_desconto = cc_cfg["colunas"]["desconto_correto_pct"]
    col_tabela = cc_cfg["colunas"].get("chave_tabela")

    obrigatorias = [col_ean, col_desconto, *([col_tabela] if col_tabela else [])]
    faltando = [c for c in obrigatorias if c not in df.columns]
    if faltando:
        raise KeyError(
            f"Colunas ausentes em Condicao_comercial ({caminho.name}): {faltando}. "
            "Ajuste condicao_comercial.colunas no config.yaml."
        )

    df["_ean_norm"] = df[col_ean].map(normalize_ean)
    df["_desconto_correto_pct"] = to_numeric(df[col_desconto])

    # a coluna vem formatada como % no Excel (célula guarda 0.2398, exibe
    # "23,98%") — converte pra escala 0-100 igual ao desconto aplicado da base
    if cc_cfg.get("desconto_em_fracao"):
        df["_desconto_correto_pct"] = df["_desconto_correto_pct"] * 100

    if col_tabela:
        df["_tabela_condicao_norm"] = df[col_tabela].map(normalize_text)
        # tokens, não só o texto — permite achar a Tabela certa por
        # correspondência de palavras (igual Rotulo x Tabela de negociação),
        # necessário quando a Tabela "correta" varia linha a linha (Bandeira)
        df["_tabela_condicao_tokens"] = df["_tabela_condicao_norm"].map(tokenizar)
        colunas_saida = ["_ean_norm", "_tabela_condicao_norm", "_tabela_condicao_tokens", "_desconto_correto_pct"]
        subset_dedup = ["_ean_norm", "_tabela_condicao_norm"]
    else:
        colunas_saida = ["_ean_norm", "_desconto_correto_pct"]
        subset_dedup = ["_ean_norm"]

    return df[colunas_saida].drop_duplicates(subset=subset_dedup, keep="first")


def _selecionar_desconto_por_tokens(
    ean_norm: pd.Series, tokens_esperados: pd.Series, df_condicao: pd.DataFrame
) -> pd.Series:
    """Escolhe o desconto correto linha a linha, comparando os tokens da
    Tabela esperada daquela linha (tokens_esperados) com os tokens de cada
    Tabela cadastrada em Condicao_comercial pro mesmo EAN — necessário
    quando a Tabela "correta" varia linha a linha (ex.: Bandeira, onde cada
    Grupo de clientes tem sua própria Tabela esperada, ao contrário de
    Feira/Canal Autorizador, que têm uma única palavra_chave fixa pra toda
    a matriz). Usa a mesma lógica de _rotulo_bate_na_tabela (cada palavra
    esperada como prefixo de alguma palavra da Tabela, em qualquer ordem).
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


def aplicar_condicao_correta(df: pd.DataFrame, df_condicao: pd.DataFrame, matriz_cfg: dict) -> pd.DataFrame:
    if "_tabela_condicao_norm" in df_condicao.columns:
        # matrizes com uma única palavra_chave fixa pra matriz toda (Feira,
        # Canal Autorizador, Sell Out, ...) filtram a condição comercial por
        # essa Tabela antes de casar por EAN, evitando pegar o desconto de
        # outra Tabela do mesmo produto.
        palavra = matriz_cfg.get("palavra_chave")
        if palavra:
            correspondencia = matriz_cfg.get("correspondencia_tabela", "contem")
            mask = _mask_tabela_bate(df_condicao["_tabela_condicao_norm"], palavra, correspondencia)
            df_condicao = df_condicao[mask]
        df_condicao = df_condicao[["_ean_norm", "_desconto_correto_pct"]].drop_duplicates(
            subset="_ean_norm", keep="first"
        )

    df = df.merge(df_condicao, how="left", on="_ean_norm")

    # preço sem desconto = líquido faturado / (1 - desconto aplicado)
    fator_aplicado = 1 - (df["_desconto_aplicado_pct"] / 100)
    df["preco_sem_desconto"] = (df["_faturado_liquido"] / fator_aplicado).where(fator_aplicado != 0)

    # preço líquido que deveria ter sido faturado com o desconto correto
    fator_correto = 1 - (df["_desconto_correto_pct"] / 100)
    df["preco_liquido_desconto_correto"] = df["preco_sem_desconto"] * fator_correto

    df["diferenca_faturamento"] = df["_faturado_liquido"] - df["preco_liquido_desconto_correto"]

    # só faz sentido preencher esses cálculos para linhas de erro operacional
    # com desconto comercial faturado real (!= 0) — desconto aplicado 0
    # normalmente indica dado ausente/errado na base, não "sem desconto de
    # fato", então o pedido continua Erro Operacional mas sem o cálculo
    erro_operacional = df["Check"] == CHECK_ERRO
    calcular = erro_operacional & (df["_desconto_aplicado_pct"] != 0)
    for col in (
        "_desconto_correto_pct",
        "preco_sem_desconto",
        "preco_liquido_desconto_correto",
        "diferenca_faturamento",
    ):
        df[col] = df[col].where(calcular)

    df = df.rename(columns={"_desconto_correto_pct": "desconto_correto_pct"})
    return df


def montar_saida(df: pd.DataFrame, cfg: dict, matriz_cfg: dict) -> pd.DataFrame:
    colunas_base_originais = list(cfg["base"]["colunas"].values())
    colunas_novas = [
        *(matriz_cfg.get("colunas_trazidas") or {}).keys(),
        "Check",
        "desconto_correto_pct",
        "preco_sem_desconto",
        "preco_liquido_desconto_correto",
        "diferenca_faturamento",
    ]
    return df[colunas_base_originais + colunas_novas].copy()


def montar_saida_consolidacao(df: pd.DataFrame, cfg: dict, matriz_cfg: dict) -> pd.DataFrame:
    """tipo: "consolidacao" — sem Check nem cruzamento de desconto, só as
    vendas encontradas mais as colunas trazidas do controle (ex.: Bandeira).
    """
    colunas_base_originais = list(cfg["base"]["colunas"].values())
    colunas_novas = list((matriz_cfg.get("colunas_trazidas") or {}).keys())
    return df[colunas_base_originais + colunas_novas].copy()


def _rotulo_bate_na_tabela(rotulo_tokens: tuple[str, ...], tabela_tokens: tuple[str, ...]) -> bool:
    """Cada palavra do Rotulo precisa ser prefixo de alguma palavra da
    Tabela de negociação, em qualquer ordem — cobre tanto abreviação
    ("AUT" prefixo de "AUTORIZADOR") quanto reordenação ("GENERICO_D1000"
    vs "D1000_GENERICO").
    """
    if not rotulo_tokens:
        return False
    return all(any(tab.startswith(rot) for tab in tabela_tokens) for rot in rotulo_tokens)


def _tokens_ou_none(valor) -> tuple[str, ...] | None:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    tokens = tokenizar(normalize_text(valor))
    return tokens or None


def carregar_painel_tabela(painel_cfg: dict) -> pd.DataFrame:
    """Lê o "Painel x Tabela" (matriz.painel_tabela no config.yaml) — guia
    oficial de qual Tabela de negociação deveria estar alocada pra cada
    Grupo de clientes: Tabela 1 (Genérico, sempre preenchida) e Tabela 2
    (CA, só quando esse grupo também tem canal CA). É o único critério pra
    decidir se um pedido está na Tabela certa (tipo "erro_bandeira") e pra
    achar o desconto correto em Condicao_comercial pela Tabela que DEVERIA
    estar sendo usada — nunca "a primeira condição cadastrada pro EAN", que
    não sabia diferenciar duas Tabelas do mesmo Grupo (ex.: RAIA CA vs
    RAIA_GENERICO, descontos diferentes pro mesmo EAN).
    """
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
            "Ajuste painel_tabela.colunas no config.yaml."
        )

    df["_grupo_norm"] = df[col_grupo].map(normalize_text)
    df["_tabela1_texto"] = df[col_tabela1]
    df["_tabela2_texto"] = df[col_tabela2]
    df["_tabela1_tokens"] = df[col_tabela1].map(_tokens_ou_none)
    df["_tabela2_tokens"] = df[col_tabela2].map(_tokens_ou_none)
    colunas_saida = ["_grupo_norm", "_tabela1_texto", "_tabela2_texto", "_tabela1_tokens", "_tabela2_tokens"]
    return df[colunas_saida].drop_duplicates(subset="_grupo_norm", keep="first")


def carregar_cnpjs_ca(cnpjs_ca_cfg: list[dict]) -> set[str]:
    """União dos CNPJs cadastrados como "CA" em qualquer uma das fontes
    configuradas em regra.cnpjs_ca (ex.: Painel_NV com rótulo NAO_VISITADO,
    rotulos_lojas.csv com SELL_OUT_CA/ATACAREJO_CONECTA_CA — os mesmos
    arquivos/rótulos das matrizes CanalAutorizador/SellOut CA/Atacarejo
    Conecta CA). Usado só pra decidir, no Painel x Tabela, se um pedido
    deve olhar a Tabela 1 (Genérico) ou a Tabela 2 (CA).
    """
    cnpjs: set[str] = set()
    for fonte_cfg in cnpjs_ca_cfg:
        try:
            df_controle = carregar_controle(fonte_cfg)
        except FileNotFoundError as erro:
            # uma fonte de cnpjs_ca faltando (ex.: rotulos_lojas.csv ainda
            # não copiado pra pasta de dados) não pode travar o desconto
            # correto de bandeiras que nem dependem dela (ex.: Carrefour,
            # Coop) — avisa e segue só com as fontes que existem.
            print(f"!!! Aviso: fonte de cnpjs_ca ignorada ({erro})")
            continue
        cnpjs |= set(df_controle["_chave_controle_norm"])
    return cnpjs


def calcular_erro_bandeira(df_base: pd.DataFrame, cfg: dict, matriz_cfg: dict) -> pd.DataFrame:
    """tipo: "erro_bandeira" — filtra a base pelos CNPJs do arquivo de
    controle (ex.: Painel_Bandeira, igual "consolidacao") e compara a
    Tabela de negociação de cada pedido com o guia "Painel x Tabela"
    (Grupo de clientes -> Tabela 1 Genérico / Tabela 2 CA): batendo com
    qualquer uma das duas (cada palavra da Tabela esperada como prefixo de
    alguma palavra da Tabela real, em qualquer ordem), o pedido está
    correto e não entra no retorno. Devolve só as linhas erradas, já com
    "tabela_correta" (a Tabela que deveria ter sido usada — Tabela 2/CA se
    o CNPJ estiver cadastrado como CA em cnpjs_ca, senão Tabela 1/Genérico;
    "Grupo de clientes não cadastrado no Painel x Tabela" quando nem isso
    dá pra saber) e "desconto_correto_pct" (a linha correspondente em
    Condicao_comercial).
    """
    colunas = cfg["base"]["colunas"]
    col_grupo = colunas.get("grupo_clientes")
    if not col_grupo:
        raise KeyError(
            "Configure base.colunas.grupo_clientes no config.yaml pra usar a matriz tipo 'erro_bandeira'."
        )

    df_controle = carregar_controle(matriz_cfg)
    cnpjs_validos = set(df_controle["_chave_controle_norm"])
    df = filtrar_por_cnpj(df_base, cnpjs_validos)
    if df.empty:
        return df

    df = aplicar_procx_cnpj(df, df_controle)
    df["_tabela_tokens"] = df["_tabela_norm"].map(tokenizar)
    df["_grupo_norm"] = df[col_grupo].map(normalize_text)

    df_painel = carregar_painel_tabela(matriz_cfg["painel_tabela"])
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

    cnpjs_ca = carregar_cnpjs_ca(matriz_cfg.get("cnpjs_ca") or [])
    is_ca = df["_cnpj_norm"].isin(cnpjs_ca)

    def _escolher(t1_tok, t2_tok, t1_txt, t2_txt, ca: bool):
        if ca and isinstance(t2_tok, tuple):
            return t2_tok, t2_txt
        if isinstance(t1_tok, tuple):
            return t1_tok, t1_txt
        return None, "Grupo de clientes não cadastrado no Painel x Tabela"

    escolhidos = [
        _escolher(t1, t2, t1t, t2t, ca)
        for t1, t2, t1t, t2t, ca in zip(
            df["_tabela1_tokens"], df["_tabela2_tokens"], df["_tabela1_texto"], df["_tabela2_texto"], is_ca
        )
    ]
    tokens_esperados = pd.Series([e[0] for e in escolhidos], index=df.index)
    df["tabela_correta"] = [e[1] for e in escolhidos]

    df_condicao = carregar_condicao_comercial(cfg)
    if "_tabela_condicao_tokens" in df_condicao.columns:
        df["desconto_correto_pct"] = _selecionar_desconto_por_tokens(df["_ean_norm"], tokens_esperados, df_condicao)
    else:
        df["desconto_correto_pct"] = float("nan")

    return df


def montar_saida_erro_bandeira(df: pd.DataFrame, cfg: dict, matriz_cfg: dict) -> pd.DataFrame:
    """tipo: "erro_bandeira" — monta a saída com os campos pedidos, nessa
    ordem exata: campos da base (sem o desconto aplicado), colunas trazidas
    do controle (id_bandeira, bandeira, perfil_bandeira, razao_social,
    cidade, estado) e a condição correta (tabela_correta,
    desconto_correto_pct).
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
    colunas_trazidas = list((matriz_cfg.get("colunas_trazidas") or {}).keys())
    colunas_condicao = ["tabela_correta", "desconto_correto_pct"]
    return df[colunas_base + colunas_trazidas + colunas_condicao].copy()


def _agregar_volume(df: pd.DataFrame, by) -> pd.DataFrame:
    """Agrupa por `by` somando qtd_pedidos e faturado_liquido, e também
    quantidade_faturada (unidades) quando a base tiver essa coluna
    configurada (base.colunas.quantidade_faturada) — não quebra quem não
    tiver essa coluna. Não reseta o índice, pra permitir reindex antes.
    """
    agregacoes = {"qtd_pedidos": ("_tabela_norm", "size"), "faturado_liquido": ("_faturado_liquido", "sum")}
    if "_quantidade_faturada" in df.columns:
        agregacoes["quantidade_faturada"] = ("_quantidade_faturada", "sum")
    return df.groupby(by).agg(**agregacoes)


def _completar_percentuais_e_medias(df: pd.DataFrame, sufixo_percentual: str = "") -> None:
    """Adiciona (in-place) percentual_faturado/percentual_quantidade_faturada
    (sobre o total do próprio df) e faturado_medio_por_pedido/
    quantidade_media_por_pedido — reaproveitado por todas as tabelas de
    resumo_volume/resumo_por_cnpj.
    """
    total_pedidos = df["qtd_pedidos"].sum()
    total_faturado = df["faturado_liquido"].sum()
    df[f"percentual_pedidos{sufixo_percentual}"] = (
        (df["qtd_pedidos"] / total_pedidos * 100).round(2) if total_pedidos else 0.0
    )
    df[f"percentual_faturado{sufixo_percentual}"] = (
        (df["faturado_liquido"] / total_faturado * 100).round(2) if total_faturado else 0.0
    )
    df["faturado_medio_por_pedido"] = (df["faturado_liquido"] / df["qtd_pedidos"]).where(df["qtd_pedidos"] > 0).round(
        2
    )
    if "quantidade_faturada" in df.columns:
        total_quantidade = df["quantidade_faturada"].sum()
        df[f"percentual_quantidade_faturada{sufixo_percentual}"] = (
            (df["quantidade_faturada"] / total_quantidade * 100).round(2) if total_quantidade else 0.0
        )
        df["quantidade_media_por_pedido"] = (
            (df["quantidade_faturada"] / df["qtd_pedidos"]).where(df["qtd_pedidos"] > 0).round(2)
        )


def calcular_resumo_volume(df_base: pd.DataFrame, matriz_cfg: dict) -> pd.DataFrame:
    """tipo: "resumo_volume" — roda na base inteira (sem filtro por CNPJ,
    sem Check): classifica cada pedido em categoria_a (Tabela de negociação
    contém alguma das palavras_chave_categoria_a) ou categoria_b (o resto),
    e soma quantidade de pedidos, faturado líquido e quantidade faturada
    (se configurada) por categoria, com o percentual de cada uma sobre o
    total e a média por pedido individual (volume por pedido).
    """
    palavras = [normalize_text(p) for p in matriz_cfg["palavras_chave_categoria_a"]]
    nome_a = matriz_cfg.get("nome_categoria_a", "A")
    nome_b = matriz_cfg.get("nome_categoria_b", "B")

    categoria = df_base["_tabela_norm"].map(lambda t: nome_a if any(p in t for p in palavras) else nome_b)

    resumo = (
        _agregar_volume(df_base.assign(categoria=categoria), "categoria")
        .reindex([nome_a, nome_b], fill_value=0)
        .reset_index()
    )
    _completar_percentuais_e_medias(resumo)
    return resumo


def calcular_resumo_mensal(df_base: pd.DataFrame, matriz_cfg: dict) -> dict[str, pd.DataFrame]:
    """Quebra mensal do "resumo_volume": CA x WE por mês (com percentual
    dentro de cada mês) + média mensal por categoria; dentro de CA, o
    volume total (todos os meses) de cada tabela específica isolada (ex.:
    só "Canal Autorizador"), e também mês a mês, com qual delas puxou mais
    volume (faturado líquido) em cada mês.
    """
    palavras = list(matriz_cfg["palavras_chave_categoria_a"])
    palavras_norm = [normalize_text(p) for p in palavras]
    nome_a = matriz_cfg.get("nome_categoria_a", "A")
    nome_b = matriz_cfg.get("nome_categoria_b", "B")

    def _tabela_especifica(tabela_norm: str) -> str | None:
        for original, norm in zip(palavras, palavras_norm):
            if norm in tabela_norm:
                return original
        return None

    df = df_base.copy()
    df["categoria"] = df["_tabela_norm"].map(lambda t: nome_a if any(p in t for p in palavras_norm) else nome_b)
    df["mes"] = df["_data_pedido"].dt.to_period("M").astype(str)

    mensal = _agregar_volume(df, ["mes", "categoria"]).reset_index()
    total_mes = mensal.groupby("mes")[[c for c in mensal.columns if c not in ("mes", "categoria")]].transform("sum")
    mensal["percentual_pedidos"] = (mensal["qtd_pedidos"] / total_mes["qtd_pedidos"] * 100).round(2)
    mensal["percentual_faturado"] = (mensal["faturado_liquido"] / total_mes["faturado_liquido"] * 100).round(2)
    # volume por pedido = faturado líquido médio de cada pedido individual naquele mês/categoria
    mensal["faturado_medio_por_pedido"] = (mensal["faturado_liquido"] / mensal["qtd_pedidos"]).round(2)
    colunas_media_mensal = ["qtd_pedidos", "faturado_liquido", "percentual_pedidos", "percentual_faturado", "faturado_medio_por_pedido"]
    if "quantidade_faturada" in mensal.columns:
        mensal["percentual_quantidade_faturada"] = (
            mensal["quantidade_faturada"] / total_mes["quantidade_faturada"] * 100
        ).round(2)
        mensal["quantidade_media_por_pedido"] = (mensal["quantidade_faturada"] / mensal["qtd_pedidos"]).round(2)
        colunas_media_mensal += ["quantidade_faturada", "percentual_quantidade_faturada", "quantidade_media_por_pedido"]

    media_mensal = mensal.groupby("categoria")[colunas_media_mensal].mean().round(2).reset_index()

    df_ca = df[df["categoria"] == nome_a].copy()
    df_ca["tabela_especifica"] = df_ca["_tabela_norm"].map(_tabela_especifica)

    # total (todos os meses juntos) por tabela específica dentro de CA —
    # ex.: quantos pedidos são só de "Canal Autorizador", isolado das outras
    # tabelas que também compõem a categoria CA
    ca_por_tabela = _agregar_volume(df_ca, "tabela_especifica").reset_index()
    _completar_percentuais_e_medias(ca_por_tabela, sufixo_percentual="_dentro_ca")

    ca_por_tabela_mes = _agregar_volume(df_ca, ["mes", "tabela_especifica"]).reset_index()
    ca_por_tabela_mes["faturado_medio_por_pedido"] = (
        ca_por_tabela_mes["faturado_liquido"] / ca_por_tabela_mes["qtd_pedidos"]
    ).round(2)
    if "quantidade_faturada" in ca_por_tabela_mes.columns:
        ca_por_tabela_mes["quantidade_media_por_pedido"] = (
            ca_por_tabela_mes["quantidade_faturada"] / ca_por_tabela_mes["qtd_pedidos"]
        ).round(2)

    renomear_maior = {
        "tabela_especifica": "tabela_maior_volume",
        "qtd_pedidos": "qtd_pedidos_maior",
        "faturado_liquido": "faturado_liquido_maior",
        "faturado_medio_por_pedido": "faturado_medio_por_pedido_maior",
        "quantidade_faturada": "quantidade_faturada_maior",
        "quantidade_media_por_pedido": "quantidade_media_por_pedido_maior",
    }
    maior_por_mes = (
        ca_por_tabela_mes.sort_values("faturado_liquido", ascending=False)
        .groupby("mes", as_index=False)
        .first()
        .rename(columns=renomear_maior)
        .sort_values("mes")
        .reset_index(drop=True)
    )

    return {
        "Mensal_CAxWE": mensal,
        "Media_Mensal": media_mensal,
        "CA_por_Tabela": ca_por_tabela,
        "CA_por_Tabela_Mes": ca_por_tabela_mes,
        "Maior_Volume_por_Mes": maior_por_mes,
    }


def _preparar_df_por_grupo(df_base: pd.DataFrame, matriz_cfg: dict) -> tuple[pd.DataFrame, str, str]:
    """Filtra a base pelas palavras_chave (ex.: "RAIA CA", "CANAL
    AUTORIZADOR") e monta a coluna "tabela_especifica" e "mes" — reaproveitado
    por calcular_resumo_por_cnpj e sua versão mensal.

    Por padrão a chave de agrupamento é o CNPJ; se matriz_cfg tiver
    coluna_agrupamento: "distribuidor", usa a coluna base.colunas.distribuidor
    (ex.: "Nome do distribuidor") em vez do CNPJ. Retorna (df filtrado,
    coluna interna a agrupar, nome da coluna na saída).
    """
    palavras = list(matriz_cfg["palavras_chave"])
    palavras_norm = [normalize_text(p) for p in palavras]

    def _tabela_especifica(tabela_norm: str) -> str | None:
        for original, norm in zip(palavras, palavras_norm):
            if norm in tabela_norm:
                return original
        return None

    coluna_agrupamento = matriz_cfg.get("coluna_agrupamento", "cnpj")
    if coluna_agrupamento == "distribuidor":
        chave_col, nome_saida = "_distribuidor_norm", "distribuidor"
    else:
        chave_col, nome_saida = "_cnpj_norm", "cnpj"

    df = df_base.copy()
    df["tabela_especifica"] = df["_tabela_norm"].map(_tabela_especifica)
    df = df[df["tabela_especifica"].notna()]
    df["mes"] = df["_data_pedido"].dt.to_period("M").astype(str)
    return df, chave_col, nome_saida


def calcular_resumo_por_cnpj(df_base: pd.DataFrame, matriz_cfg: dict) -> pd.DataFrame:
    """tipo: "resumo_por_cnpj" — roda na base inteira (sem filtro por CNPJ
    de controle, sem Check): filtra só as tabelas de palavras_chave (ex.:
    "RAIA CA", "DPSP CA", ou "CANAL AUTORIZADOR") e acumula (total, sem
    quebra mensal) quantidade de pedidos, faturado líquido e quantidade
    faturada (se configurada), dentro de cada tabela — agrupado por CNPJ ou
    por distribuidor (ver _preparar_df_por_grupo).
    """
    df, chave_col, nome_saida = _preparar_df_por_grupo(df_base, matriz_cfg)

    resumo = _agregar_volume(df, ["tabela_especifica", chave_col]).reset_index().rename(columns={chave_col: nome_saida})
    resumo["faturado_medio_por_pedido"] = (resumo["faturado_liquido"] / resumo["qtd_pedidos"]).round(2)
    if "quantidade_faturada" in resumo.columns:
        resumo["quantidade_media_por_pedido"] = (resumo["quantidade_faturada"] / resumo["qtd_pedidos"]).round(2)
    resumo = resumo.sort_values(["tabela_especifica", "faturado_liquido"], ascending=[True, False]).reset_index(
        drop=True
    )
    return resumo


def calcular_resumo_por_cnpj_mensal(df_base: pd.DataFrame, matriz_cfg: dict) -> pd.DataFrame:
    """Versão mensal de calcular_resumo_por_cnpj: mesma filtragem e
    agrupamento (CNPJ ou distribuidor), mas abrindo por mês em vez de
    acumular tudo — usada quando a matriz tem abrir_por_mes: true.
    """
    df, chave_col, nome_saida = _preparar_df_por_grupo(df_base, matriz_cfg)

    mensal = (
        _agregar_volume(df, ["tabela_especifica", chave_col, "mes"]).reset_index().rename(columns={chave_col: nome_saida})
    )
    mensal["faturado_medio_por_pedido"] = (mensal["faturado_liquido"] / mensal["qtd_pedidos"]).round(2)
    if "quantidade_faturada" in mensal.columns:
        mensal["quantidade_media_por_pedido"] = (mensal["quantidade_faturada"] / mensal["qtd_pedidos"]).round(2)
    mensal = mensal.sort_values(["tabela_especifica", nome_saida, "mes"]).reset_index(drop=True)
    return mensal


def calcular_resumo_cnpj_grupo(df_base: pd.DataFrame, cfg: dict, matriz_cfg: dict) -> pd.DataFrame:
    """tipo: "resumo_cnpj_grupo" — filtra exclusivamente pela palavra_chave
    da matriz (ex.: "CANAL AUTORIZADOR") e agrupa por CNPJ + Grupo de
    clientes + mês, somando quantidade de pedidos, faturado líquido e
    quantidade faturada (se configurada) de cada mês. Exige
    base.colunas.grupo_clientes configurado.
    """
    colunas = cfg["base"]["colunas"]
    if "grupo_clientes" not in colunas:
        raise KeyError(
            "Configure base.colunas.grupo_clientes no config.yaml pra usar a matriz tipo 'resumo_cnpj_grupo'."
        )

    palavra = normalize_text(matriz_cfg["palavra_chave"])
    df = df_base[df_base["_tabela_norm"].str.contains(palavra, na=False)].copy()
    df["grupo_clientes"] = df[colunas["grupo_clientes"]]
    df["mes"] = df["_data_pedido"].dt.to_period("M").astype(str)

    resumo = (
        _agregar_volume(df, ["_cnpj_norm", "grupo_clientes", "mes"]).reset_index().rename(columns={"_cnpj_norm": "cnpj"})
    )
    return resumo.sort_values(["cnpj", "mes"]).reset_index(drop=True)


def _salvar_saida(df_saida: pd.DataFrame, pasta_saida: Path, nome_arquivo: str) -> None:
    caminho_final = pasta_saida / nome_arquivo
    df_saida.to_excel(caminho_final, index=False)
    print(f"Resultado salvo em {caminho_final}")

    if "diferenca_faturamento" in df_saida.columns:
        impacto = df_saida["diferenca_faturamento"].sum(skipna=True)
        print(f"Impacto financeiro total dos erros operacionais: R$ {impacto:,.2f}")


def rodar_matriz(nome_matriz: str, matriz_cfg: dict, df_base: pd.DataFrame, cfg: dict) -> pd.DataFrame | None:
    print(f"\n=== Matriz: {nome_matriz} ===")
    tipo = matriz_cfg.get("tipo", "tabela")

    if tipo == "resumo_volume":
        # roda na base inteira, sem arquivo de controle nem CNPJ — não passa
        # pelo carregar_controle/filtro das outras matrizes
        resumo = calcular_resumo_volume(df_base, matriz_cfg)
        abas_mensais = calcular_resumo_mensal(df_base, matriz_cfg)

        print(resumo.to_string(index=False))
        print("\nMédia mensal por categoria:")
        print(abas_mensais["Media_Mensal"].to_string(index=False))
        print("\nTabela com maior volume por mês (dentro de CA):")
        print(abas_mensais["Maior_Volume_por_Mes"].to_string(index=False))

        pasta_saida = BASE_DIR / cfg["saida"]["pasta"]
        pasta_saida.mkdir(parents=True, exist_ok=True)
        nome_arquivo = matriz_cfg.get("nome_arquivo_saida") or f"{nome_matriz}_analise.xlsx"
        caminho_final = pasta_saida / nome_arquivo
        with pd.ExcelWriter(caminho_final) as writer:
            resumo.to_excel(writer, sheet_name="Resumo", index=False)
            for nome_aba, df_aba in abas_mensais.items():
                df_aba.to_excel(writer, sheet_name=nome_aba, index=False)
        print(f"Resultado salvo em {caminho_final}")
        return resumo

    if tipo == "resumo_por_cnpj":
        # roda na base inteira, sem arquivo de controle nem CNPJ de entrada
        resumo = calcular_resumo_por_cnpj(df_base, matriz_cfg)
        if resumo.empty:
            print("Nenhuma venda encontrada para as palavras-chave configuradas.")
            return None
        print(resumo.to_string(index=False))
        pasta_saida = BASE_DIR / cfg["saida"]["pasta"]
        pasta_saida.mkdir(parents=True, exist_ok=True)
        nome_arquivo = matriz_cfg.get("nome_arquivo_saida") or f"{nome_matriz}_analise.xlsx"

        if matriz_cfg.get("abrir_por_mes"):
            mensal = calcular_resumo_por_cnpj_mensal(df_base, matriz_cfg)
            caminho_final = pasta_saida / nome_arquivo
            with pd.ExcelWriter(caminho_final) as writer:
                resumo.to_excel(writer, sheet_name="Total", index=False)
                mensal.to_excel(writer, sheet_name="Mensal", index=False)
            print(f"Resultado salvo em {caminho_final}")
        else:
            _salvar_saida(resumo, pasta_saida, nome_arquivo)
        return resumo

    if tipo == "resumo_cnpj_grupo":
        # filtra exclusivamente pela palavra_chave da matriz (ex.: "CANAL
        # AUTORIZADOR") — não roda na base inteira
        resumo = calcular_resumo_cnpj_grupo(df_base, cfg, matriz_cfg)
        if resumo.empty:
            print(f"Nenhuma linha encontrada para a palavra-chave '{matriz_cfg['palavra_chave']}'.")
            return None
        print(f"{len(resumo)} linhas geradas (CNPJ x grupo de clientes x mês).")
        pasta_saida = BASE_DIR / cfg["saida"]["pasta"]
        pasta_saida.mkdir(parents=True, exist_ok=True)
        nome_arquivo = matriz_cfg.get("nome_arquivo_saida") or f"{nome_matriz}_analise.xlsx"
        _salvar_saida(resumo, pasta_saida, nome_arquivo)
        return resumo

    if tipo == "erro_bandeira":
        # filtra pelos CNPJs do controle (ex.: Painel_Bandeira) e compara
        # com o Painel x Tabela — só as linhas erradas voltam
        df_erros = calcular_erro_bandeira(df_base, cfg, matriz_cfg)
        if df_erros.empty:
            print("Nenhum pedido em Tabela errada — todos batem com o Painel x Tabela.")
            return None
        df_saida = montar_saida_erro_bandeira(df_erros, cfg, matriz_cfg)
        col_grupo = cfg["base"]["colunas"]["grupo_clientes"]
        col_data = cfg["base"]["colunas"]["data_pedido"]
        df_saida = df_saida.sort_values([col_grupo, col_data]).reset_index(drop=True)
        print(f"{len(df_saida)} pedidos em Tabela errada.")
        pasta_saida = BASE_DIR / cfg["saida"]["pasta"]
        pasta_saida.mkdir(parents=True, exist_ok=True)
        nome_arquivo = matriz_cfg.get("nome_arquivo_saida") or "Análise Erro Bandeira.xlsx"
        _salvar_saida(df_saida, pasta_saida, nome_arquivo)
        return df_saida

    df_controle = carregar_controle(matriz_cfg)

    if tipo in ("cnpj", "consolidacao"):
        cnpjs_validos = set(df_controle["_chave_controle_norm"])
        df_filtrado = filtrar_por_cnpj(df_base, cnpjs_validos)
        if df_filtrado.empty:
            print("Nenhuma venda encontrada para os CNPJs do arquivo de controle.")
            return None
        print(f"{len(df_filtrado)} linhas filtradas pelos {len(cnpjs_validos)} CNPJs do controle.")
    else:
        df_filtrado = filtrar_matriz(df_base, matriz_cfg["palavra_chave"])
        if df_filtrado.empty:
            print(f"Nenhuma linha encontrada para a palavra-chave '{matriz_cfg['palavra_chave']}'.")
            return None
        print(f"{len(df_filtrado)} linhas filtradas em 'Tabela de negociação'.")

    df_merge = aplicar_procx_cnpj(df_filtrado, df_controle)
    pasta_saida = BASE_DIR / cfg["saida"]["pasta"]
    pasta_saida.mkdir(parents=True, exist_ok=True)

    if tipo == "consolidacao":
        # sem Check, sem cruzamento de desconto — só consolida as vendas
        # encontradas com as colunas trazidas do controle (ex.: Bandeira)
        df_saida = montar_saida_consolidacao(df_merge, cfg, matriz_cfg)
        nome_arquivo = matriz_cfg.get("nome_arquivo_saida") or f"{nome_matriz}_analise.xlsx"
        _salvar_saida(df_saida, pasta_saida, nome_arquivo)
        return df_saida

    df_merge["Check"] = calcular_check(df_merge, matriz_cfg)
    print(df_merge["Check"].value_counts(dropna=False).to_string())

    df_condicao = carregar_condicao_comercial(cfg)
    df_final = aplicar_condicao_correta(df_merge, df_condicao, matriz_cfg)

    df_saida = montar_saida(df_final, cfg, matriz_cfg)
    nome_arquivo = matriz_cfg.get("nome_arquivo_saida") or f"{nome_matriz}_analise.xlsx"
    _salvar_saida(df_saida, pasta_saida, nome_arquivo)
    return df_saida


def perguntar_quais_matrizes(ativas: list[dict]) -> list[dict]:
    """Menu interativo (console) pra escolher qual análise rodar, sem perder
    as outras já configuradas. Só é chamado quando dá pra interagir (terminal
    de verdade) e nem --matrizes nem --todas foram usados.
    """
    print("\nQual análise você quer rodar?")
    print("  0 - Todas as análises ativas")
    for i, m in enumerate(ativas, start=1):
        print(f"  {i} - {m['nome']} (tipo: {m.get('tipo', 'tabela')})")

    escolha = input("Digite o(s) número(s) separados por vírgula (Enter = todas): ").strip()
    if not escolha or escolha == "0":
        return ativas

    indices_validos = set()
    for pedaco in escolha.split(","):
        pedaco = pedaco.strip()
        if pedaco.isdigit() and 1 <= int(pedaco) <= len(ativas):
            indices_validos.add(int(pedaco))
        else:
            print(f"Ignorando opção inválida: '{pedaco}'")

    return [m for i, m in enumerate(ativas, start=1) if i in indices_validos]


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--config", default=str(BASE_DIR / "config.yaml"), help="Caminho do config.yaml")
    parser.add_argument(
        "--matrizes",
        help="Lista separada por vírgula das matrizes a rodar (ignora o 'ativo' do config e o menu). Ex: Feira,CanalAutorizador",
    )
    parser.add_argument(
        "--todas", action="store_true", help="Roda todas as matrizes ativas sem mostrar o menu interativo"
    )
    parser.add_argument("--listar", action="store_true", help="Lista as matrizes configuradas e sai")
    args = parser.parse_args()

    cfg = carregar_config(args.config)

    if args.listar:
        for m in cfg["matrizes"]:
            palavra_chave = m.get("palavra_chave")
            detalhe = f", palavra_chave='{palavra_chave}'" if palavra_chave else ""
            print(f"- {m['nome']} (tipo={m.get('tipo', 'tabela')}, ativo={m['ativo']}{detalhe})")
        return

    if args.matrizes:
        selecionadas = {nome.strip() for nome in args.matrizes.split(",")}
        matrizes_a_rodar = [m for m in cfg["matrizes"] if m["nome"] in selecionadas]
    else:
        ativas = [m for m in cfg["matrizes"] if m.get("ativo")]
        if not ativas:
            matrizes_a_rodar = []
        elif args.todas:
            matrizes_a_rodar = ativas
        else:
            try:
                matrizes_a_rodar = perguntar_quais_matrizes(ativas)
            except EOFError:
                # rodando sem console interativo (ex.: tarefa agendada) — não
                # trava esperando input, só roda tudo que está ativo
                print("Sem entrada interativa disponível — rodando todas as matrizes ativas.")
                matrizes_a_rodar = ativas

    if not matrizes_a_rodar:
        print("Nenhuma matriz selecionada. Marque 'ativo: true' no config.yaml, use --matrizes ou escolha no menu.")
        return

    df_base = carregar_base(cfg)
    print(f"Base carregada: {len(df_base)} linhas.")

    for matriz_cfg in matrizes_a_rodar:
        try:
            rodar_matriz(matriz_cfg["nome"], matriz_cfg, df_base, cfg)
        except Exception as e:
            # um erro numa matriz (ex.: coluna ausente por causa de outra
            # planilha) não pode impedir as demais de rodar e salvar seu
            # resultado — sem isso, uma matriz travada escondia até o
            # arquivo mais recente da matriz seguinte da lista.
            print(f"\n!!! Erro ao rodar a matriz '{matriz_cfg['nome']}': {e}")
            print("Seguindo para a próxima matriz...\n")


if __name__ == "__main__":
    main()
