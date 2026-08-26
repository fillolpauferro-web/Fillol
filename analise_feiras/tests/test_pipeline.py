import os
import sys
from pathlib import Path

import pandas as pd
import yaml

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from pipeline import (  # noqa: E402
    CHECK_ERRO,
    CHECK_OK,
    _bate_mapa_bandeira,
    _rotulo_bate_na_tabela,
    carregar_base,
    perguntar_quais_matrizes,
    rodar_matriz,
)
from utils import (  # noqa: E402
    normalize_cnpj,
    normalize_cnpj_raiz,
    normalize_ean,
    normalize_text,
    read_table_mais_recente,
    remover_prefixo_tabela_agregadora,
    to_datetime,
    to_numeric,
    tokenizar,
)


def test_remover_prefixo_tabela_agregadora():
    assert remover_prefixo_tabela_agregadora(normalize_text("Tabela Agregadora - Carrefour CA")) == "CARREFOUR CA"
    assert remover_prefixo_tabela_agregadora(normalize_text("Tabela Agregadora - Canal Autorizador")) == (
        "CANAL AUTORIZADOR"
    )
    # sem o prefixo, não mexe em nada
    assert remover_prefixo_tabela_agregadora(normalize_text("Carrefour CA")) == "CARREFOUR CA"


def test_rotulo_bate_na_tabela_ignora_ordem_e_aceita_abreviacao():
    # caso real 1: mesmas palavras, ordem trocada, com prefixo extra
    rot = tokenizar(normalize_text("GENERICO_D1000"))
    tab = tokenizar(normalize_text("Tabela Agregadora - D1000_GENERICO"))
    assert _rotulo_bate_na_tabela(rot, tab)

    # caso real 2: abreviação (AUT é prefixo de AUTORIZADOR)
    rot = tokenizar(normalize_text("CANAL_AUT"))
    tab = tokenizar(normalize_text("Canal Autorizador"))
    assert _rotulo_bate_na_tabela(rot, tab)

    # não pode bater com uma tabela genérica que não tem a palavra específica
    rot = tokenizar(normalize_text("GENERICO_D1000"))
    tab = tokenizar(normalize_text("DEFAULT GENERICO CA"))
    assert not _rotulo_bate_na_tabela(rot, tab)

    # rótulo vazio nunca bate
    assert not _rotulo_bate_na_tabela((), tab)


def test_bate_mapa_bandeira_cobre_nomes_sem_relacao_textual():
    mapa = [
        {"bandeira_contem": "DROGAO SUPER", "tabela_contem": "DROGAO SUPER"},
        {"bandeira_contem": "PACHECO", "tabela_contem": "DPSP"},
    ]

    # casos reais confirmados pelo usuário
    tab = tokenizar(normalize_text("Tabela Agregadora - DROGAO SUPER_GENERICO"))
    assert _bate_mapa_bandeira(normalize_text("DROGAO SUPER SP"), tab, mapa)

    # "Pacheco" e "DPSP_CA" não têm nenhuma palavra em comum — só o mapa
    # manual resolve, o match por Rotulo/token não teria como
    tab = tokenizar(normalize_text("DPSP_CA"))
    assert _bate_mapa_bandeira(normalize_text("PACHECO"), tab, mapa)

    # bandeira sem item no mapa não bate
    assert not _bate_mapa_bandeira(normalize_text("OUTRA REDE"), tab, mapa)

    # mapa vazio nunca bate
    assert not _bate_mapa_bandeira(normalize_text("PACHECO"), tab, [])


def test_carregar_base_ignora_prefixo_tabela_agregadora(tmp_path: Path):
    # pedidos lançados na versão "Tabela Agregadora - X" são tão válidos
    # quanto na versão direta "X" — carregar_base já precisa tratar como
    # equivalentes pra Feira/CanalAutorizador/regra da Bandeira funcionarem
    # sem precisar saber qual das duas formas a base usou.
    base_df = pd.DataFrame(
        {
            "Tabela de negociação": ["Tabela Agregadora - Carrefour CA", "Carrefour CA"],
            "CNPJ": ["11.111.111/0001-11", "22.222.222/0001-22"],
            "EAN": ["1111111111111", "2222222222222"],
            "Id pedido": ["1", "2"],
            "Data do pedido (original)": ["10/05/2026 10:00", "10/05/2026 10:00"],
            "Faturado líquido (R$)": ["27,3", "27,3"],
            "Desconto comercial faturado (%)": ["20", "20"],
        }
    )
    base_df.to_excel(tmp_path / "base_pedidos.xlsx", index=False)

    cfg = {
        "base": {
            "arquivo": "base_pedidos.xlsx",
            "aba": None,
            "colunas": {
                "tabela_negociacao": "Tabela de negociação",
                "cnpj": "CNPJ",
                "ean": "EAN",
                "id_pedido": "Id pedido",
                "data_pedido": "Data do pedido (original)",
                "faturado_liquido": "Faturado líquido (R$)",
                "desconto_aplicado_pct": "Desconto comercial faturado (%)",
            },
        }
    }

    import pipeline

    pipeline.BASE_DIR = tmp_path
    df_base = carregar_base(cfg)

    assert df_base["_tabela_norm"].tolist() == ["CARREFOUR CA", "CARREFOUR CA"]
    # a coluna original de saída continua com o texto de verdade
    assert df_base["Tabela de negociação"].tolist() == ["Tabela Agregadora - Carrefour CA", "Carrefour CA"]


def test_to_numeric_aceita_formato_brasileiro_e_internacional():
    casos = pd.Series(
        [
            "56,87",  # pt-BR simples
            "1.234,56",  # pt-BR com milhar
            "0.3",  # float nativo do Excel convertido para string (decimal com ponto)
            "1234.56",  # idem, sem milhar
            "300",  # inteiro puro
            "'42,5",  # aspas de "número como texto" do Excel
        ]
    )
    resultado = to_numeric(casos).tolist()
    assert resultado == [56.87, 1234.56, 0.3, 1234.56, 300.0, 42.5]


def test_to_datetime_aceita_data_muito_no_futuro():
    # "31/12/9999" (ou similar) é usado como marcador de "sem prazo definido"
    # em algumas planilhas de controle; datetime64[ns] estoura em ~2262, então
    # isso não pode quebrar o parsing das outras datas da mesma coluna.
    casos = pd.Series(["05/05/2026", "31/12/9999", "2026-05-10 09:00:00", None])
    resultado = to_datetime(casos)
    assert resultado.iloc[0] == pd.Timestamp("2026-05-05")
    assert resultado.iloc[1] == pd.Timestamp("9999-12-31")
    assert resultado.iloc[2] == pd.Timestamp("2026-05-10 09:00:00")
    assert pd.isna(resultado.iloc[3])


def test_normalize_cnpj_preenche_com_zero_a_esquerda():
    # Excel guardado como número perde o zero à esquerda do CNPJ (13
    # dígitos em vez de 14) — TEXTO(CNPJ;"00000000000000") corrige isso.
    assert normalize_cnpj("1672858000165") == "01672858000165"
    assert normalize_cnpj("01.672.858/0001-65") == "01672858000165"
    assert normalize_cnpj(None) == ""


def test_normalize_cnpj_remove_decimal_zero_residual():
    # célula numérica sem casas decimais vira string terminando em ".0"
    # (ex.: pandas lendo uma coluna float) — não pode virar um dígito extra
    assert normalize_cnpj("1672858000165.0") == "01672858000165"
    assert normalize_cnpj_raiz("42225938.0") == "42225938"
    assert normalize_cnpj_raiz("9156879.0") == "09156879"
    # CNPJ mascarado não pode ser afetado (não termina em ".0")
    assert normalize_cnpj("01.672.858/0001-65") == "01672858000165"


def test_normalize_ean_remove_aspas_das_duas_pontas():
    assert normalize_ean("'7896422511865'") == "7896422511865"
    assert normalize_ean("'7896422511865") == "7896422511865"
    assert normalize_ean(" 7896422511865 ") == "7896422511865"
    assert normalize_ean("7896422511865.0") == "7896422511865"
    assert normalize_ean(None) == ""


def test_read_table_mais_recente_pega_o_arquivo_mais_novo(tmp_path: Path):
    # nome do arquivo muda a cada exportação (ex.: Painel_NV_XXXX) — só o
    # mais recente deve valer, não a concatenação de todos
    antigo = tmp_path / "Painel_NV_2026_01.xlsx"
    novo = tmp_path / "Painel_NV_2026_02.xlsx"
    pd.DataFrame({"x": [1]}).to_excel(antigo, index=False)
    pd.DataFrame({"x": [2]}).to_excel(novo, index=False)
    os.utime(antigo, (1_000_000_000, 1_000_000_000))
    os.utime(novo, (2_000_000_000, 2_000_000_000))

    resultado = read_table_mais_recente(str(tmp_path / "Painel_NV_*.xlsx"))
    assert resultado["x"].iloc[0] == "2"


def test_perguntar_quais_matrizes(monkeypatch):
    ativas = [{"nome": "Feira", "tipo": "tabela"}, {"nome": "CanalAutorizador", "tipo": "cnpj"}]

    monkeypatch.setattr("builtins.input", lambda prompt="": "2")
    assert [m["nome"] for m in perguntar_quais_matrizes(ativas)] == ["CanalAutorizador"]

    monkeypatch.setattr("builtins.input", lambda prompt="": "")
    assert perguntar_quais_matrizes(ativas) == ativas

    monkeypatch.setattr("builtins.input", lambda prompt="": "1,2")
    assert len(perguntar_quais_matrizes(ativas)) == 2

    monkeypatch.setattr("builtins.input", lambda prompt="": "99,abc")
    assert perguntar_quais_matrizes(ativas) == []


def _montar_config(tmp_path: Path) -> dict:
    return {
        "base": {
            "arquivo": "base_pedidos.xlsx",
            "aba": None,
            "colunas": {
                "tabela_negociacao": "Tabela de negociação",
                "cnpj": "CNPJ",
                "ean": "EAN",
                "id_pedido": "Id pedido",
                "data_pedido": "Data do pedido (original)",
                "faturado_liquido": "Faturado líquido (R$)",
                "desconto_aplicado_pct": "Desconto comercial faturado (%)",
            },
        },
        "condicao_comercial": {
            "arquivo": "condicao_comercial.xlsx",
            "aba": "Dados",
            "colunas": {
                "chave_ean": "EAN FORMATADO",
                "desconto_correto_pct": "Desconto Atual",
            },
            "desconto_em_fracao": True,
        },
        "matrizes": [
            {
                "nome": "Feira",
                "ativo": True,
                "palavra_chave": "FEIRA",
                "arquivo_controle": "Controle_Feiras.xlsx",
                "aba_controle": "dados",
                "chave_controle": "CNPJ AJUSTADO",
                "colunas_trazidas": {
                    "inicio_real": "Início Real",
                    "termino_real": "Término Real",
                },
                "colunas_data": ["inicio_real", "termino_real"],
            }
        ],
        "saida": {"pasta": "saida"},
    }


def test_pipeline_end_to_end(tmp_path: Path):
    # pedido 1: CNPJ cadastrado e dentro do período -> OK
    # pedido 2: CNPJ não cadastrado no controle -> Erro Operacional,
    #           desconto correto vem de condicao_comercial por EAN;
    #           CNPJ e EAN vêm com aspas simples nas duas pontas, como na
    #           exportação real, pra testar a normalização
    # pedido 3: CNPJ cadastrado, mas pedido feito FORA do período
    #           inicio_real/termino_real -> Erro Operacional mesmo cadastrado
    # pedido 4: fora da matriz Feira (Tabela de negociação não é feira)
    # pedido 5: CNPJ não cadastrado (Erro Operacional) mas desconto comercial
    #           faturado = 0 -> mantém Erro Operacional, sem calcular preço/desconto
    base_df = pd.DataFrame(
        {
            "Tipo de cliente": [
                "ASSOCIATIVISMO",
                "ASSOCIATIVISMO",
                "ASSOCIATIVISMO",
                "REDE INDIRETA",
                "ASSOCIATIVISMO",
            ],
            "CNPJ": [
                "01.672.858/0001-65",
                "'06052566000143'",
                "01.672.858/0001-65",
                "06.052.566/0001-43",
                "06.052.566/0001-43",
            ],
            "Id pedido": ["17275", "17300", "17400", "17778", "17500"],
            "EAN": ["7896422511865", "'7896422511865'", "7896422511865", "7896422514651", "7896422511865"],
            "Tabela de negociação": [
                "FEIRA NEGOCIOS CA",
                "FEIRA NEGOCIOS CA",
                "FEIRA NEGOCIOS CA",
                "DEFAULT GENERICO CA",
                "FEIRA NEGOCIOS CA",
            ],
            "Data do pedido (original)": [
                "10/05/2026 18:34",
                "10/05/2026 10:00",
                "15/07/2026 09:00",  # fora do período do controle (maio)
                "05/05/2026 19:16",
                "10/05/2026 10:00",
            ],
            "Faturado líquido (R$)": ["27,3", "27,3", "27,3", "53,42", "27,3"],
            "Desconto comercial faturado (%)": ["56,87", "56,87", "56,87", "69,36", "0"],
        }
    )

    controle_df = pd.DataFrame(
        {
            "CNPJ AJUSTADO": ["01.672.858/0001-65"],
            "Início Real": ["01/05/2026"],
            "Término Real": ["31/05/2026"],
        }
    )

    condicao_df = pd.DataFrame(
        {
            "EAN FORMATADO": ["7896422511865"],
            # armazenado como fração (célula formatada como % no Excel real)
            "Desconto Atual": [0.30],
        }
    )

    (tmp_path / "saida").mkdir()
    base_df.to_excel(tmp_path / "base_pedidos.xlsx", index=False)
    with pd.ExcelWriter(tmp_path / "Controle_Feiras.xlsx") as w:
        controle_df.to_excel(w, sheet_name="dados", index=False)
    with pd.ExcelWriter(tmp_path / "condicao_comercial.xlsx") as w:
        condicao_df.to_excel(w, sheet_name="Dados", index=False)

    cfg = _montar_config(tmp_path)
    import pipeline

    pipeline.BASE_DIR = tmp_path

    df_base = carregar_base(cfg)
    assert len(df_base) == 5

    resultado = rodar_matriz("Feira", cfg["matrizes"][0], df_base, cfg)
    assert resultado is not None
    assert len(resultado) == 4  # só as 4 linhas de FEIRA NEGOCIOS CA

    checks = dict(zip(resultado["Id pedido"], resultado["Check"]))
    assert checks["17275"] == CHECK_OK
    assert checks["17300"] == CHECK_ERRO
    assert checks["17400"] == CHECK_ERRO  # cadastrado, mas fora do período
    assert checks["17500"] == CHECK_ERRO  # desconto aplicado = 0

    linha_ok = resultado.loc[resultado["Id pedido"] == "17275"].iloc[0]
    assert pd.Timestamp(linha_ok["inicio_real"]) == pd.Timestamp("2026-05-01")
    assert pd.Timestamp(linha_ok["termino_real"]) == pd.Timestamp("2026-05-31")
    assert pd.isna(linha_ok["diferenca_faturamento"])  # sem erro operacional, sem cálculo

    linha_erro = resultado.loc[resultado["Id pedido"] == "17300"].iloc[0]
    preco_sem_desconto_esperado = 27.3 / (1 - 0.5687)
    preco_correto_esperado = preco_sem_desconto_esperado * (1 - 0.30)
    assert round(linha_erro["preco_sem_desconto"], 2) == round(preco_sem_desconto_esperado, 2)
    # confirma que o EAN com aspas nas duas pontas casou certo com a condição comercial
    assert round(linha_erro["preco_liquido_desconto_correto"], 2) == round(preco_correto_esperado, 2)

    linha_fora_periodo = resultado.loc[resultado["Id pedido"] == "17400"].iloc[0]
    # mesmo fora do período, as datas do controle continuam vindo na saída
    assert pd.Timestamp(linha_fora_periodo["inicio_real"]) == pd.Timestamp("2026-05-01")

    linha_desconto_zero = resultado.loc[resultado["Id pedido"] == "17500"].iloc[0]
    # Erro Operacional continua valendo, mas sem calcular preço/desconto
    assert linha_desconto_zero["Check"] == CHECK_ERRO
    assert pd.isna(linha_desconto_zero["preco_sem_desconto"])
    assert pd.isna(linha_desconto_zero["preco_liquido_desconto_correto"])
    assert pd.isna(linha_desconto_zero["diferenca_faturamento"])
    assert pd.isna(linha_desconto_zero["desconto_correto_pct"])

    assert (tmp_path / "saida" / "Feira_analise.xlsx").exists()


def test_matriz_tipo_cnpj_canal_autorizador(tmp_path: Path):
    # pedido 9001: CNPJ está no Painel_NV com Rótulo CANAL_AUT e a Tabela de
    #              negociação bate com "Canal Autorizador" -> OK
    # pedido 9002: mesmo CNPJ do Painel_NV, mas lançado numa Tabela diferente
    #              -> Erro Operacional; desconto correto tem que vir da linha
    #              "Canal Autorizador" da condicao_comercial (0.20), não da
    #              linha "Feira Negócios CA" (0.50) do mesmo EAN
    # pedido 9003: CNPJ está no Painel_NV, mas com outro Rótulo (não
    #              CANAL_AUT) -> nem entra na análise dessa matriz
    base_df = pd.DataFrame(
        {
            "Tipo de cliente": ["CANAL AUTORIZADO", "CANAL AUTORIZADO", "CANAL AUTORIZADO"],
            "CNPJ": ["11.222.333/0001-81", "11.222.333/0001-81", "44.555.666/0001-92"],
            "Id pedido": ["9001", "9002", "9003"],
            "EAN": ["1111111111111", "1111111111111", "1111111111111"],
            "Tabela de negociação": ["Canal Autorizador CA", "Associativismo CA", "Canal Autorizador CA"],
            "Data do pedido (original)": ["10/05/2026 10:00", "11/05/2026 10:00", "12/05/2026 10:00"],
            "Faturado líquido (R$)": ["27,3", "27,3", "27,3"],
            "Desconto comercial faturado (%)": ["56,87", "56,87", "56,87"],
        }
    )

    painel_nv_df = pd.DataFrame(
        {
            "CNPJ ajustado": ["11.222.333/0001-81", "44.555.666/0001-92"],
            "Rotulo": ["CANAL_AUT", "OUTRO_CANAL"],
        }
    )

    condicao_df = pd.DataFrame(
        {
            "EAN FORMATADO": ["1111111111111", "1111111111111"],
            "Tabela": ["Canal Autorizador", "Feira Negócios CA"],
            "Desconto Atual": [0.20, 0.50],
        }
    )

    (tmp_path / "saida").mkdir()
    base_df.to_excel(tmp_path / "base_pedidos.xlsx", index=False)
    with pd.ExcelWriter(tmp_path / "Painel_NV_2026_08.xlsx") as w:
        painel_nv_df.to_excel(w, sheet_name="Dados", index=False)
    with pd.ExcelWriter(tmp_path / "condicao_comercial.xlsx") as w:
        condicao_df.to_excel(w, sheet_name="Dados", index=False)

    cfg = _montar_config(tmp_path)
    cfg["condicao_comercial"]["colunas"]["chave_tabela"] = "Tabela"
    matriz_cfg = {
        "nome": "CanalAutorizador",
        "tipo": "cnpj",
        "ativo": True,
        "palavra_chave": "CANAL AUTORIZADOR",
        "arquivo_controle": "Painel_NV_*.xlsx",
        "aba_controle": "Dados",
        "chave_controle": "CNPJ ajustado",
        "coluna_rotulo_controle": "Rotulo",
        "rotulo_valido_controle": "CANAL_AUT",
        "colunas_trazidas": {},
        "colunas_data": [],
    }
    cfg["matrizes"].append(matriz_cfg)

    import pipeline

    pipeline.BASE_DIR = tmp_path

    df_base = carregar_base(cfg)
    resultado = rodar_matriz("CanalAutorizador", matriz_cfg, df_base, cfg)

    assert resultado is not None
    # pedido 9003 tem outro Rótulo no Painel_NV -> fora da análise
    assert set(resultado["Id pedido"]) == {"9001", "9002"}

    checks = dict(zip(resultado["Id pedido"], resultado["Check"]))
    assert checks["9001"] == CHECK_OK
    assert checks["9002"] == CHECK_ERRO

    linha_erro = resultado.loc[resultado["Id pedido"] == "9002"].iloc[0]
    preco_sem_desconto = 27.3 / (1 - 0.5687)
    preco_correto = preco_sem_desconto * (1 - 0.20)  # linha "Canal Autorizador", não a de Feira
    assert round(linha_erro["preco_liquido_desconto_correto"], 2) == round(preco_correto, 2)

    assert (tmp_path / "saida" / "CanalAutorizador_analise.xlsx").exists()


def test_matriz_tipo_consolidacao_bandeira(tmp_path: Path):
    # pedido 8001 e 8002: CNPJs que estão no Painel_Bandeira -> entram na
    #                     consolidação, cada um com a Bandeira do seu grupo
    # pedido 8003: CNPJ fora do Painel_Bandeira -> não entra
    base_df = pd.DataFrame(
        {
            "Tipo de cliente": ["ASSOCIATIVISMO", "ASSOCIATIVISMO", "ASSOCIATIVISMO"],
            "CNPJ": ["11.111.111/0001-11", "22.222.222/0001-22", "33.333.333/0001-33"],
            "Id pedido": ["8001", "8002", "8003"],
            "EAN": ["1111111111111", "2222222222222", "3333333333333"],
            "Tabela de negociação": ["DEFAULT GENERICO CA", "DEFAULT GENERICO CA", "DEFAULT GENERICO CA"],
            "Data do pedido (original)": ["10/05/2026 10:00", "11/05/2026 10:00", "12/05/2026 10:00"],
            "Faturado líquido (R$)": ["27,3", "40,0", "15,0"],
            "Desconto comercial faturado (%)": ["56,87", "20", "10"],
        }
    )

    painel_bandeira_df = pd.DataFrame(
        {
            "CNPJ Ajustado": ["11.111.111/0001-11", "22.222.222/0001-22"],
            "id_bandeira": ["10", "20"],
            "desc_bandeira": ["REDE A", "REDE B"],
            "perfil_bandeira": ["CALENDARIO", "CALENDARIO"],
            "razao_social": ["CLIENTE UM", "CLIENTE DOIS"],
            "cidade": ["SAO PAULO", "RIO DE JANEIRO"],
            "estado": ["SP", "RJ"],
        }
    )

    (tmp_path / "saida").mkdir()
    base_df.to_excel(tmp_path / "base_pedidos.xlsx", index=False)
    with pd.ExcelWriter(tmp_path / "Painel_Bandeira_2026_08.xlsx") as w:
        painel_bandeira_df.to_excel(w, sheet_name="Dados", index=False)

    cfg = _montar_config(tmp_path)
    matriz_cfg = {
        "nome": "Bandeira",
        "tipo": "consolidacao",
        "ativo": True,
        "arquivo_controle": "Painel_Bandeira_*.xlsx",
        "aba_controle": "Dados",
        "chave_controle": "CNPJ Ajustado",
        "colunas_trazidas": {
            "id_bandeira": "id_bandeira",
            "bandeira": "desc_bandeira",
            "perfil_bandeira": "perfil_bandeira",
            "razao_social": "razao_social",
            "cidade": "cidade",
            "estado": "estado",
        },
        "colunas_data": [],
        "nome_arquivo_saida": "historico_bandeiras.xlsx",
    }
    cfg["matrizes"].append(matriz_cfg)

    import pipeline

    pipeline.BASE_DIR = tmp_path

    df_base = carregar_base(cfg)
    resultado = rodar_matriz("Bandeira", matriz_cfg, df_base, cfg)

    assert resultado is not None
    # pedido 8003 tem CNPJ fora do Painel_Bandeira -> não entra
    assert set(resultado["Id pedido"]) == {"8001", "8002"}

    # sem Check nem colunas de desconto — é só consolidação
    assert "Check" not in resultado.columns
    assert "desconto_correto_pct" not in resultado.columns
    assert "diferenca_faturamento" not in resultado.columns

    bandeiras = dict(zip(resultado["Id pedido"], resultado["bandeira"]))
    assert bandeiras["8001"] == "REDE A"
    assert bandeiras["8002"] == "REDE B"

    razoes = dict(zip(resultado["Id pedido"], resultado["razao_social"]))
    assert razoes["8001"] == "CLIENTE UM"

    assert (tmp_path / "saida" / "historico_bandeiras.xlsx").exists()
    assert not (tmp_path / "saida" / "Bandeira_analise.xlsx").exists()


def test_matriz_bandeira_com_regra(tmp_path: Path):
    # pedido 9101: CNPJ 11.111.111/0001-11 (raiz 11111111), Tabela de
    #              negociação real "Tabela Agregadora - D1000_GENERICO"
    #              (mesmas palavras do Rotulo GENERICO_D1000, ordem trocada
    #              e com prefixo extra — caso real observado na base) -> OK
    # pedido 9102: CNPJ 22.222.222/0001-22 (raiz 22222222), Tabela NÃO bate
    #              com o Rotulo esperado (GENERICO_D500) -> Erro Operacional,
    #              com desconto correto calculado via EAN
    # pedido 9103: CNPJ 33.333.333/0001-33 (raiz 33333333) está no
    #              Painel_Bandeira mas NÃO tem linha no regra.xlsx -> tem
    #              que virar Erro Operacional sem quebrar (bug real: a
    #              coluna de tokens do Rotulo vira NaN pra CNPJ sem
    #              correspondência, e iterar sobre NaN estourava TypeError)
    # pedido 9104: CNPJ 44.444.444/0001-44 também SEM linha no regra.xlsx,
    #              mas a bandeira é "PACHECO" e a Tabela é "DPSP_CA" — sem
    #              nenhuma palavra em comum, só o mapa_bandeira_tabela
    #              (de-para manual) resgata pra OK
    base_df = pd.DataFrame(
        {
            "Tipo de cliente": ["ASSOCIATIVISMO", "ASSOCIATIVISMO", "ASSOCIATIVISMO", "ASSOCIATIVISMO"],
            "CNPJ": [
                "11.111.111/0001-11",
                "22.222.222/0001-22",
                "33.333.333/0001-33",
                "44.444.444/0001-44",
            ],
            "Id pedido": ["9101", "9102", "9103", "9104"],
            "EAN": ["1111111111111", "2222222222222", "3333333333333", "4444444444444"],
            "Tabela de negociação": [
                "Tabela Agregadora - D1000_GENERICO",
                "FEIRA NEGOCIOS CA",
                "QUALQUER OUTRA TABELA",
                "DPSP_CA",
            ],
            "Data do pedido (original)": [
                "10/05/2026 10:00",
                "11/05/2026 10:00",
                "12/05/2026 10:00",
                "13/05/2026 10:00",
            ],
            "Faturado líquido (R$)": ["27,3", "40,0", "15,0", "18,0"],
            "Desconto comercial faturado (%)": ["56,87", "20", "10", "30"],
        }
    )

    painel_bandeira_df = pd.DataFrame(
        {
            "CNPJ Ajustado": [
                "11.111.111/0001-11",
                "22.222.222/0001-22",
                "33.333.333/0001-33",
                "44.444.444/0001-44",
            ],
            "desc_bandeira": ["REDE A", "REDE B", "REDE C", "PACHECO"],
        }
    )

    regra_df = pd.DataFrame(
        {
            "Raiz CNPJ": ["11111111", "22222222"],
            "Rotulo": ["GENERICO_D1000", "GENERICO_D500"],
        }
    )

    condicao_df = pd.DataFrame(
        {
            "EAN FORMATADO": ["2222222222222"],
            "Desconto Atual": [0.25],
        }
    )

    (tmp_path / "saida").mkdir()
    base_df.to_excel(tmp_path / "base_pedidos.xlsx", index=False)
    with pd.ExcelWriter(tmp_path / "Painel_Bandeira_2026_08.xlsx") as w:
        painel_bandeira_df.to_excel(w, sheet_name="Dados", index=False)
    with pd.ExcelWriter(tmp_path / "regra.xlsx") as w:
        regra_df.to_excel(w, sheet_name="Dados", index=False)
    with pd.ExcelWriter(tmp_path / "condicao_comercial.xlsx") as w:
        condicao_df.to_excel(w, sheet_name="Dados", index=False)

    cfg = _montar_config(tmp_path)
    matriz_cfg = {
        "nome": "Bandeira",
        "tipo": "consolidacao",
        "ativo": True,
        "arquivo_controle": "Painel_Bandeira_*.xlsx",
        "aba_controle": "Dados",
        "chave_controle": "CNPJ Ajustado",
        "colunas_trazidas": {"bandeira": "desc_bandeira"},
        "colunas_data": [],
        "nome_arquivo_saida": "historico_bandeiras.xlsx",
        "regra": {
            "arquivo": "regra.xlsx",
            "aba": "Dados",
            "colunas": {"chave_raiz_cnpj": "Raiz CNPJ", "rotulo": "Rotulo"},
            "nome_arquivo_saida": "Bandeiras_Analise.xlsx",
            "mapa_bandeira_tabela": [
                {"bandeira_contem": "PACHECO", "tabela_contem": "DPSP"},
            ],
        },
    }
    cfg["matrizes"].append(matriz_cfg)

    import pipeline

    pipeline.BASE_DIR = tmp_path

    df_base = carregar_base(cfg)
    resultado = rodar_matriz("Bandeira", matriz_cfg, df_base, cfg)

    assert resultado is not None
    assert (tmp_path / "saida" / "historico_bandeiras.xlsx").exists()
    assert (tmp_path / "saida" / "Bandeiras_Analise.xlsx").exists()

    # o resultado retornado é o da segunda camada (com Check)
    assert "Check" in resultado.columns
    checks = dict(zip(resultado["Id pedido"], resultado["Check"]))
    assert checks["9101"] == CHECK_OK
    assert checks["9102"] == CHECK_ERRO
    assert checks["9103"] == CHECK_ERRO  # CNPJ sem linha no regra.xlsx
    assert checks["9104"] == CHECK_OK  # sem linha no regra.xlsx, mas resgatado pelo mapa_bandeira_tabela

    linha_erro = resultado.loc[resultado["Id pedido"] == "9102"].iloc[0]
    preco_sem_desconto = 40.0 / (1 - 0.20)
    preco_correto = preco_sem_desconto * (1 - 0.25)
    assert round(linha_erro["preco_liquido_desconto_correto"], 2) == round(preco_correto, 2)

    linha_ok = resultado.loc[resultado["Id pedido"] == "9101"].iloc[0]
    assert pd.isna(linha_ok["diferenca_faturamento"])


def test_matriz_tipo_resumo_volume(tmp_path: Path):
    # maio/2026: pedido 1 (Carrefour CA, 100), pedido 2 (Raia CA, 200) em CA;
    #            pedido 3 (Default Generico CA, 50) em WE.
    #            -> maior volume de CA no mês: Raia CA (200)
    # junho/2026: pedido 4 ("Tabela Agregadora - DPSP CA" -> DPSP CA, 150) e
    #            pedido 6 (Carrefour CA, 400) em CA; pedido 5
    #            (Araujo_Generico, 300) em WE.
    #            -> maior volume de CA no mês: Carrefour CA (400)
    # Roda na base inteira, sem filtro por CNPJ nem arquivo de controle.
    base_df = pd.DataFrame(
        {
            "Tipo de cliente": ["X", "X", "X", "X", "X", "X"],
            "CNPJ": [
                "11.111.111/0001-11",
                "22.222.222/0001-22",
                "33.333.333/0001-33",
                "44.444.444/0001-44",
                "55.555.555/0001-55",
                "66.666.666/0001-66",
            ],
            "Id pedido": ["1", "2", "3", "4", "5", "6"],
            "EAN": ["1", "2", "3", "4", "5", "6"],
            "Tabela de negociação": [
                "Carrefour CA",
                "Raia CA",
                "Default Generico CA",
                "Tabela Agregadora - DPSP CA",
                "Araujo_Generico",
                "Carrefour CA",
            ],
            "Data do pedido (original)": [
                "10/05/2026 10:00",
                "11/05/2026 10:00",
                "12/05/2026 10:00",
                "10/06/2026 10:00",
                "11/06/2026 10:00",
                "12/06/2026 10:00",
            ],
            "Faturado líquido (R$)": ["100,00", "200,00", "50,00", "150,00", "300,00", "400,00"],
            "Desconto comercial faturado (%)": ["10", "10", "10", "10", "10", "10"],
        }
    )

    base_df.to_excel(tmp_path / "base_pedidos.xlsx", index=False)
    (tmp_path / "saida").mkdir()

    cfg = _montar_config(tmp_path)
    matriz_cfg = {
        "nome": "ResumoCAxWE",
        "tipo": "resumo_volume",
        "ativo": True,
        "palavras_chave_categoria_a": [
            "CANAL AUTORIZADOR",
            "CARREFOUR CA",
            "DPSP CA",
            "PANVEL CA",
            "RAIA CA",
        ],
        "nome_categoria_a": "CA",
        "nome_categoria_b": "WE",
        "nome_arquivo_saida": "Resumo_CA_WE.xlsx",
    }

    import pipeline

    pipeline.BASE_DIR = tmp_path

    df_base = carregar_base(cfg)
    resultado = rodar_matriz("ResumoCAxWE", matriz_cfg, df_base, cfg)

    assert resultado is not None
    linhas = resultado.set_index("categoria")

    assert linhas.loc["CA", "qtd_pedidos"] == 4
    assert linhas.loc["WE", "qtd_pedidos"] == 2
    assert round(linhas.loc["CA", "faturado_liquido"], 2) == 850.0
    assert round(linhas.loc["WE", "faturado_liquido"], 2) == 350.0
    # volume por pedido = faturado líquido médio de cada pedido individual
    assert round(linhas.loc["CA", "faturado_medio_por_pedido"], 2) == 212.5
    assert round(linhas.loc["WE", "faturado_medio_por_pedido"], 2) == 175.0

    caminho_saida = tmp_path / "saida" / "Resumo_CA_WE.xlsx"
    assert caminho_saida.exists()

    abas = pd.read_excel(caminho_saida, sheet_name=None)
    assert set(abas.keys()) == {
        "Resumo",
        "Mensal_CAxWE",
        "Media_Mensal",
        "CA_por_Tabela_Mes",
        "Maior_Volume_por_Mes",
    }

    mensal = abas["Mensal_CAxWE"].set_index(["mes", "categoria"])
    assert mensal.loc[("2026-05", "CA"), "qtd_pedidos"] == 2
    assert round(mensal.loc[("2026-05", "CA"), "faturado_liquido"], 2) == 300.0
    assert round(mensal.loc[("2026-05", "CA"), "percentual_faturado"], 2) == 85.71
    assert round(mensal.loc[("2026-05", "CA"), "faturado_medio_por_pedido"], 2) == 150.0
    assert round(mensal.loc[("2026-06", "CA"), "faturado_liquido"], 2) == 550.0
    assert round(mensal.loc[("2026-06", "CA"), "faturado_medio_por_pedido"], 2) == 275.0

    media = abas["Media_Mensal"].set_index("categoria")
    assert round(media.loc["CA", "faturado_liquido"], 2) == 425.0
    assert round(media.loc["WE", "faturado_liquido"], 2) == 175.0
    assert round(media.loc["CA", "faturado_medio_por_pedido"], 2) == 212.5
    assert round(media.loc["WE", "faturado_medio_por_pedido"], 2) == 175.0

    maior = abas["Maior_Volume_por_Mes"].set_index("mes")
    assert maior.loc["2026-05", "tabela_maior_volume"] == "RAIA CA"
    assert round(maior.loc["2026-05", "faturado_liquido_maior"], 2) == 200.0
    assert round(maior.loc["2026-05", "faturado_medio_por_pedido_maior"], 2) == 200.0
    assert maior.loc["2026-06", "tabela_maior_volume"] == "CARREFOUR CA"
    assert round(maior.loc["2026-06", "faturado_liquido_maior"], 2) == 400.0
    assert round(maior.loc["2026-06", "faturado_medio_por_pedido_maior"], 2) == 400.0


def test_main_continua_apos_erro_em_uma_matriz(tmp_path: Path, monkeypatch):
    # Uma matriz mal configurada (arquivo de controle inexistente) não pode
    # travar o resto: a matriz seguinte da lista ainda precisa rodar e
    # salvar seu resultado.
    base_df = pd.DataFrame(
        {
            "Tipo de cliente": ["ASSOCIATIVISMO", "ASSOCIATIVISMO"],
            "CNPJ": ["11.111.111/0001-11", "22.222.222/0001-22"],
            "Id pedido": ["1", "2"],
            "EAN": ["1111111111111", "2222222222222"],
            "Tabela de negociação": ["FEIRA NEGOCIOS CA", "DEFAULT GENERICO CA"],
            "Data do pedido (original)": ["10/05/2026 10:00", "11/05/2026 10:00"],
            "Faturado líquido (R$)": ["27,3", "40,0"],
            "Desconto comercial faturado (%)": ["56,87", "20"],
        }
    )
    painel_bandeira_df = pd.DataFrame(
        {
            "CNPJ Ajustado": ["11.111.111/0001-11"],
            "desc_bandeira": ["REDE A"],
        }
    )

    (tmp_path / "saida").mkdir()
    base_df.to_excel(tmp_path / "base_pedidos.xlsx", index=False)
    with pd.ExcelWriter(tmp_path / "Painel_Bandeira_2026.xlsx") as w:
        painel_bandeira_df.to_excel(w, sheet_name="Dados", index=False)

    cfg = {
        "base": {
            "arquivo": "base_pedidos.xlsx",
            "aba": None,
            "colunas": {
                "tabela_negociacao": "Tabela de negociação",
                "cnpj": "CNPJ",
                "ean": "EAN",
                "id_pedido": "Id pedido",
                "data_pedido": "Data do pedido (original)",
                "faturado_liquido": "Faturado líquido (R$)",
                "desconto_aplicado_pct": "Desconto comercial faturado (%)",
            },
        },
        "condicao_comercial": {
            "arquivo": "condicao_comercial.xlsx",
            "aba": "Dados",
            "colunas": {"chave_ean": "EAN FORMATADO", "desconto_correto_pct": "Desconto Atual"},
            "desconto_em_fracao": True,
        },
        "matrizes": [
            {
                "nome": "Quebrada",
                "tipo": "tabela",
                "ativo": True,
                "palavra_chave": "FEIRA",
                "arquivo_controle": "nao_existe.xlsx",
                "aba_controle": None,
                "chave_controle": "CNPJ",
                "colunas_trazidas": {},
                "colunas_data": [],
            },
            {
                "nome": "Bandeira",
                "tipo": "consolidacao",
                "ativo": True,
                "arquivo_controle": "Painel_Bandeira_*.xlsx",
                "aba_controle": "Dados",
                "chave_controle": "CNPJ Ajustado",
                "colunas_trazidas": {"bandeira": "desc_bandeira"},
                "colunas_data": [],
                "nome_arquivo_saida": "historico_bandeiras.xlsx",
            },
        ],
        "saida": {"pasta": "saida"},
    }

    config_path = tmp_path / "config.yaml"
    config_path.write_text(yaml.dump(cfg), encoding="utf-8")

    import pipeline

    pipeline.BASE_DIR = tmp_path
    monkeypatch.setattr(sys, "argv", ["pipeline.py", "--config", str(config_path), "--todas"])

    pipeline.main()

    assert (tmp_path / "saida" / "historico_bandeiras.xlsx").exists()
