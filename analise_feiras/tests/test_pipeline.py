import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from pipeline import CHECK_ERRO, CHECK_OK, carregar_base, rodar_matriz  # noqa: E402
from utils import normalize_cnpj, normalize_ean, to_datetime, to_numeric  # noqa: E402


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


def test_normalize_ean_remove_aspas_das_duas_pontas():
    assert normalize_ean("'7896422511865'") == "7896422511865"
    assert normalize_ean("'7896422511865") == "7896422511865"
    assert normalize_ean(" 7896422511865 ") == "7896422511865"
    assert normalize_ean(None) == ""


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
    base_df = pd.DataFrame(
        {
            "Tipo de cliente": ["ASSOCIATIVISMO", "ASSOCIATIVISMO", "ASSOCIATIVISMO", "REDE INDIRETA"],
            "CNPJ": [
                "01.672.858/0001-65",
                "'06052566000143'",
                "01.672.858/0001-65",
                "06.052.566/0001-43",
            ],
            "Id pedido": ["17275", "17300", "17400", "17778"],
            "EAN": ["7896422511865", "'7896422511865'", "7896422511865", "7896422514651"],
            "Tabela de negociação": [
                "FEIRA NEGOCIOS CA",
                "FEIRA NEGOCIOS CA",
                "FEIRA NEGOCIOS CA",
                "DEFAULT GENERICO CA",
            ],
            "Data do pedido (original)": [
                "10/05/2026 18:34",
                "10/05/2026 10:00",
                "15/07/2026 09:00",  # fora do período do controle (maio)
                "05/05/2026 19:16",
            ],
            "Faturado líquido (R$)": ["27,3", "27,3", "27,3", "53,42"],
            "Desconto comercial faturado (%)": ["56,87", "56,87", "56,87", "69,36"],
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
    assert len(df_base) == 4

    resultado = rodar_matriz("Feira", cfg["matrizes"][0], df_base, cfg)
    assert resultado is not None
    assert len(resultado) == 3  # só as 3 linhas de FEIRA NEGOCIOS CA

    checks = dict(zip(resultado["Id pedido"], resultado["Check"]))
    assert checks["17275"] == CHECK_OK
    assert checks["17300"] == CHECK_ERRO
    assert checks["17400"] == CHECK_ERRO  # cadastrado, mas fora do período

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

    assert (tmp_path / "saida" / "Feira_analise.xlsx").exists()
