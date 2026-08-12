import sys
from pathlib import Path

import pandas as pd
import yaml

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from pipeline import CHECK_ANTES, CHECK_OK, carregar_base, carregar_config, rodar_matriz  # noqa: E402


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
            "arquivo": "Condicao_comercial.xlsx",
            "aba": "dados",
            "colunas": {
                "chave_cnpj": "CNPJ",
                "chave_ean": "EAN",
                "desconto_correto_pct": "Desconto (%)",
            },
        },
        "matrizes": [
            {
                "nome": "Feira",
                "ativo": True,
                "palavra_chave": "FEIRA",
                "arquivo_controle": "Controle_Feiras.xlsx",
                "aba_controle": "dados",
                "chave_controle": "Tabela de negociação",
                "colunas_trazidas": {
                    "feira": "Feira",
                    "data_inicial": "Data inicial",
                    "vigencia": "Vigência",
                },
            }
        ],
        "saida": {"pasta": "saida"},
    }


def test_pipeline_end_to_end(tmp_path: Path):
    # pedido 1: dentro da vigência da feira -> OK
    # pedido 2: antes da vigência -> erro operacional, com condição comercial correta cadastrada
    base_df = pd.DataFrame(
        {
            "Tipo de cliente": ["ASSOCIATIVISMO", "ASSOCIATIVISMO", "REDE INDIRETA"],
            "CNPJ": ["01.672.858/0001-65", "01.672.858/0001-65", "06.052.566/0001-43"],
            "Id pedido": ["17275", "17300", "17778"],
            "EAN": ["7896422511865", "7896422511865", "7896422514651"],
            "Tabela de negociação": ["FEIRA NEGOCIOS CA", "FEIRA NEGOCIOS CA", "DEFAULT GENERICO CA"],
            "Data do pedido (original)": ["10/05/2026 18:34", "01/01/2026 10:00", "05/05/2026 19:16"],
            "Faturado líquido (R$)": ["27,3", "27,3", "53,42"],
            "Desconto comercial faturado (%)": ["56,87", "56,87", "69,36"],
        }
    )

    controle_df = pd.DataFrame(
        {
            "Tabela de negociação": ["FEIRA NEGOCIOS CA"],
            "Feira": ["Feira de Negócios CA 2026"],
            "Data inicial": ["01/05/2026"],
            "Vigência": ["31/05/2026"],
        }
    )

    condicao_df = pd.DataFrame(
        {
            "CNPJ": ["01.672.858/0001-65"],
            "EAN": ["7896422511865"],
            "Desconto (%)": ["30"],
        }
    )

    (tmp_path / "saida").mkdir()
    base_df.to_excel(tmp_path / "base_pedidos.xlsx", index=False)
    with pd.ExcelWriter(tmp_path / "Controle_Feiras.xlsx") as w:
        controle_df.to_excel(w, sheet_name="dados", index=False)
    with pd.ExcelWriter(tmp_path / "Condicao_comercial.xlsx") as w:
        condicao_df.to_excel(w, sheet_name="dados", index=False)

    cfg = _montar_config(tmp_path)
    import pipeline

    pipeline.BASE_DIR = tmp_path

    df_base = carregar_base(cfg)
    assert len(df_base) == 3  # a linha REDE INDIRETA existe na base mas não é da matriz Feira

    resultado = rodar_matriz("Feira", cfg["matrizes"][0], df_base, cfg)
    assert resultado is not None
    assert len(resultado) == 2  # só as 2 linhas de FEIRA NEGOCIOS CA

    checks = dict(zip(resultado["Id pedido"], resultado["Check"]))
    assert checks["17275"] == CHECK_OK
    assert checks["17300"] == CHECK_ANTES

    linha_erro = resultado.loc[resultado["Id pedido"] == "17300"].iloc[0]
    assert round(linha_erro["preco_sem_desconto"], 2) == round(27.3 / (1 - 0.5687), 2)
    preco_sem_desconto_esperado = 27.3 / (1 - 0.5687)
    preco_correto_esperado = preco_sem_desconto_esperado * (1 - 0.30)
    assert round(linha_erro["preco_liquido_desconto_correto"], 2) == round(preco_correto_esperado, 2)

    assert (tmp_path / "saida" / "Feira_filtrado.xlsx").exists()
    assert (tmp_path / "saida" / "Feira_analise.xlsx").exists()
