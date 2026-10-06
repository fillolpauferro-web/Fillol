import sys
from pathlib import Path

import pandas as pd
import yaml

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import analise_erro_bandeira as erro_bandeira  # noqa: E402
import consolidado_erros_operacionais as consolidado  # noqa: E402
import pipeline  # noqa: E402


def _colunas_base() -> dict:
    return {
        "tabela_negociacao": "Tabela de negociação",
        "cnpj": "CNPJ",
        "ean": "EAN",
        "id_pedido": "Id pedido",
        "tipo_cliente": "Tipo de cliente",
        "data_pedido": "Data do pedido (original)",
        "faturado_liquido": "Faturado líquido (R$)",
        "desconto_aplicado_pct": "Desconto comercial faturado (%)",
        "numero_nota": "Numero da Nota",
        "quantidade_faturada": "Quantidade faturada",
        "distribuidor": "Nome do distribuidor",
        "grupo_clientes": "Grupo de clientes",
    }


def test_consolidado_junta_erros_do_pipeline_e_da_bandeira(tmp_path: Path, capsys):
    # pedido P1: CNPJ cadastrado como NAO_VISITADO no Painel_NV, mas a
    #            Tabela de negociação não é "Canal Autorizador" -> Check
    #            Erro Operacional na matriz CanalAutorizador (pipeline.py)
    # pedido P2: CNPJ do Painel_Bandeira, Grupo "RAIA DROGASIL", Tabela
    #            errada (não bate com RAIA_GENERICO nem RAIA CA) -> erro
    #            na Análise Erro Bandeira (analise_erro_bandeira.py)
    base_df = pd.DataFrame(
        {
            "Tabela de negociação": ["FEIRA ERRADA", "QUALQUER OUTRA TABELA"],
            "CNPJ": ["11.111.111/0001-11", "22.222.222/0001-22"],
            "EAN": ["1111111111111", "2222222222222"],
            "Id pedido": ["P1", "P2"],
            "Tipo de cliente": ["REDES CORPORATIVAS", "REDES CORPORATIVAS"],
            "Data do pedido (original)": ["10/05/2026 10:00", "11/05/2026 10:00"],
            "Faturado líquido (R$)": ["100,0", "200,0"],
            "Desconto comercial faturado (%)": ["20", "25"],
            "Numero da Nota": ["9001", "9002"],
            "Quantidade faturada": ["5", "9"],
            "Nome do distribuidor": ["PANPHARMA", "PANPHARMA"],
            "Grupo de clientes": ["OUTRO GRUPO", "RAIA DROGASIL"],
        }
    )

    painel_nv_df = pd.DataFrame({"CNPJ ajustado": ["11.111.111/0001-11"], "Rotulo": ["NAO_VISITADO"]})
    painel_bandeira_df = pd.DataFrame(
        {
            "CNPJ Ajustado": ["22.222.222/0001-22"],
            "id_bandeira": ["999"],
            "desc_bandeira": ["RAIA DROGASIL"],
            "perfil_bandeira": ["CALENDARIO"],
            "razao_social": ["RAIA DROGASIL SA"],
            "cidade": ["LAURO DE FREITAS"],
            "estado": ["BA"],
        }
    )
    condicao_df = pd.DataFrame(
        {
            "EAN FORMATADO": ["1111111111111", "2222222222222"],
            "Tabela": ["Canal Autorizador", "RAIA_GENERICO"],
            "Desconto Atual": [0.30, 0.35],
        }
    )
    painel_tabela_df = pd.DataFrame(
        {"Grupo de clientes": ["RAIA DROGASIL"], "Tabela 1": ["RAIA_GENERICO"], "Tabela 2": ["RAIA CA"]}
    )

    (tmp_path / "saida").mkdir()
    base_df.to_excel(tmp_path / "base_pedidos.xlsx", index=False)
    with pd.ExcelWriter(tmp_path / "Painel_NV_2026_08.xlsx") as w:
        painel_nv_df.to_excel(w, sheet_name="Dados", index=False)
    with pd.ExcelWriter(tmp_path / "Painel_Bandeira_2026_08.xlsx") as w:
        painel_bandeira_df.to_excel(w, sheet_name="Dados", index=False)
    with pd.ExcelWriter(tmp_path / "condicao_comercial.xlsx") as w:
        condicao_df.to_excel(w, sheet_name="Dados", index=False)
    with pd.ExcelWriter(tmp_path / "Panel_x_Tabela.xlsx") as w:
        painel_tabela_df.to_excel(w, sheet_name="Planilha1", index=False)

    colunas = _colunas_base()

    cfg = {
        "base": {"arquivo": "base_pedidos.xlsx", "aba": None, "colunas": colunas},
        "condicao_comercial": {
            "arquivo": "condicao_comercial.xlsx",
            "aba": "Dados",
            "colunas": {
                "chave_ean": "EAN FORMATADO",
                "desconto_correto_pct": "Desconto Atual",
                "chave_tabela": "Tabela",
            },
            "desconto_em_fracao": True,
        },
        "matrizes": [
            {
                "nome": "CanalAutorizador",
                "tipo": "cnpj",
                "ativo": True,
                "palavra_chave": "CANAL AUTORIZADOR",
                "arquivo_controle": "Painel_NV_*.xlsx",
                "aba_controle": "Dados",
                "chave_controle": "CNPJ ajustado",
                "coluna_rotulo_controle": "Rotulo",
                "rotulo_valido_controle": "NAO_VISITADO",
                "colunas_trazidas": {},
                "colunas_data": [],
            }
        ],
        "saida": {"pasta": "saida"},
    }

    cfg_bandeira = {
        "base": {"arquivo": "base_pedidos.xlsx", "aba": None, "colunas": colunas},
        "condicao_comercial": cfg["condicao_comercial"],
        "controle_bandeira": {
            "arquivo": "Painel_Bandeira_*.xlsx",
            "aba": "Dados",
            "chave": "CNPJ Ajustado",
            "colunas_trazidas": {
                "id_bandeira": "id_bandeira",
                "bandeira": "desc_bandeira",
                "perfil_bandeira": "perfil_bandeira",
                "razao_social": "razao_social",
                "cidade": "cidade",
                "estado": "estado",
            },
        },
        "painel_tabela": {
            "arquivo": "Panel_x_Tabela.xlsx",
            "aba": "Planilha1",
            "colunas": {"grupo_clientes": "Grupo de clientes", "tabela_1": "Tabela 1", "tabela_2": "Tabela 2"},
        },
        "cnpjs_ca": [],
        "saida": {"pasta": "saida", "nome_arquivo": "Análise Erro Bandeira.xlsx"},
    }

    with open(tmp_path / "config.yaml", "w", encoding="utf-8") as f:
        yaml.safe_dump(cfg, f, allow_unicode=True)
    with open(tmp_path / "config_erro_bandeira.yaml", "w", encoding="utf-8") as f:
        yaml.safe_dump(cfg_bandeira, f, allow_unicode=True)

    pipeline.BASE_DIR = tmp_path
    erro_bandeira.BASE_DIR = tmp_path
    consolidado.BASE_DIR = tmp_path

    argv_original = sys.argv
    sys.argv = [
        "consolidado_erros_operacionais.py",
        "--config",
        str(tmp_path / "config.yaml"),
        "--config-bandeira",
        str(tmp_path / "config_erro_bandeira.yaml"),
    ]
    try:
        consolidado.main()
    finally:
        sys.argv = argv_original

    saida_texto = capsys.readouterr().out
    assert "Aviso: matriz" in saida_texto  # as outras 4 matrizes de Check não existem nesse cfg mínimo

    caminho_saida = tmp_path / "saida" / "Consolidado_Erros_Operacionais.xlsx"
    assert caminho_saida.exists()

    resultado = pd.read_excel(caminho_saida)
    assert list(resultado.columns) == consolidado.COLUNAS_SAIDA
    assert set(zip(resultado["fonte"], resultado["id_pedido"].astype(str))) == {
        ("CanalAutorizador", "P1"),
        ("Bandeira", "P2"),
    }

    por_pedido = resultado.set_index("id_pedido")
    p1 = por_pedido.loc["P1"]
    assert p1["tabela_correta"] == "CANAL AUTORIZADOR"
    assert round(p1["desconto_correto_pct"], 2) == 30.0
    preco_sem_desconto_p1 = 100.0 / (1 - 0.20)
    assert round(p1["preco_liquido_desconto_correto"], 2) == round(preco_sem_desconto_p1 * (1 - 0.30), 2)
    assert round(p1["diferenca_faturamento"], 2) == round(100.0 - preco_sem_desconto_p1 * (1 - 0.30), 2)

    p2 = por_pedido.loc["P2"]
    assert p2["tabela_correta"] == "RAIA_GENERICO"
    assert round(p2["desconto_correto_pct"], 2) == 35.0
    preco_sem_desconto_p2 = 200.0 / (1 - 0.25)
    assert round(p2["preco_liquido_desconto_correto"], 2) == round(preco_sem_desconto_p2 * (1 - 0.35), 2)
