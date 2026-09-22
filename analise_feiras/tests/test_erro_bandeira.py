import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from utils import normalize_cnpj, normalize_text, tokenizar  # noqa: E402

import analise_erro_bandeira as erro_bandeira  # noqa: E402


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
                "tipo_cliente": "Tipo de cliente",
                "data_pedido": "Data do pedido (original)",
                "faturado_liquido": "Faturado líquido (R$)",
                "desconto_aplicado_pct": "Desconto comercial faturado (%)",
                "numero_nota": "Numero da Nota",
                "quantidade_faturada": "Quantidade faturada",
                "distribuidor": "Nome do distribuidor",
                "grupo_clientes": "Grupo de clientes",
            },
        },
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
            "colunas": {
                "grupo_clientes": "Grupo de clientes",
                "tabela_1": "Tabela 1",
                "tabela_2": "Tabela 2",
            },
        },
        "cnpjs_ca": [
            {
                "arquivo": "rotulos_lojas_*.csv",
                "chave": "cnpj",
                "coluna_rotulo": "rotulo",
                "rotulo_valido": "SELL_OUT_CA",
            }
        ],
        "saida": {"pasta": "saida", "nome_arquivo": "Análise Erro Bandeira.xlsx"},
    }


def test_rotulo_bate_na_tabela_ignora_ordem_e_aceita_abreviacao():
    # caso real 1: mesmas palavras, ordem trocada, com prefixo extra
    rot = tokenizar(normalize_text("GENERICO_D1000"))
    tab = tokenizar(normalize_text("Tabela Agregadora - D1000_GENERICO"))
    assert erro_bandeira._rotulo_bate_na_tabela(rot, tab)

    # caso real 2: abreviação (AUT é prefixo de AUTORIZADOR)
    rot = tokenizar(normalize_text("CANAL_AUT"))
    tab = tokenizar(normalize_text("Canal Autorizador"))
    assert erro_bandeira._rotulo_bate_na_tabela(rot, tab)

    # não pode bater com uma tabela genérica que não tem a palavra específica
    rot = tokenizar(normalize_text("GENERICO_D1000"))
    tab = tokenizar(normalize_text("DEFAULT GENERICO CA"))
    assert not erro_bandeira._rotulo_bate_na_tabela(rot, tab)

    # rótulo vazio nunca bate
    assert not erro_bandeira._rotulo_bate_na_tabela((), tab)
    assert not erro_bandeira._rotulo_bate_na_tabela(None, tab)


def test_carregar_cnpjs_ca_ignora_fonte_com_arquivo_ausente(tmp_path: Path, capsys):
    # Bug real (na versão anterior, dentro do pipeline.py): uma fonte de
    # cnpjs_ca com arquivo ausente travava a análise inteira. Aqui, uma
    # fonte ausente é só ignorada (com aviso), sem quebrar as outras.
    rotulos_df = pd.DataFrame({"cnpj": ["11.111.111/0001-11"], "rotulo": ["SELL_OUT_CA"]})
    rotulos_df.to_csv(tmp_path / "rotulos_lojas_20260904.csv", sep=";", index=False)

    erro_bandeira.BASE_DIR = tmp_path

    cfg = {
        "cnpjs_ca": [
            {
                "arquivo": "rotulos_lojas_*.csv",
                "chave": "cnpj",
                "coluna_rotulo": "rotulo",
                "rotulo_valido": "SELL_OUT_CA",
            },
            {
                "arquivo": "painel_nv_ausente_*.xlsx",  # não existe em tmp_path
                "chave": "cnpj",
                "coluna_rotulo": "rotulo",
                "rotulo_valido": "NAO_VISITADO",
            },
        ]
    }

    cnpjs = erro_bandeira.carregar_cnpjs_ca(cfg)

    assert cnpjs == {normalize_cnpj("11.111.111/0001-11")}
    assert "Aviso" in capsys.readouterr().out


def test_calcular_erro_bandeira_e_main(tmp_path: Path, capsys):
    # Grupo "RAIA DROGASIL" no Painel x Tabela: Tabela 1 RAIA_GENERICO,
    # Tabela 2 RAIA CA (mesmo EAN, descontos diferentes: 40% e 60,90%).
    # pedido 7001: Tabela errada, CNPJ cadastrado como SELL_OUT_CA (CA) ->
    #              entra no relatório com tabela_correta="RAIA CA", 60,90%
    # pedido 7002: Tabela errada, CNPJ sem cadastro CA -> entra com
    #              tabela_correta="RAIA_GENERICO", 40%
    # pedido 7003: Tabela "RAIA CA" (bate com a Tabela 2) -> correto, FORA
    # pedido 7004: Tabela "RAIA_GENERICO" (bate com a Tabela 1) -> correto, FORA
    # pedido 7005: Grupo "GRUPO DESCONHECIDO" (fora do Painel x Tabela) ->
    #              entra com tabela_correta avisando que não está cadastrado
    # pedido 7006: CNPJ fora do controle_bandeira -> nem entra na análise
    base_df = pd.DataFrame(
        {
            "Tabela de negociação": [
                "QUALQUER OUTRA TABELA",
                "QUALQUER OUTRA TABELA",
                "RAIA CA",
                "RAIA_GENERICO",
                "QUALQUER TABELA",
                "RAIA CA",
            ],
            "CNPJ": [
                "10.000.000/0001-00",
                "20.000.000/0001-00",
                "30.000.000/0001-00",
                "40.000.000/0001-00",
                "50.000.000/0001-00",
                "99.999.999/0001-99",
            ],
            "EAN": [
                "5555555555555",
                "5555555555555",
                "5555555555555",
                "5555555555555",
                "6666666666666",
                "5555555555555",
            ],
            "Id pedido": ["7001", "7002", "7003", "7004", "7005", "7006"],
            "Tipo de cliente": ["REDES CORPORATIVAS"] * 6,
            "Data do pedido (original)": ["10/05/2026 10:00"] * 6,
            "Faturado líquido (R$)": ["55,0"] * 6,
            "Desconto comercial faturado (%)": ["20"] * 6,
            "Numero da Nota": ["1001", "1002", "1003", "1004", "1005", "1006"],
            "Quantidade faturada": ["7"] * 6,
            "Nome do distribuidor": ["PANPHARMA"] * 6,
            "Grupo de clientes": [
                "RAIA DROGASIL",
                "RAIA DROGASIL",
                "RAIA DROGASIL",
                "RAIA DROGASIL",
                "GRUPO DESCONHECIDO",
                "RAIA DROGASIL",
            ],
        }
    )

    painel_bandeira_df = pd.DataFrame(
        {
            "CNPJ Ajustado": [
                "10.000.000/0001-00",
                "20.000.000/0001-00",
                "30.000.000/0001-00",
                "40.000.000/0001-00",
                "50.000.000/0001-00",
            ],
            "id_bandeira": ["10366"] * 5,
            "desc_bandeira": ["RAIA DROGASIL", "RAIA DROGASIL", "RAIA DROGASIL", "RAIA DROGASIL", "OUTRA REDE"],
            "perfil_bandeira": ["CALENDARIO"] * 5,
            "razao_social": ["RAIA DROGASIL SA"] * 5,
            "cidade": ["TRES LAGOAS"] * 5,
            "estado": ["MS"] * 5,
        }
    )

    painel_tabela_df = pd.DataFrame(
        {
            "Grupo de clientes": ["RAIA DROGASIL"],
            "Tabela 1": ["RAIA_GENERICO"],
            "Tabela 2": ["RAIA CA"],
        }
    )

    condicao_df = pd.DataFrame(
        {
            "EAN FORMATADO": ["5555555555555", "5555555555555"],
            "Tabela": ["RAIA_GENERICO", "RAIA CA"],
            "Desconto Atual": [0.40, 0.6090],
        }
    )

    rotulos_lojas_df = pd.DataFrame({"cnpj": ["10.000.000/0001-00"], "rotulo": ["SELL_OUT_CA"]})

    (tmp_path / "saida").mkdir()
    base_df.to_excel(tmp_path / "base_pedidos.xlsx", index=False)
    with pd.ExcelWriter(tmp_path / "Painel_Bandeira_2026_08.xlsx") as w:
        painel_bandeira_df.to_excel(w, sheet_name="Dados", index=False)
    with pd.ExcelWriter(tmp_path / "condicao_comercial.xlsx") as w:
        condicao_df.to_excel(w, sheet_name="Dados", index=False)
    with pd.ExcelWriter(tmp_path / "Panel_x_Tabela.xlsx") as w:
        painel_tabela_df.to_excel(w, sheet_name="Planilha1", index=False)
    rotulos_lojas_df.to_csv(tmp_path / "rotulos_lojas_20260904_141843.csv", sep=";", index=False)

    cfg = _montar_config(tmp_path)
    erro_bandeira.BASE_DIR = tmp_path

    df_base = erro_bandeira.carregar_base(cfg)
    print(f"Base carregada: {len(df_base)} linhas.")
    df_erros = erro_bandeira.calcular_erro_bandeira(df_base, cfg)

    assert not df_erros.empty
    # 7003/7004 (Tabela correta) e 7006 (fora do controle_bandeira) não entram
    assert set(df_erros["Id pedido"]) == {"7001", "7002", "7005"}

    resultado = erro_bandeira.montar_saida(df_erros, cfg)

    # Desconto comercial faturado (%) só é usado internamente pro cálculo
    # reverso — não faz parte da saída pedida
    assert "Desconto comercial faturado (%)" not in resultado.columns
    for coluna in (
        "Tabela de negociação",
        "CNPJ",
        "EAN",
        "Id pedido",
        "Tipo de cliente",
        "Data do pedido (original)",
        "Faturado líquido (R$)",
        "Numero da Nota",
        "Quantidade faturada",
        "Nome do distribuidor",
        "Grupo de clientes",
        "id_bandeira",
        "bandeira",
        "perfil_bandeira",
        "razao_social",
        "cidade",
        "estado",
        "tabela_correta",
        "desconto_correto_pct",
        "preco_sem_desconto",
        "faturamento_correto",
        "impacto_financeiro",
    ):
        assert coluna in resultado.columns

    tabelas_corretas = dict(zip(resultado["Id pedido"], resultado["tabela_correta"]))
    assert tabelas_corretas["7001"] == "RAIA CA"  # CNPJ CA
    assert tabelas_corretas["7002"] == "RAIA_GENERICO"  # CNPJ não-CA
    assert tabelas_corretas["7005"] == erro_bandeira.TABELA_NAO_CADASTRADA

    descontos = dict(zip(resultado["Id pedido"], resultado["desconto_correto_pct"]))
    assert round(descontos["7001"], 2) == 60.90
    assert round(descontos["7002"], 2) == 40.00
    assert pd.isna(descontos["7005"])

    # faturado líquido 55,0, desconto aplicado 20% -> preço sem desconto
    # 55,0 / 0,80 = 68,75
    precos = dict(zip(resultado["Id pedido"], resultado["preco_sem_desconto"]))
    assert round(precos["7001"], 2) == 68.75
    assert round(precos["7002"], 2) == 68.75

    faturamentos_corretos = dict(zip(resultado["Id pedido"], resultado["faturamento_correto"]))
    assert round(faturamentos_corretos["7001"], 2) == round(68.75 * (1 - 0.6090), 2)
    assert round(faturamentos_corretos["7002"], 2) == round(68.75 * (1 - 0.40), 2)
    assert pd.isna(faturamentos_corretos["7005"])  # Tabela correta desconhecida

    impactos = dict(zip(resultado["Id pedido"], resultado["impacto_financeiro"]))
    assert round(impactos["7001"], 2) == round(55.0 - 68.75 * (1 - 0.6090), 2)
    assert round(impactos["7002"], 2) == round(55.0 - 68.75 * (1 - 0.40), 2)
    assert pd.isna(impactos["7005"])

    # roda o script inteiro (main) também, ponta a ponta
    import sys as _sys

    monkeypatch_argv = _sys.argv
    _sys.argv = ["analise_erro_bandeira.py", "--config", str(tmp_path / "config_erro_bandeira.yaml")]
    try:
        # carregar_config lê do disco — grava o config nesse caminho
        import yaml

        with open(tmp_path / "config_erro_bandeira.yaml", "w", encoding="utf-8") as f:
            yaml.safe_dump(cfg, f, allow_unicode=True)
        erro_bandeira.main()
    finally:
        _sys.argv = monkeypatch_argv

    caminho_saida = tmp_path / "saida" / "Análise Erro Bandeira.xlsx"
    assert caminho_saida.exists()
    saida_lida = pd.read_excel(caminho_saida)
    assert set(saida_lida["Id pedido"].astype(str)) == {"7001", "7002", "7005"}


def test_calcular_erro_bandeira_sem_erro_nenhum(tmp_path: Path):
    # todos os pedidos batem com a Tabela 1/Tabela 2 do seu grupo -> nenhum
    # entra no relatório
    base_df = pd.DataFrame(
        {
            "Tabela de negociação": ["RAIA CA", "RAIA_GENERICO"],
            "CNPJ": ["10.000.000/0001-00", "20.000.000/0001-00"],
            "EAN": ["5555555555555", "5555555555555"],
            "Id pedido": ["8001", "8002"],
            "Tipo de cliente": ["REDES CORPORATIVAS"] * 2,
            "Data do pedido (original)": ["10/05/2026 10:00"] * 2,
            "Faturado líquido (R$)": ["55,0"] * 2,
            "Desconto comercial faturado (%)": ["20"] * 2,
            "Numero da Nota": ["1001", "1002"],
            "Quantidade faturada": ["7"] * 2,
            "Nome do distribuidor": ["PANPHARMA"] * 2,
            "Grupo de clientes": ["RAIA DROGASIL", "RAIA DROGASIL"],
        }
    )

    painel_bandeira_df = pd.DataFrame(
        {
            "CNPJ Ajustado": ["10.000.000/0001-00", "20.000.000/0001-00"],
            "id_bandeira": ["10366", "10366"],
            "desc_bandeira": ["RAIA DROGASIL", "RAIA DROGASIL"],
            "perfil_bandeira": ["CALENDARIO", "CALENDARIO"],
            "razao_social": ["RAIA DROGASIL SA", "RAIA DROGASIL SA"],
            "cidade": ["TRES LAGOAS", "TRES LAGOAS"],
            "estado": ["MS", "MS"],
        }
    )

    painel_tabela_df = pd.DataFrame(
        {"Grupo de clientes": ["RAIA DROGASIL"], "Tabela 1": ["RAIA_GENERICO"], "Tabela 2": ["RAIA CA"]}
    )

    (tmp_path / "saida").mkdir()
    base_df.to_excel(tmp_path / "base_pedidos.xlsx", index=False)
    with pd.ExcelWriter(tmp_path / "Painel_Bandeira_2026_08.xlsx") as w:
        painel_bandeira_df.to_excel(w, sheet_name="Dados", index=False)
    with pd.ExcelWriter(tmp_path / "Panel_x_Tabela.xlsx") as w:
        painel_tabela_df.to_excel(w, sheet_name="Planilha1", index=False)

    cfg = _montar_config(tmp_path)
    erro_bandeira.BASE_DIR = tmp_path

    df_base = erro_bandeira.carregar_base(cfg)
    df_erros = erro_bandeira.calcular_erro_bandeira(df_base, cfg)

    assert df_erros.empty
