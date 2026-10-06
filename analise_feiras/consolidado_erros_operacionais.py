"""Consolidado de Erros Operacionais — junta numa única planilha os
pedidos com Check = Erro Operacional de todas as matrizes de desconto
(Feira, Canal Autorizador, Sell Out, Atacarejo Conecta, Atacarejo Conecta
CA, rodadas via pipeline.py/config.yaml) mais os erros da Análise Erro
Bandeira (script separado, analise_erro_bandeira.py/
config_erro_bandeira.yaml) — pra ter uma visão única de "todos os erros
operacionais de desconto", sem precisar abrir um arquivo por matriz.

Cada matriz continua salvando seu próprio arquivo normalmente (nada muda
nisso); este script só roda tudo de novo e consolida as linhas de erro
num arquivo extra, "Consolidado_Erros_Operacionais.xlsx", com uma coluna
"fonte" indicando de qual matriz veio cada linha.

Colunas da saída: fonte, tabela_negociacao, cnpj, ean, id_pedido,
tipo_cliente, data_pedido, desconto_aplicado_pct, faturado_liquido,
numero_nota, quantidade_faturada, distribuidor, grupo_clientes,
tabela_correta, desconto_correto_pct, preco_sem_desconto,
preco_liquido_desconto_correto, diferenca_faturamento.

"tabela_correta" é a palavra_chave fixa da matriz (Feira, Canal
Autorizador, Sell Out, Atacarejo Conecta, Atacarejo Conecta CA) ou, pra
Bandeira, a Tabela específica daquele pedido (Painel x Tabela).

Uso:
    python consolidado_erros_operacionais.py
    python consolidado_erros_operacionais.py --config config.yaml --config-bandeira config_erro_bandeira.yaml
"""

from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

import analise_erro_bandeira as erro_bandeira
import pipeline

BASE_DIR = Path(__file__).resolve().parent

# matrizes de Check (tipo "tabela"/"cnpj") do pipeline.py que entram no
# consolidado — Bandeira é tratada separadamente (vem de outro script).
MATRIZES_CHECK = ["Feira", "CanalAutorizador", "SellOut", "AtacarejoConecta", "AtacarejoConectaCA"]

COLUNAS_SAIDA = [
    "fonte",
    "tabela_negociacao",
    "cnpj",
    "ean",
    "id_pedido",
    "tipo_cliente",
    "data_pedido",
    "desconto_aplicado_pct",
    "faturado_liquido",
    "numero_nota",
    "quantidade_faturada",
    "distribuidor",
    "grupo_clientes",
    "tabela_correta",
    "desconto_correto_pct",
    "preco_sem_desconto",
    "preco_liquido_desconto_correto",
    "diferenca_faturamento",
]


def _erros_matriz_check(nome_matriz: str, cfg: dict, df_base: pd.DataFrame) -> pd.DataFrame:
    """Roda uma matriz de Check do pipeline.py (salva o arquivo dela
    normalmente, como sempre) e devolve só as linhas de Erro Operacional,
    já no formato comum do consolidado.
    """
    matriz_cfg = next((m for m in cfg["matrizes"] if m["nome"] == nome_matriz), None)
    if matriz_cfg is None:
        print(f"!!! Aviso: matriz '{nome_matriz}' não encontrada no config.yaml — pulando.")
        return pd.DataFrame(columns=COLUNAS_SAIDA)

    resultado = pipeline.rodar_matriz(nome_matriz, matriz_cfg, df_base, cfg)
    if resultado is None or "Check" not in resultado.columns:
        return pd.DataFrame(columns=COLUNAS_SAIDA)

    erros = resultado[resultado["Check"] == pipeline.CHECK_ERRO]
    if erros.empty:
        return pd.DataFrame(columns=COLUNAS_SAIDA)

    colunas = cfg["base"]["colunas"]
    return pd.DataFrame(
        {
            "fonte": nome_matriz,
            "tabela_negociacao": erros[colunas["tabela_negociacao"]],
            "cnpj": erros[colunas["cnpj"]],
            "ean": erros[colunas["ean"]],
            "id_pedido": erros[colunas["id_pedido"]],
            "tipo_cliente": erros[colunas["tipo_cliente"]],
            "data_pedido": erros[colunas["data_pedido"]],
            "desconto_aplicado_pct": erros[colunas["desconto_aplicado_pct"]],
            "faturado_liquido": erros[colunas["faturado_liquido"]],
            "numero_nota": erros[colunas["numero_nota"]],
            "quantidade_faturada": erros[colunas["quantidade_faturada"]],
            "distribuidor": erros[colunas["distribuidor"]],
            "grupo_clientes": erros[colunas["grupo_clientes"]],
            "tabela_correta": matriz_cfg.get("palavra_chave"),
            "desconto_correto_pct": erros["desconto_correto_pct"],
            "preco_sem_desconto": erros["preco_sem_desconto"],
            "preco_liquido_desconto_correto": erros["preco_liquido_desconto_correto"],
            "diferenca_faturamento": erros["diferenca_faturamento"],
        }
    )


def _erros_bandeira(cfg_bandeira: dict) -> pd.DataFrame:
    """Roda a Análise Erro Bandeira (analise_erro_bandeira.py) e devolve
    os erros já no formato comum do consolidado.
    """
    df_base = erro_bandeira.carregar_base(cfg_bandeira)
    df_erros = erro_bandeira.calcular_erro_bandeira(df_base, cfg_bandeira)
    if df_erros.empty:
        return pd.DataFrame(columns=COLUNAS_SAIDA)

    colunas = cfg_bandeira["base"]["colunas"]
    return pd.DataFrame(
        {
            "fonte": "Bandeira",
            "tabela_negociacao": df_erros[colunas["tabela_negociacao"]],
            "cnpj": df_erros[colunas["cnpj"]],
            "ean": df_erros[colunas["ean"]],
            "id_pedido": df_erros[colunas["id_pedido"]],
            "tipo_cliente": df_erros[colunas["tipo_cliente"]],
            "data_pedido": df_erros[colunas["data_pedido"]],
            "desconto_aplicado_pct": df_erros[colunas["desconto_aplicado_pct"]],
            "faturado_liquido": df_erros[colunas["faturado_liquido"]],
            "numero_nota": df_erros[colunas["numero_nota"]],
            "quantidade_faturada": df_erros[colunas["quantidade_faturada"]],
            "distribuidor": df_erros[colunas["distribuidor"]],
            "grupo_clientes": df_erros[colunas["grupo_clientes"]],
            "tabela_correta": df_erros["tabela_correta"],
            "desconto_correto_pct": df_erros["desconto_correto_pct"],
            "preco_sem_desconto": df_erros["preco_sem_desconto"],
            "preco_liquido_desconto_correto": df_erros["faturamento_correto"],
            "diferenca_faturamento": df_erros["impacto_financeiro"],
        }
    )


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--config", default=str(BASE_DIR / "config.yaml"), help="Caminho do config.yaml")
    parser.add_argument(
        "--config-bandeira",
        default=str(BASE_DIR / "config_erro_bandeira.yaml"),
        help="Caminho do config_erro_bandeira.yaml",
    )
    args = parser.parse_args()

    pipeline.BASE_DIR = BASE_DIR
    erro_bandeira.BASE_DIR = BASE_DIR

    cfg = pipeline.carregar_config(args.config)
    cfg_bandeira = erro_bandeira.carregar_config(args.config_bandeira)

    df_base = pipeline.carregar_base(cfg)
    print(f"Base carregada: {len(df_base)} linhas.")

    partes = []
    for nome in MATRIZES_CHECK:
        print(f"\n--- Calculando erros de {nome} ---")
        try:
            partes.append(_erros_matriz_check(nome, cfg, df_base))
        except Exception as e:
            print(f"!!! Erro ao calcular '{nome}': {e}. Pulando.")

    print("\n--- Calculando erros de Bandeira ---")
    try:
        partes.append(_erros_bandeira(cfg_bandeira))
    except Exception as e:
        print(f"!!! Erro ao calcular 'Bandeira': {e}. Pulando.")

    consolidado = pd.concat(partes, ignore_index=True) if partes else pd.DataFrame(columns=COLUNAS_SAIDA)
    if consolidado.empty:
        print("\nNenhum erro operacional de desconto encontrado em nenhuma matriz.")
        return

    print(f"\n{len(consolidado)} pedidos com erro operacional de desconto, no total.")
    print(consolidado["fonte"].value_counts().to_string())

    pasta_saida = BASE_DIR / cfg["saida"]["pasta"]
    pasta_saida.mkdir(parents=True, exist_ok=True)
    caminho_final = pasta_saida / "Consolidado_Erros_Operacionais.xlsx"
    consolidado.to_excel(caminho_final, index=False)
    print(f"Resultado salvo em {caminho_final}")

    impacto = consolidado["diferenca_faturamento"].sum(skipna=True)
    print(f"Impacto financeiro total (faturado - correto): R$ {impacto:,.2f}")


if __name__ == "__main__":
    main()
