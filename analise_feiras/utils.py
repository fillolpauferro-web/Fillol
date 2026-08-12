"""Funções utilitárias compartilhadas pelo pipeline de análise de feiras."""

from __future__ import annotations

import glob
import re
import unicodedata
from pathlib import Path

import pandas as pd


_ENCODINGS_CANDIDATAS = ("utf-8-sig", "utf-8", "cp1252", "latin1")


def _detectar_encoding(caminho: Path) -> str:
    """Tenta abrir o arquivo com cada encoding até uma que não quebre.
    latin1 nunca falha, então sempre sobra alguma opção no fim da lista.
    """
    for encoding in _ENCODINGS_CANDIDATAS:
        try:
            with open(caminho, "r", encoding=encoding) as f:
                f.readline()
            return encoding
        except UnicodeDecodeError:
            continue
    return "latin1"


def _detectar_separador(caminho: Path, encoding: str) -> str:
    """CSV exportado de Excel em português normalmente usa ';' (porque ','
    é o separador decimal); CSV "internacional" usa ','. Decide pela
    primeira linha do arquivo.
    """
    with open(caminho, "r", encoding=encoding, errors="replace") as f:
        primeira_linha = f.readline()
    return ";" if primeira_linha.count(";") > primeira_linha.count(",") else ","


def read_table(caminho: str | Path, aba: str | None = None) -> pd.DataFrame:
    """Lê .xlsx/.xls/.csv de forma transparente.

    `aba=None` usa a primeira aba (comportamento padrão do pandas para Excel).
    Para CSV/TSV, detecta automaticamente encoding e separador (';' ou ',').
    """
    caminho = Path(caminho)
    if not caminho.exists():
        raise FileNotFoundError(
            f"Arquivo não encontrado: {caminho}. Verifique o caminho no config.yaml."
        )

    if caminho.suffix.lower() in (".csv", ".tsv"):
        encoding = _detectar_encoding(caminho)
        sep = "\t" if caminho.suffix.lower() == ".tsv" else _detectar_separador(caminho, encoding)
        return pd.read_csv(
            caminho, sep=sep, dtype=str, keep_default_na=False, na_values=[""], encoding=encoding
        )

    aba_real = _resolver_aba(caminho, aba) if aba else 0
    return pd.read_excel(caminho, sheet_name=aba_real, dtype=str)


def _resolver_aba(caminho: Path, aba: str) -> str:
    """Acha o nome exato da aba ignorando maiúsculas/minúsculas e espaços
    nas pontas (ex.: config.yaml pede "dados" mas a planilha tem "DADOS").
    """
    nomes_abas = pd.ExcelFile(caminho).sheet_names
    aba_normalizada = aba.strip().lower()
    for nome in nomes_abas:
        if nome.strip().lower() == aba_normalizada:
            return nome
    raise ValueError(
        f"Aba '{aba}' não encontrada em {caminho.name}. Abas disponíveis: {nomes_abas}"
    )


def read_table_or_glob(padrao: str | Path, aba: str | None = None) -> pd.DataFrame:
    """Lê um arquivo único, ou vários arquivos de uma vez se `padrao` tiver
    curinga (* ? []) — nesse caso empilha todos num único DataFrame.

    Ex.: "dados/planilha_base_vendas_*.xlsx" lê e junta todos os arquivos
    que começam com esse prefixo.
    """
    padrao_str = str(padrao)
    if not any(ch in padrao_str for ch in "*?["):
        return read_table(padrao_str, aba)

    arquivos = sorted(glob.glob(padrao_str))
    if not arquivos:
        raise FileNotFoundError(
            f"Nenhum arquivo encontrado para o padrão: {padrao_str}. "
            "Verifique o caminho/prefixo no config.yaml."
        )

    print(f"{len(arquivos)} arquivo(s) encontrado(s) para '{Path(padrao_str).name}':")
    for arq in arquivos:
        print(f"  - {Path(arq).name}")

    partes = [read_table(arq, aba) for arq in arquivos]
    return pd.concat(partes, ignore_index=True)


def normalize_text(valor) -> str:
    """Maiúsculas, sem acento, sem espaço nas pontas — para comparações robustas."""
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ""
    texto = str(valor).strip().upper()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = re.sub(r"\s+", " ", texto)
    return texto


def normalize_cnpj(valor) -> str:
    """Mantém só os dígitos do CNPJ, para casar bases com/sem máscara."""
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ""
    return re.sub(r"\D", "", str(valor))


def to_datetime(series: pd.Series) -> pd.Series:
    """Converte para data aceitando formatos mistos: aaaa-mm-dd (ISO, sem
    ambiguidade) e dd/mm/aaaa (padrão BR, dayfirst=True).

    Usar dayfirst=True direto com format="mixed" faz o pandas inverter dia/mês
    também em timestamps ISO (bug observado: "2026-05-10" virava 10/10 em vez
    de 10/05), então cada formato é tratado separadamente.
    """
    texto = series.astype(str).str.strip()
    resultado = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")

    eh_iso = texto.str.match(r"^\d{4}-\d{2}-\d{2}")
    if eh_iso.any():
        resultado.loc[eh_iso] = pd.to_datetime(texto.loc[eh_iso], errors="coerce")
    if (~eh_iso).any():
        resultado.loc[~eh_iso] = pd.to_datetime(texto.loc[~eh_iso], errors="coerce", dayfirst=True)

    return resultado


def to_numeric(series: pd.Series) -> pd.Series:
    """Converte texto numérico pt-BR (1.234,56 ou 1234,56) para float."""
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")

    limpo = (
        series.astype(str)
        .str.strip()
        .str.replace(r"^'", "", regex=True)  # aspas de "número como texto" do Excel
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
    )
    return pd.to_numeric(limpo, errors="coerce")
