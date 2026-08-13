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


def read_table_mais_recente(padrao: str | Path, aba: str | None = None) -> pd.DataFrame:
    """Como read_table, mas aceita curinga (* ? []) no caminho — usa só o
    arquivo mais recente (data de modificação) entre os que baterem com o
    padrão, ao invés de juntar todos como read_table_or_glob faz.

    Útil para arquivos de controle cujo nome muda a cada exportação (ex.:
    "Painel_NV_2026_08.xlsx"), onde só a versão mais atual deve valer.
    """
    padrao_str = str(padrao)
    if not any(ch in padrao_str for ch in "*?["):
        return read_table(padrao_str, aba)

    arquivos = glob.glob(padrao_str)
    if not arquivos:
        raise FileNotFoundError(
            f"Nenhum arquivo encontrado para o padrão: {padrao_str}. "
            "Verifique o caminho/prefixo no config.yaml."
        )

    mais_recente = max(arquivos, key=lambda p: Path(p).stat().st_mtime)
    print(f"Usando o arquivo mais recente para '{Path(padrao_str).name}': {Path(mais_recente).name}")
    return read_table(mais_recente, aba)


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
    """Mantém só os dígitos do CNPJ e preenche com zero à esquerda até 14
    dígitos (equivalente a =TEXTO(CNPJ;"00000000000000") do Excel) — o Excel
    costuma perder o zero à esquerda de CNPJ guardado como número.
    """
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ""
    digitos = re.sub(r"\D", "", str(valor))
    return digitos.zfill(14) if digitos else ""


def normalize_cnpj_raiz(valor) -> str:
    """Mantém só os dígitos e preenche com zero à esquerda até 8 dígitos —
    a raiz do CNPJ (os 8 primeiros dígitos, sem filial/dígitos verificadores).
    Não reaproveita normalize_cnpj porque ali o preenchimento é até 14
    dígitos (CNPJ completo), não 8.
    """
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ""
    digitos = re.sub(r"\D", "", str(valor))
    return digitos.zfill(8) if digitos else ""


def normalize_ean(valor) -> str:
    """Remove aspas simples (marcador de "número como texto" do Excel/CSV,
    que pode vir só no início ou nas duas pontas) e espaços nas pontas.
    """
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ""
    return str(valor).strip().strip("'")


def to_datetime(series: pd.Series) -> pd.Series:
    """Converte para data aceitando formatos mistos: aaaa-mm-dd (ISO, sem
    ambiguidade) e dd/mm/aaaa (padrão BR, dayfirst=True).

    Usar dayfirst=True direto com format="mixed" faz o pandas inverter dia/mês
    também em timestamps ISO (bug observado: "2026-05-10" virava 10/10 em vez
    de 10/05), então cada formato é tratado separadamente.

    As duas partes são combinadas com pd.concat (não com atribuição num
    Series datetime64[ns] pré-criado) porque forçar dtype "ns" de antemão
    estoura (OutOfBoundsDatetime) em datas tipo "31/12/9999" usadas como
    "sem prazo definido" — o pandas sabe lidar com essas datas usando uma
    resolução mais larga (us/s), só não quando o dtype de destino já está
    travado em nanossegundos.
    """
    texto = series.astype(str).str.strip()
    eh_iso = texto.str.match(r"^\d{4}-\d{2}-\d{2}")

    partes = []
    if eh_iso.any():
        partes.append(pd.to_datetime(texto[eh_iso], errors="coerce"))
    if (~eh_iso).any():
        partes.append(pd.to_datetime(texto[~eh_iso], errors="coerce", dayfirst=True))

    if not partes:
        return pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")

    return pd.concat(partes).reindex(series.index)


def to_numeric(series: pd.Series) -> pd.Series:
    """Converte texto numérico para float, aceitando dois formatos:

    - pt-BR (texto digitado numa célula/CSV): "1.234,56" ou "56,87"
      -> "." é separador de milhar, "," é decimal.
    - internacional (célula Excel com número nativo, convertida pra string
      pelo pandas): "0.3", "1234.56" -> "." já é decimal, sem vírgula.

    Só remove o "." como milhar quando a string também tem vírgula — senão
    "0.3" vira "03" (=3) por engano, que é o formato que sai quando o Excel
    guarda um float nativo (ex.: célula formatada como %).
    """
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")

    texto = series.astype(str).str.strip().str.replace(r"^'", "", regex=True)
    tem_virgula = texto.str.contains(",", regex=False)

    limpo = texto.copy()
    limpo.loc[tem_virgula] = (
        limpo.loc[tem_virgula].str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    )
    return pd.to_numeric(limpo, errors="coerce")
