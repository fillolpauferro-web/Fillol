# -*- coding: utf-8 -*-
"""
Converte os relatórios de Ressarcimento (Excel/CSV) em Parquet e roda a
análise de desconto líquido / Margem OL de Santa Cruz e Panpharma.

Uso (deixe este arquivo na mesma pasta de analise_desconto_liquido.py):
    pip install pandas pyarrow openpyxl python-calamine
    python converter_e_analisar.py

Opções:
    --origem   pasta com os .xlsx/.xls/.csv   (padrão: ...\\Desktop\\Ressarcimento)
    --destino  pasta onde salvar os .parquet  (padrão: ...\\Ressarcimento\\Parquet)
    --refazer  reconverte mesmo se o .parquet já existir
    --so-converter  não roda a análise no final
"""

import argparse
import sys
import time
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
import analise_desconto_liquido as analise  # noqa: E402

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
        if any(analise.norm(v) == "distribuidor" for v in previa.iloc[i].tolist()):
            return i
    return 0


def ler_relatorio(arq: Path) -> pd.DataFrame:
    if arq.suffix.lower() == ".csv":
        for sep, enc in ((";", "utf-8-sig"), (";", "latin-1"), (",", "utf-8-sig")):
            try:
                previa = pd.read_csv(arq, sep=sep, encoding=enc, header=None, nrows=10, dtype=str)
                if previa.shape[1] > 5:
                    break
            except Exception:  # noqa: BLE001
                continue
        h = _linha_cabecalho(previa)
        df = pd.read_csv(arq, sep=sep, encoding=enc, skiprows=h, dtype=str, decimal=",")
    else:
        motor = _motor_excel(arq)
        previa = pd.read_excel(arq, header=None, nrows=10, engine=motor)
        h = _linha_cabecalho(previa)
        df = pd.read_excel(arq, header=h, engine=motor)

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
                      and not p.name.startswith("Analise_Desconto"))
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
    args = ap.parse_args()

    origem, destino = Path(args.origem), Path(args.destino)
    converter(origem, destino, args.refazer)
    if args.so_converter:
        return

    print("\n" + "#" * 78 + "\nANÁLISE\n" + "#" * 78)
    sys.argv = [sys.argv[0], "--pasta", str(destino),
                "--saida", str(origem / "Analise_Desconto_SantaCruz_Panpharma.xlsx")]
    analise.main()


if __name__ == "__main__":
    main()
