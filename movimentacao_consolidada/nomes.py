"""Resolve um nome único de cliente por CNPJ raiz. O mesmo CNPJ pode
aparecer no export do SAP com grafias diferentes de Client Name (ex.:
espaçamento/abreviação distintos) -- aqui a gente escolhe a grafia mais
frequente por CNPJ e aplica client_name_overrides.csv por cima pra forçar
casos específicos.
"""
import pandas as pd

from config import CLIENT_NAME_OVERRIDES_FILE


def load_overrides() -> dict:
    try:
        overrides = pd.read_csv(CLIENT_NAME_OVERRIDES_FILE, comment="#", dtype=str, encoding="utf-8")
    except UnicodeDecodeError:
        overrides = pd.read_csv(CLIENT_NAME_OVERRIDES_FILE, comment="#", dtype=str, encoding="cp1252")
    if overrides.empty:
        return {}
    overrides["cnpj_raiz"] = overrides["cnpj_raiz"].str.strip().str.zfill(8)
    return dict(zip(overrides["cnpj_raiz"], overrides["cliente"].str.strip()))


def _nome_mais_frequente(df: pd.DataFrame) -> dict:
    contagem = (
        df.groupby(["cnpj_raiz", "cliente"])
        .size()
        .reset_index(name="n")
        .sort_values(["cnpj_raiz", "n", "cliente"], ascending=[True, False, True])
    )
    return contagem.drop_duplicates("cnpj_raiz").set_index("cnpj_raiz")["cliente"].to_dict()


def normalizar_nomes(df: pd.DataFrame, avisar: bool = True) -> pd.DataFrame:
    """Sobrescreve df['cliente'] com um nome único por cnpj_raiz."""
    canonico = _nome_mais_frequente(df)
    overrides = load_overrides()

    if avisar:
        variacoes = df.groupby("cnpj_raiz")["cliente"].nunique()
        multiplos = variacoes[variacoes > 1]
        if not multiplos.empty:
            print(
                f"Aviso: {len(multiplos)} CNPJ(s) com mais de uma grafia de "
                "Client Name -- consolidando num nome único:"
            )
            for cnpj in multiplos.index:
                nomes = sorted(df.loc[df["cnpj_raiz"] == cnpj, "cliente"].unique())
                escolhido = overrides.get(cnpj, canonico.get(cnpj))
                print(f"  {cnpj}: {nomes} -> \"{escolhido}\"")

    canonico.update(overrides)
    df = df.copy()
    df["cliente"] = df["cnpj_raiz"].map(canonico).fillna(df["cliente"])
    return df


def aplicar_overrides(df: pd.DataFrame) -> pd.DataFrame:
    """Só aplica client_name_overrides.csv, sem lógica de nome mais
    frequente -- usado em arquivos pequenos (ex.: Saldo Inicial) onde não
    faz sentido calcular frequência."""
    overrides = load_overrides()
    if not overrides:
        return df
    df = df.copy()
    df["cliente"] = df["cnpj_raiz"].map(overrides).fillna(df["cliente"])
    return df
