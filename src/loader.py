import unicodedata
from pathlib import Path

import pandas as pd

# ------------------------------------------
#  RUTAS DE ENTRADA
# ------------------------------------------
# Se calculan desde la raiz del proyecto (un nivel arriba de src/), no con la
# ruta absoluta de X:. Asi la carpeta se puede mover o renombrar sin que el
# proyecto deje de encontrar sus insumos.
BASE_DIR = Path(__file__).resolve().parents[1]
INPUT_DIR = BASE_DIR / "data" / "input"

ruta_clientes = INPUT_DIR / "base_cliente.xlsx"
ruta_contratos = INPUT_DIR / "base_contratos.xlsx"
ruta_inpc = INPUT_DIR / "base_inpc.xlsx"
ruta_coincidencia = INPUT_DIR / "1ra Parte" / "coincidencia.xlsx"


# ------------------------------------------
#  UTILIDADES
# ------------------------------------------
def normalizar_columnas(df):
    # Limpia encabezados: trim, minúsculas, guiones bajos y sin acentos
    df.columns = (
        df.columns
        .str.strip()
        .str.lower()
        .str.replace(" ", "_")
        .str.normalize("NFKD")
        .str.encode("ascii", errors="ignore")
        .str.decode("utf-8")
    )
    return df


def convertir_fechas(df):
    for col in df.columns:
        if "fecha" in col or "fec" in col:
            # Coerce a datetime para evitar errores posteriores en merges/cálculos
            df[col] = pd.to_datetime(df[col], errors="coerce")
    return df


def _norm_txt(val) -> str:
    """Normaliza texto para cruces: trim, upper, sin acentos, espacios colapsados."""
    if pd.isna(val):
        return ""
    s = str(val).strip().upper()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = " ".join(s.split())
    return s


def anexar_orden_a_clientes(df_clientes: pd.DataFrame, df_coincidencia: pd.DataFrame) -> pd.DataFrame:
    """
    Agrega columna `orden` a la base de clientes usando coincidencia.xlsx.
    Cruce por RFC, CLIENTE, PLAZA y SUCURSAL (normalizados).
    """
    if df_clientes is None or df_clientes.empty:
        return df_clientes
    if df_coincidencia is None or df_coincidencia.empty:
        # Mantener compatibilidad aunque no exista coincidencia.xlsx
        if "orden" not in df_clientes.columns:
            df_clientes = df_clientes.copy()
            df_clientes["orden"] = pd.NA
        return df_clientes

    clientes = df_clientes.copy()
    coinc = df_coincidencia.copy()

    for col in ["rfc", "cliente", "plaza", "sucursal"]:
        if col not in clientes.columns:
            clientes[col] = ""
        if col not in coinc.columns:
            coinc[col] = ""

    coinc["orden"] = pd.to_numeric(coinc.get("orden"), errors="coerce")
    coinc = coinc.dropna(subset=["orden"])

    key_cols = []
    for col in ["rfc", "cliente", "plaza", "sucursal"]:
        k = f"{col}_key"
        clientes[k] = clientes[col].map(_norm_txt)
        coinc[k] = coinc[col].map(_norm_txt)
        key_cols.append(k)

    coinc = coinc.drop_duplicates(subset=key_cols, keep="first")
    clientes = clientes.merge(coinc[key_cols + ["orden"]], on=key_cols, how="left")
    clientes = clientes.drop(columns=key_cols)

    # Orden entero (si aplica)
    try:
        clientes["orden"] = pd.to_numeric(clientes["orden"], errors="coerce").astype("Int64")
    except Exception:
        # Si algo raro viene en la columna, se deja como está.
        pass

    return clientes


def completar_rfc_contratos(df_contratos: pd.DataFrame, df_coincidencia: pd.DataFrame) -> pd.DataFrame:
    """
    Rellena RFC faltante en contratos usando coincidencia.xlsx (clave oficial: orden).
    No modifica registros con RFC ya presente.
    """
    if df_contratos is None or df_contratos.empty:
        return df_contratos
    if df_coincidencia is None or df_coincidencia.empty:
        return df_contratos
    if "orden" not in df_contratos.columns or "orden" not in df_coincidencia.columns:
        return df_contratos

    contratos = df_contratos.copy()
    coincidencia = df_coincidencia.copy()

    contratos["orden_key"] = pd.to_numeric(contratos["orden"], errors="coerce")
    coincidencia["orden_key"] = pd.to_numeric(coincidencia["orden"], errors="coerce")
    coincidencia = coincidencia.dropna(subset=["orden_key"])

    rfc_col = "rfc_del_subarrendatario" if "rfc_del_subarrendatario" in contratos.columns else "rfc"
    if rfc_col not in contratos.columns:
        return df_contratos

    mapa_rfc = (
        coincidencia[["orden_key", "rfc"]]
        .copy()
        .assign(rfc=lambda d: d["rfc"].astype(str).str.strip())
        .dropna(subset=["orden_key"])
        .drop_duplicates(subset=["orden_key"], keep="first")
        .set_index("orden_key")["rfc"]
    )

    contratos = contratos.merge(
        mapa_rfc.rename("rfc_from_coinc"),
        left_on="orden_key",
        right_index=True,
        how="left",
    )

    def _needs_fill(val):
        return pd.isna(val) or str(val).strip() == ""

    mask_fill = contratos[rfc_col].apply(_needs_fill)
    contratos.loc[mask_fill, rfc_col] = contratos.loc[mask_fill, "rfc_from_coinc"]
    contratos = contratos.drop(columns=["orden_key", "rfc_from_coinc"])
    return contratos


# ------------------------------------------
#  FUNCIONES DE CARGA
# ------------------------------------------
def load_clientes():
    print("Cargando base de clientes...")
    df = pd.read_excel(ruta_clientes)
    df = normalizar_columnas(df)
    # Estandarizar nombre de columna importe_mtto (quitar punto final si existe)
    if "importe_mtto." in df.columns and "importe_mtto" not in df.columns:
        df = df.rename(columns={"importe_mtto.": "importe_mtto"})
    df = convertir_fechas(df)
    # Solo usar registros con ORDEN numerico.
    if "orden" in df.columns:
        total = len(df)
        df["orden"] = pd.to_numeric(df["orden"], errors="coerce")
        df = df.dropna(subset=["orden"]).copy()
        try:
            df["orden"] = df["orden"].astype("Int64")
        except Exception:
            pass
        omitidas = total - len(df)
        if omitidas:
            print(f"Aviso: {omitidas} filas de base_cliente omitidas por ORDEN no numerico.")
    return df


def load_contratos():
    print("Cargando base de contratos...")
    df = pd.read_excel(ruta_contratos)
    df = normalizar_columnas(df)
    df = convertir_fechas(df)
    # Solo usar registros con ORDEN numerico.
    if "orden" in df.columns:
        total = len(df)
        df["orden"] = pd.to_numeric(df["orden"], errors="coerce")
        df = df.dropna(subset=["orden"]).copy()
        try:
            df["orden"] = df["orden"].astype("Int64")
        except Exception:
            pass
        omitidas = total - len(df)
        if omitidas:
            print(f"Aviso: {omitidas} filas de base_contratos omitidas por ORDEN no numerico.")
    return df


def load_inpc():
    print("Cargando base INPC...")
    df = pd.read_excel(ruta_inpc)
    df = normalizar_columnas(df)

    # Asegurar formato de mes y año
    if "mes" in df.columns:
        df["mes"] = df["mes"].astype(str).str.zfill(2)

    if "anio" in df.columns:
        df["anio"] = df["anio"].astype(int)

    return df


def load_coincidencia():
    print("Cargando coincidencia (RFC/CLIENTE/PLAZA/SUCURSAL -> ORDEN)...")
    df = pd.read_excel(ruta_coincidencia)
    df = normalizar_columnas(df)
    return df


# ------------------------------------------
#  FUNCIÓN GENERAL
# ------------------------------------------
def load_all():
    clientes = load_clientes()
    contratos = load_contratos()
    inpc = load_inpc()



    print("Bases cargadas correctamente.")
    return clientes, contratos, inpc


