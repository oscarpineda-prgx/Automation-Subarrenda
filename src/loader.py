import pandas as pd
import unicodedata

# ------------------------------------------
#  RUTAS ABSOLUTAS EN CITRIX
# ------------------------------------------
ruta_clientes = r'\\amer.prgx.com\Citrix\UserHomeDir\opined01\Desktop\Automation-Subarrenda\data\input\base_cliente.xlsx'
ruta_contratos = r'\\amer.prgx.com\Citrix\UserHomeDir\opined01\Desktop\Automation-Subarrenda\data\input\base_contratos.xlsx'
ruta_inpc = r'\\amer.prgx.com\Citrix\UserHomeDir\opined01\Desktop\Automation-Subarrenda\data\input\base_inpc.xlsx'
ruta_coincidencia = r'\\amer.prgx.com\Citrix\UserHomeDir\opined01\Desktop\Automation-Subarrenda\data\input\coincidencia.xlsx'


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
    return df


def load_contratos():
    print("Cargando base de contratos...")
    df = pd.read_excel(ruta_contratos)
    df = normalizar_columnas(df)
    df = convertir_fechas(df)
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

    try:
        coincidencia = load_coincidencia()
        clientes = anexar_orden_a_clientes(clientes, coincidencia)
        if "orden" in clientes.columns:
            faltantes = int(clientes["orden"].isna().sum())
            if faltantes:
                print(f"Aviso: {faltantes} filas en base_cliente sin ORDEN (sin match en coincidencia.xlsx).")
    except FileNotFoundError:
        print("coincidencia.xlsx no encontrado: se continúa sin ORDEN en base_cliente.")
    except Exception as exc:
        print(f"No se pudo cargar/anexar coincidencia.xlsx: {exc}")

    print("Bases cargadas correctamente.")
    return clientes, contratos, inpc
