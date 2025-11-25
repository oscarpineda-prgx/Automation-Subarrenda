import pandas as pd

# ------------------------------------------
#  RUTAS ABSOLUTAS EN CITRIX
# ------------------------------------------
ruta_clientes = r'\\amer.prgx.com\Citrix\UserHomeDir\opined01\Desktop\Automation-Subarrenda\data\input\base_cliente.xlsx'
ruta_contratos = r'\\amer.prgx.com\Citrix\UserHomeDir\opined01\Desktop\Automation-Subarrenda\data\input\base_contratos.xlsx'
ruta_inpc = r'\\amer.prgx.com\Citrix\UserHomeDir\opined01\Desktop\Automation-Subarrenda\data\input\base_inpc.xlsx'


# ------------------------------------------
#  UTILIDADES
# ------------------------------------------
def normalizar_columnas(df):
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
            df[col] = pd.to_datetime(df[col], errors="coerce")
    return df


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


# ------------------------------------------
#  FUNCIÓN GENERAL
# ------------------------------------------
def load_all():
    clientes = load_clientes()
    contratos = load_contratos()
    inpc = load_inpc()

    print("Bases cargadas correctamente.")
    return clientes, contratos, inpc
