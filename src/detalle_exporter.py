import os
from typing import List, Optional, Tuple

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

from src.format_excel import format_workbook

# Meses en texto (enero=1, ...)
_MESES = [
    "ENERO",
    "FEBRERO",
    "MARZO",
    "ABRIL",
    "MAYO",
    "JUNIO",
    "JULIO",
    "AGOSTO",
    "SEPTIEMBRE",
    "OCTUBRE",
    "NOVIEMBRE",
    "DICIEMBRE",
]
_MES_MAP = {m: i + 1 for i, m in enumerate(_MESES)}


def _sanitizar_nombre_archivo(texto: str) -> str:
    """Limpia texto para usarlo como nombre de archivo en Windows."""
    if texto is None:
        return "sin_nombre"
    nombre = str(texto)
    for ch in '<>:"/\\|?*':
        nombre = nombre.replace(ch, "")
    nombre = nombre.strip().replace(" ", "_")
    return nombre or "sin_nombre"


def _mes_a_texto(num: Optional[int]) -> Optional[str]:
    if num is None or not (1 <= num <= 12):
        return None
    return _MESES[num - 1]


def _coerce_mes(val) -> Optional[int]:
    if pd.isna(val):
        return None
    if isinstance(val, (int, float)):
        num = int(val)
        return num if 1 <= num <= 12 else None
    if isinstance(val, str):
        txt = val.strip().upper()
        if txt.isdigit():
            num = int(txt)
            return num if 1 <= num <= 12 else None
        if txt in _MES_MAP:
            return _MES_MAP[txt]
        for nombre, num in _MES_MAP.items():
            if nombre.startswith(txt[:3]):
                return num
    return None


def _buscar_mes_fda(df_clientes: Optional[pd.DataFrame], rfc: str, cliente: str) -> Tuple[Optional[int], Optional[str]]:
    """Busca el mes de incremento FDA en base_clientes por RFC (y cliente si existe). Usa solo columna mes3."""
    if df_clientes is None or df_clientes.empty:
        return None, None

    df = df_clientes.copy()
    df["rfc_key"] = df["rfc"].astype(str).str.strip().str.upper() if "rfc" in df.columns else ""
    df["cliente_key"] = ""
    if "cliente" in df.columns:
        df["cliente_key"] = df["cliente"].astype(str).str.strip().str.upper()
    elif "nombre_del_subarrendatario" in df.columns:
        df["cliente_key"] = df["nombre_del_subarrendatario"].astype(str).str.strip().str.upper()

    clave_rfc = str(rfc).strip().upper()
    clave_cli = str(cliente).strip().upper()

    filtrado = df[df["rfc_key"] == clave_rfc]
    if not filtrado.empty and filtrado["cliente_key"].notna().any():
        prefer_cli = filtrado[filtrado["cliente_key"] == clave_cli]
        if not prefer_cli.empty:
            filtrado = prefer_cli

    if filtrado.empty or "mes3" not in filtrado.columns:
        return None, None

    for val in filtrado["mes3"]:
        mes_num = _coerce_mes(val)
        if mes_num:
            return mes_num, _mes_a_texto(mes_num)
    return None, None


def _buscar_mes_auditoria(df_contratos: Optional[pd.DataFrame], rfc: str, cliente: str) -> Tuple[Optional[int], Optional[str]]:
    """Toma el mes de la fecha de firma del contrato (auditoria) por RFC + cliente."""
    if df_contratos is None or df_contratos.empty:
        return None, None

    df = df_contratos.copy()
    df["rfc_key"] = df["rfc_del_subarrendatario"].astype(str).str.strip().str.upper() if "rfc_del_subarrendatario" in df.columns else ""
    df["cliente_key"] = ""
    if "nombre_del_subarrendatario" in df.columns:
        df["cliente_key"] = df["nombre_del_subarrendatario"].astype(str).str.strip().str.upper()

    clave_rfc = str(rfc).strip().upper()
    clave_cli = str(cliente).strip().upper()

    filtrado = df[df["rfc_key"] == clave_rfc]
    prefer_cli = filtrado[filtrado["cliente_key"] == clave_cli]
    if not prefer_cli.empty:
        filtrado = prefer_cli

    if filtrado.empty or "fecha_de_firma_del_contrato" not in filtrado.columns:
        return None, None

    fechas = pd.to_datetime(filtrado["fecha_de_firma_del_contrato"], errors="coerce").dropna()
    if fechas.empty:
        return None, None
    mes_num = int(fechas.iloc[0].month)
    return mes_num, _mes_a_texto(mes_num)


def _agregar_tabla_incrementos(
    ruta: str,
    mes_fda_num: Optional[int],
    mes_fda_txt: Optional[str],
    mes_aud_num: Optional[int],
    mes_aud_txt: Optional[str],
    diferencia: Optional[int],
) -> None:
    """Agrega tabla pequeña con meses de incremento en la misma hoja (a la derecha)."""
    wb = load_workbook(ruta)
    ws = wb.active

    start_col = ws.max_column + 2  # un espacio en blanco
    start_row = 2  # debajo de encabezados

    fill = PatternFill("solid", fgColor="1F4E78")
    font = Font(color="FFFFFF", bold=True)
    center = Alignment(horizontal="center", vertical="center")

    filas = [
        ("MES DE INCREMENTO FDA", mes_fda_txt, mes_fda_num),
        ("MES DE INCREMENTO AUDITORIA", mes_aud_txt, mes_aud_num),
        ("Diferencia", None, diferencia),
    ]

    for offset, (label, val_txt, val_num) in enumerate(filas):
        row = start_row + offset
        c_label = ws.cell(row=row, column=start_col, value=label)
        c_label.fill = fill
        c_label.font = font
        c_label.alignment = center

        c_txt = ws.cell(row=row, column=start_col + 1, value=val_txt if val_txt is not None else "")
        c_num = ws.cell(row=row, column=start_col + 2, value=val_num if val_num is not None else "")
        c_txt.alignment = center
        c_num.alignment = center

    wb.save(ruta)


def exportar_detalles_individuales(
    detalle: pd.DataFrame,
    df_clientes: Optional[pd.DataFrame] = None,
    df_contratos: Optional[pd.DataFrame] = None,
    output_dir: str = "data/output/detalle_individual",
    aplicar_formato: bool = True,
) -> List[str]:
    """
    Genera un archivo Excel por cada subarrendatario (clave RFC + cliente) y agrega
    una tabla con meses de incremento (FDA vs auditoria).
    """
    if detalle.empty:
        print("Detalle vacio: no se generan archivos individuales.")
        return []
    if not {"rfc", "cliente"}.issubset(detalle.columns):
        raise ValueError("Detalle no contiene columnas requeridas: rfc y cliente.")

    os.makedirs(output_dir, exist_ok=True)

    base = detalle.copy()
    base["rfc"] = base["rfc"].astype(str).str.strip()
    base["cliente"] = base["cliente"].astype(str).str.strip()
    base = base[(base["rfc"] != "") & (base["cliente"] != "")]
    if base.empty:
        print("No hay registros con RFC y cliente para exportar.")
        return []

    rutas = []
    for (rfc, cliente), df_grp in base.groupby(["rfc", "cliente"]):
        nombre_archivo = f"{_sanitizar_nombre_archivo(rfc)}_{_sanitizar_nombre_archivo(cliente)}.xlsx"
        ruta = os.path.join(output_dir, nombre_archivo)
        df_sorted = df_grp.sort_values("fecha") if "fecha" in df_grp.columns else df_grp
        df_sorted.to_excel(ruta, index=False)

        if aplicar_formato:
            try:
                format_workbook(ruta)
            except Exception as exc:
                print(f"No se pudo formatear {ruta}: {exc}")

        rutas.append(ruta)

    print(f"Archivos individuales generados: {len(rutas)} en {output_dir}")
    return rutas
