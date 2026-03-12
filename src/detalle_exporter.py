import os
from typing import List, Optional, Tuple

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.drawing.image import Image

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
_ACCOUNTING_FMT = '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'


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
    # Normaliza mes en texto/numero a entero 1-12
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


def _buscar_mes_fda(df_clientes: Optional[pd.DataFrame], rfc: str, subarrendatario: str) -> Tuple[Optional[int], Optional[str]]:
    """Fallback: Busca el mes de incremento FDA en base_clientes por RFC (y cliente si existe). Usa columna mes3."""
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
    clave_cli = str(subarrendatario).strip().upper()

    filtrado = df[df["rfc_key"] == clave_rfc]
    # Si hay mismo RFC con clientes distintos, prioriza match exacto de cliente
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


def _buscar_mes_fda_por_orden(
    df_clientes: Optional[pd.DataFrame],
    orden,
    rfc: str,
    subarrendatario: str,
) -> Tuple[Optional[int], Optional[str]]:
    """Busca mes FDA por orden si está disponible; si no hay match, hace fallback al método anterior."""
    if df_clientes is None or df_clientes.empty:
        return None, None
    df = df_clientes.copy()
    if "orden" not in df.columns:
        return _buscar_mes_fda(df_clientes, rfc, subarrendatario)

    ord_key = pd.to_numeric(orden, errors="coerce")
    df["orden_key"] = pd.to_numeric(df["orden"], errors="coerce")
    filtrado = df[df["orden_key"] == ord_key] if pd.notna(ord_key) else pd.DataFrame()
    if filtrado.empty:
        return _buscar_mes_fda(df_clientes, rfc, subarrendatario)

    if "mes3" not in filtrado.columns:
        return None, None
    for val in filtrado["mes3"]:
        mes_num = _coerce_mes(val)
        if mes_num:
            return mes_num, _mes_a_texto(mes_num)
    return None, None


def _buscar_fecha_firma(df_contratos: Optional[pd.DataFrame], rfc: str, subarrendatario: str) -> Optional[pd.Timestamp]:
    """Devuelve la fecha de firma del contrato (auditoria) por RFC + subarrendatario (fallback)."""
    if df_contratos is None or df_contratos.empty:
        return None

    df = df_contratos.copy()
    df["rfc_key"] = df["rfc_del_subarrendatario"].astype(str).str.strip().str.upper() if "rfc_del_subarrendatario" in df.columns else ""
    df["cliente_key"] = ""
    if "nombre_del_subarrendatario" in df.columns:
        df["cliente_key"] = df["nombre_del_subarrendatario"].astype(str).str.strip().str.upper()

    clave_rfc = str(rfc).strip().upper()
    clave_cli = str(subarrendatario).strip().upper()

    filtrado = df[df["rfc_key"] == clave_rfc]
    prefer_cli = filtrado[filtrado["cliente_key"] == clave_cli]
    if not prefer_cli.empty:
        filtrado = prefer_cli

    if filtrado.empty or "fecha_de_firma_del_contrato" not in filtrado.columns:
        return None

    fechas = pd.to_datetime(filtrado["fecha_de_firma_del_contrato"], errors="coerce").dropna()
    if fechas.empty:
        return None
    return fechas.iloc[0]


def _buscar_fecha_firma_por_orden(
    df_contratos: Optional[pd.DataFrame],
    orden,
    rfc: str,
    subarrendatario: str,
) -> Optional[pd.Timestamp]:
    """Devuelve la fecha de firma del contrato (auditoria) por orden; fallback RFC/cliente."""
    if df_contratos is None or df_contratos.empty:
        return None
    df = df_contratos.copy()
    if "orden" not in df.columns:
        return _buscar_fecha_firma(df_contratos, rfc, subarrendatario)

    ord_key = pd.to_numeric(orden, errors="coerce")
    df["orden_key"] = pd.to_numeric(df["orden"], errors="coerce")
    filtrado = df[df["orden_key"] == ord_key] if pd.notna(ord_key) else pd.DataFrame()
    if filtrado.empty:
        return _buscar_fecha_firma(df_contratos, rfc, subarrendatario)

    if "fecha_de_firma_del_contrato" not in filtrado.columns:
        return None
    fechas = pd.to_datetime(filtrado["fecha_de_firma_del_contrato"], errors="coerce").dropna()
    if fechas.empty:
        return None
    return fechas.iloc[0]


def _buscar_mes_auditoria(df_contratos: Optional[pd.DataFrame], rfc: str, subarrendatario: str) -> Tuple[Optional[int], Optional[str]]:
    """Toma el mes de la fecha de firma del contrato (auditoria) por RFC + subarrendatario (fallback)."""
    fecha = _buscar_fecha_firma(df_contratos, rfc, subarrendatario)
    if fecha is None:
        return None, None
    mes_num = int(pd.to_datetime(fecha).month)
    return mes_num, _mes_a_texto(mes_num)


def _buscar_mes_auditoria_por_orden(
    df_contratos: Optional[pd.DataFrame],
    orden,
    rfc: str,
    subarrendatario: str,
) -> Tuple[Optional[int], Optional[str]]:
    """Toma el mes de la fecha de firma del contrato (auditoria) por orden; fallback RFC/cliente."""
    fecha = _buscar_fecha_firma_por_orden(df_contratos, orden, rfc, subarrendatario)
    if fecha is None:
        return None, None
    mes_num = int(pd.to_datetime(fecha).month)
    return mes_num, _mes_a_texto(mes_num)


def _buscar_mes_auditoria_desde_detalle(df_detalle: pd.DataFrame) -> Tuple[Optional[int], Optional[str]]:
    """
    Toma el mes de incremento de auditoria desde la columna mes_incremento del detalle.
    """
    if df_detalle is None or df_detalle.empty or "mes_incremento" not in df_detalle.columns:
        return None, None
    for val in df_detalle["mes_incremento"]:
        mes_num = _coerce_mes(val)
        if mes_num:
            return mes_num, _mes_a_texto(mes_num)
    return None, None


def _buscar_fecha_primer_cobro(df_detalle: pd.DataFrame) -> Optional[pd.Timestamp]:
    """Devuelve la primera fecha con importe_renta_c > 0 dentro del detalle individual."""
    if "importe_renta_c" not in df_detalle.columns or "fecha" not in df_detalle.columns:
        return None

    df = df_detalle.copy()
    df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce")
    df = df.dropna(subset=["fecha"])
    df = df.sort_values("fecha")

    cobros = df[pd.to_numeric(df["importe_renta_c"], errors="coerce").fillna(0) > 0]
    if cobros.empty:
        return None
    return cobros.iloc[0]["fecha"]


def _diff_meses(fecha_a: Optional[pd.Timestamp], fecha_b: Optional[pd.Timestamp]) -> Optional[int]:
    """Devuelve diferencia absoluta en meses entre dos fechas (ajusta si el día final es menor)."""
    if fecha_a is None or fecha_b is None:
        return None
    f1 = pd.to_datetime(fecha_a, errors="coerce")
    f2 = pd.to_datetime(fecha_b, errors="coerce")
    if pd.isna(f1) or pd.isna(f2):
        return None
    meses = (f2.year - f1.year) * 12 + (f2.month - f1.month)
    if f2.day < f1.day:
        meses -= 1
    return abs(int(meses))

def _calcular_total_diferencia(df_detalle: pd.DataFrame) -> Optional[float]:
    """
    Suma diferencia_base_vs_aud desde la primera hasta la ultima fecha con cobro (>0).
    Si no hay fecha valida, cae al comportamiento previo (solo suma filas con cobro).
    """
    req_cols = {"importe_renta_c", "diferencia_base_vs_aud"}
    if not req_cols.issubset(df_detalle.columns):
        return None

    renta = pd.to_numeric(df_detalle["importe_renta_c"], errors="coerce").fillna(0)
    diferencia = pd.to_numeric(df_detalle["diferencia_base_vs_aud"], errors="coerce").fillna(0)
    cobros_mask = renta > 0
    if not cobros_mask.any():
        return 0.0

    if "fecha" not in df_detalle.columns:
        return float(diferencia[cobros_mask].sum())

    fechas = pd.to_datetime(df_detalle["fecha"], errors="coerce")
    cobros_validos = cobros_mask & fechas.notna()
    if not cobros_validos.any():
        return float(diferencia[cobros_mask].sum())

    fecha_inicio = fechas[cobros_validos].min()
    fecha_fin = fechas[cobros_validos].max()
    en_rango = fechas.between(fecha_inicio, fecha_fin)
    return float(diferencia[en_rango.fillna(False)].sum())


def _desglosar_importe_iva(total_diferencia: Optional[float]) -> Tuple[Optional[float], Optional[float]]:
    """
    Desglosa total en:
    - importe (base): total / 1.16
    - iva: total - importe
    """
    if total_diferencia is None:
        return None, None
    total = round(float(total_diferencia), 2)
    importe = round(total / 1.16, 2)
    iva = round(total - importe, 2)
    return importe, iva


def _inserta_encabezado_presentacion(
    ruta: str,
    rfc: str,
    subarrendatario: str,
    total_diferencia: Optional[float],
    orden: Optional[int] = None,
    logo_path: str = "data/input/Picture1.png",
) -> None:
    """Inserta logo, titulos y total de diferencia en la parte superior de la hoja."""
    wb = load_workbook(ruta)
    ws = wb.active

    # Reservar espacio arriba del encabezado original para el bloque de titulo
    espacio = 8
    ws.insert_rows(1, espacio)

    # Logo a la izquierda
    try:
        if os.path.isfile(logo_path):
            img = Image(logo_path)
            img.width = 180
            img.height = 80
            ws.add_image(img, "B2")
        else:
            print(f"Logo no encontrado en {logo_path}")
    except Exception as exc:
        print(f"No se pudo insertar logo en {ruta}: {exc}")

    titulo = ws.cell(row=3, column=6, value="DETALLE DIFERENCIA DE COBRO SUBARRENDAMIENTOS")
    titulo.font = Font(bold=True, size=14)
    titulo.alignment = Alignment(horizontal="center")
    ws.merge_cells(start_row=3, start_column=6, end_row=3, end_column=13)

    if orden is not None and not pd.isna(orden):
        subtitulo_txt = f"ORDEN {int(orden)} - {str(rfc).strip()} - {str(subarrendatario).strip()}"
    else:
        subtitulo_txt = f"{str(rfc).strip()} - {str(subarrendatario).strip()}"
    subtitulo = ws.cell(row=4, column=6, value=subtitulo_txt)
    subtitulo.font = Font(bold=True, size=12)
    subtitulo.alignment = Alignment(horizontal="center")
    ws.merge_cells(start_row=4, start_column=6, end_row=4, end_column=13)

    etiqueta_total = ws.cell(row=6, column=14, value="Total Diferencia base vs auditoria")
    etiqueta_total.font = Font(bold=True)
    etiqueta_total.alignment = Alignment(horizontal="center")
    ws.merge_cells(start_row=6, start_column=14, end_row=6, end_column=16)

    # Desglose IVA / IMPORTE (base) a la izquierda del total
    etiqueta_iva = ws.cell(row=6, column=6, value="IVA")
    etiqueta_iva.font = Font(bold=True)
    etiqueta_iva.alignment = Alignment(horizontal="center")
    etiqueta_importe = ws.cell(row=7, column=6, value="IMPORTE")
    etiqueta_importe.font = Font(bold=True)
    etiqueta_importe.alignment = Alignment(horizontal="center")

    importe_base, iva = _desglosar_importe_iva(total_diferencia)
    celda_iva = ws.cell(row=6, column=7, value="" if iva is None else iva)
    celda_importe = ws.cell(row=7, column=7, value="" if importe_base is None else importe_base)
    celda_iva.alignment = Alignment(horizontal="center")
    celda_importe.alignment = Alignment(horizontal="center")
    if total_diferencia is not None:
        celda_iva.number_format = _ACCOUNTING_FMT
        celda_importe.number_format = _ACCOUNTING_FMT

    valor_total = "" if total_diferencia is None else round(float(total_diferencia), 2)
    celda_total = ws.cell(row=6, column=17, value=valor_total)
    celda_total.alignment = Alignment(horizontal="center")
    if total_diferencia is not None:
        celda_total.number_format = _ACCOUNTING_FMT

    wb.save(ruta)


def _agregar_tabla_incrementos(
    ruta: str,
    mes_fda_num: Optional[int],
    mes_fda_txt: Optional[str],
    mes_aud_num: Optional[int],
    mes_aud_txt: Optional[str],
    diferencia: Optional[int],
) -> int:
    """Agrega tabla pequeña con meses de incremento en la misma hoja (a la derecha)."""
    wb = load_workbook(ruta)
    ws = wb.active

    # Ubica la tabla a la derecha del detalle, dejando una columna en blanco
    start_col = ws.max_column + 2
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
    return start_col


def _agregar_tabla_fechas(
    ruta: str,
    start_col: int,
    fecha_firma: Optional[pd.Timestamp],
    mes_firma: Optional[int],
    fecha_cobro: Optional[pd.Timestamp],
    mes_cobro: Optional[int],
    diferencia: Optional[int],
) -> None:
    """Agrega tabla de fechas (firma vs primer cobro) debajo de la tabla de incrementos."""
    wb = load_workbook(ruta)
    ws = wb.active

    start_row = 6  # deja un espacio en blanco después de la tabla anterior

    fill = PatternFill("solid", fgColor="1F4E78")
    font = Font(color="FFFFFF", bold=True)
    center = Alignment(horizontal="center", vertical="center")

    def _fmt_fecha(val: Optional[pd.Timestamp]):
        if pd.isna(val) or val is None:
            return ""
        try:
            return pd.to_datetime(val).to_pydatetime()
        except Exception:
            return val

    filas = [
        ("FECHA FIRMA CONTRATO", _fmt_fecha(fecha_firma), mes_firma),
        ("FECHA PRIMER COBRO", _fmt_fecha(fecha_cobro), mes_cobro),
        ("Diferencia", None, diferencia),
    ]

    for offset, (label, val_fecha, val_mes) in enumerate(filas):
        row = start_row + offset
        c_label = ws.cell(row=row, column=start_col, value=label)
        c_label.fill = fill
        c_label.font = font
        c_label.alignment = center

        c_fecha = ws.cell(row=row, column=start_col + 1, value=val_fecha if val_fecha is not None else "")
        # Marcar formato fecha cuando aplique
        try:
            import datetime as _dt
            if isinstance(val_fecha, (_dt.datetime, _dt.date)):
                c_fecha.number_format = "mm/dd/yyyy"
        except Exception:
            pass
        c_mes = ws.cell(row=row, column=start_col + 2, value=val_mes if val_mes is not None else "")
        c_fecha.alignment = center
        c_mes.alignment = center

    wb.save(ruta)


def exportar_detalles_individuales(
    detalle: pd.DataFrame,
    df_clientes: Optional[pd.DataFrame] = None,
    df_contratos: Optional[pd.DataFrame] = None,
    output_dir: str = "data/output/detalle_individual",
    aplicar_formato: bool = True,
) -> List[str]:
    """
    Genera un archivo Excel por cada contrato (clave ORDEN) y agrega
    dos tablas: meses de incremento (FDA vs auditoria) y fechas (firma vs primer cobro).
    """
    if detalle.empty:
        print("Detalle vacio: no se generan archivos individuales.")
        return []
    if not {"rfc", "subarrendatario"}.issubset(detalle.columns):
        raise ValueError("Detalle no contiene columnas requeridas: rfc y subarrendatario.")

    os.makedirs(output_dir, exist_ok=True)

    base = detalle.copy()
    base["rfc"] = base["rfc"].astype(str).str.strip()
    base["subarrendatario"] = base["subarrendatario"].astype(str).str.strip()
    base = base[(base["rfc"] != "") & (base["subarrendatario"] != "")]

    usar_orden = "orden" in base.columns
    if usar_orden:
        base["orden"] = pd.to_numeric(base["orden"], errors="coerce")
        base = base.dropna(subset=["orden"])
    if base.empty:
        print("No hay registros con RFC y subarrendatario para exportar.")
        return []

    rutas = []
    if usar_orden:
        group_iter = base.groupby(["orden"], dropna=True)
    else:
        group_iter = base.groupby(["rfc", "subarrendatario"])

    for key, df_grp in group_iter:
        if usar_orden:
            orden_val = int(pd.to_numeric(key, errors="coerce"))
            rfc = str(df_grp["rfc"].iloc[0]).strip()
            subarrendatario = str(df_grp["subarrendatario"].iloc[0]).strip()
            mes_fda_num, mes_fda_txt = _buscar_mes_fda_por_orden(df_clientes, orden_val, rfc, subarrendatario)
            mes_aud_num, mes_aud_txt = _buscar_mes_auditoria_desde_detalle(df_grp)
            fecha_firma = _buscar_fecha_firma_por_orden(df_contratos, orden_val, rfc, subarrendatario)
        else:
            orden_val = None
            rfc, subarrendatario = key
            mes_fda_num, mes_fda_txt = _buscar_mes_fda(df_clientes, rfc, subarrendatario)
            mes_aud_num, mes_aud_txt = _buscar_mes_auditoria_desde_detalle(df_grp)
            fecha_firma = _buscar_fecha_firma(df_contratos, rfc, subarrendatario)

        diferencia = (
            abs(int(mes_fda_num) - int(mes_aud_num))
            if mes_fda_num is not None and mes_aud_num is not None
            else None
        )
        mes_firma = int(pd.to_datetime(fecha_firma).month) if fecha_firma is not None else None
        fecha_primer_cobro = _buscar_fecha_primer_cobro(df_grp)
        mes_primer_cobro = int(pd.to_datetime(fecha_primer_cobro).month) if fecha_primer_cobro is not None else None
        diferencia_cobro = _diff_meses(fecha_firma, fecha_primer_cobro)

        # Limitar registros al 31/12/2025 en la tabla de detalle
        df_filtrado = df_grp.copy()
        if "fecha" in df_filtrado.columns:
            max_fecha = pd.Timestamp(year=2025, month=12, day=31)
            fechas = pd.to_datetime(df_filtrado["fecha"], errors="coerce")
            df_filtrado = df_filtrado[fechas <= max_fecha]

        if orden_val is not None:
            nombre_archivo = (
                f"{_sanitizar_nombre_archivo(str(orden_val))}_"
                f"{_sanitizar_nombre_archivo(rfc)}_"
                f"{_sanitizar_nombre_archivo(subarrendatario)}.xlsx"
            )
        else:
            nombre_archivo = f"{_sanitizar_nombre_archivo(rfc)}_{_sanitizar_nombre_archivo(subarrendatario)}.xlsx"
        ruta = os.path.join(output_dir, nombre_archivo)
        df_sorted = df_filtrado.sort_values("fecha") if "fecha" in df_filtrado.columns else df_filtrado
        df_sorted.to_excel(ruta, index=False)

        total_diferencia = _calcular_total_diferencia(df_sorted)

        if aplicar_formato:
            try:
                format_workbook(ruta)
            except Exception as exc:
                print(f"No se pudo formatear {ruta}: {exc}")

        try:
            _inserta_encabezado_presentacion(
                ruta,
                rfc=rfc,
                subarrendatario=subarrendatario,
                total_diferencia=total_diferencia,
                orden=orden_val,
            )

            start_col = _agregar_tabla_incrementos(
                ruta,
                mes_fda_num=mes_fda_num,
                mes_fda_txt=mes_fda_txt,
                mes_aud_num=mes_aud_num,
                mes_aud_txt=mes_aud_txt,
                diferencia=diferencia,
            )
            _agregar_tabla_fechas(
                ruta,
                start_col=start_col,
                fecha_firma=fecha_firma,
                mes_firma=mes_firma,
                fecha_cobro=fecha_primer_cobro,
                mes_cobro=mes_primer_cobro,
                diferencia=diferencia_cobro,
            )
        except Exception as exc:
            print(f"No se pudo agregar tabla de incrementos a {ruta}: {exc}")

        rutas.append(ruta)

    print(f"Archivos individuales generados: {len(rutas)} en {output_dir}")
    return rutas
