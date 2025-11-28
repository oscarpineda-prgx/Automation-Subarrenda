import sys
import pandas as pd
import os

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.drawing.image import Image

from src.format_excel import format_workbook

# Permitir imports desde src
sys.path.append("src")
from loader import load_all
from detalle_builder import generar_detalle_todos


def generar_resumen(detalle: pd.DataFrame) -> pd.DataFrame:
    """
    Genera un resumen agrupando por RFC y subarrendatario.
    - Agrupa por rfc y subarrendatario.
    - Suma diferencia_base_vs_aud.
    - Agrega sumas de mtto/subtotales/totales solo sobre filas con cobro (importe_renta_c > 0).
    - Incluye cuota de mantenimiento (%).
    - Incluye el primer importe_renta_c > 0 y el primer importe_renta_auditoria > 0 (o 0 si no existen).
    - Marca si hay cobro (importe_renta_c > 0 en alguna fila) o NO HAY COBRO.
    - No elimina RFC duplicados si el subarrendatario difiere (clave = rfc + subarrendatario).
    """
    if detalle.empty:
        return pd.DataFrame(
            columns=[
                "rfc",
                "subarrendatario",
                "area",
                "f_contrato_ini",
                "f_contrato_fin",
                "renta_mensual",
                "duracion_meses",
                "cuota_mantenimiento",
                "sum_importe_mtto_c",
                "sum_importe_mtto_a",
                "sum_subtotal_c",
                "sum_subtotal_a",
                "sum_total_c",
                "sum_total_a",
                "sum_dif_base_vs_aud",
                "primer_importe_renta_c",
                "primer_importe_renta_a",
                "estatus_cobro",
                "count_meses_sin_cobro",
            ]
        )

    def _primer_importe(df_base: pd.DataFrame, columna: str) -> pd.Series:
        """
        Devuelve serie con el primer valor > 0 por rfc/subarrendatario para la columna indicada.
        Si no hay valores > 0, devuelve 0.
        """
        if columna not in df_base.columns:
            return pd.Series(dtype=float)

        ordenado = (
            df_base.sort_values("fecha")
            if "fecha" in df_base.columns
            else df_base.copy()
        )
        serie = (
            ordenado[ordenado[columna] > 0]
            .groupby(["rfc", "subarrendatario"])[columna]
            .first()
            .round(2)
        )
        return serie

    base_primeros = detalle.copy()
    primeros_cliente = _primer_importe(base_primeros, "importe_renta_c")
    primeros_auditoria = _primer_importe(base_primeros, "importe_renta_auditoria")

    def _sumas_condicion_cobro(df_base: pd.DataFrame) -> pd.DataFrame:
        """
        Suma importes solo en filas con cobro (importe_renta_c > 0).
        Devuelve dataframe con columnas de sumas y cuota de mantenimiento.
        """
        cols = [
            "rfc",
            "subarrendatario",
            "importe_mtto_c_cobro_sum",
            "importe_mtto_auditoria_cobro_sum",
            "subtotal_c_cobro_sum",
            "subtotal_a_cobro_sum",
            "total_c_cobro_sum",
            "total_a_cobro_sum",
            "cuota_mantenimiento_pct",
        ]
        if df_base.empty:
            return pd.DataFrame(columns=cols)

        claves = df_base[["rfc", "subarrendatario"]].drop_duplicates()

        con_cobro = df_base[df_base["importe_renta_c"].fillna(0) > 0].copy()
        sumas = (
            con_cobro.assign(
                importe_mtto_c=lambda d: pd.to_numeric(d.get("importe_mtto_c", 0), errors="coerce").fillna(0),
                importe_mtto_auditoria=lambda d: pd.to_numeric(d.get("importe_mtto_auditoria", 0), errors="coerce").fillna(0),
                subtotal_c=lambda d: pd.to_numeric(d.get("subtotal_c", 0), errors="coerce").fillna(0),
                subtotal_a=lambda d: pd.to_numeric(d.get("subtotal_a", 0), errors="coerce").fillna(0),
                total_c=lambda d: pd.to_numeric(d.get("total_c", 0), errors="coerce").fillna(0),
                total_a=lambda d: pd.to_numeric(d.get("total_a", 0), errors="coerce").fillna(0),
            )
            .groupby(["rfc", "subarrendatario"])
            .agg(
                importe_mtto_c_cobro_sum=("importe_mtto_c", "sum"),
                importe_mtto_auditoria_cobro_sum=("importe_mtto_auditoria", "sum"),
                subtotal_c_cobro_sum=("subtotal_c", "sum"),
                subtotal_a_cobro_sum=("subtotal_a", "sum"),
                total_c_cobro_sum=("total_c", "sum"),
                total_a_cobro_sum=("total_a", "sum"),
            )
            .reset_index()
        )

        # Cuota por RFC (solo RFC, sin depender de que haya cobro)
        cuota_map = None
        if "cuota_mantenimiento" in df_base.columns:
            cuota_map = (
                df_base.assign(cuota_mantenimiento=lambda d: pd.to_numeric(d["cuota_mantenimiento"], errors="coerce"))
                .groupby("rfc")["cuota_mantenimiento"]
                .first()
            )

        # Ensamblar con todas las claves rfc+subarrendatario
        sumas = claves.merge(sumas, on=["rfc", "subarrendatario"], how="left")
        sumas["cuota_mantenimiento_pct"] = sumas["rfc"].map(cuota_map) if cuota_map is not None else None

        for col in cols:
            if col not in sumas.columns:
                sumas[col] = None
        return sumas[cols]

    sumas_cobro = _sumas_condicion_cobro(detalle)

    def _count_meses_sin_cobro(df_base: pd.DataFrame) -> pd.Series:
        """
        Cuenta meses sin cobro dentro del rango entre el primer y ultimo importe_renta_c > 0.
        Si nunca hay cobro, devuelve 0.
        """
        if "importe_renta_c" not in df_base.columns:
            return pd.Series(dtype=int)

        def contar(grp: pd.DataFrame) -> int:
            ordenado = grp.sort_values("fecha") if "fecha" in grp.columns else grp.copy()
            vals = pd.to_numeric(ordenado["importe_renta_c"], errors="coerce").fillna(0)
            pos = vals[vals > 0]
            if pos.empty:
                return 0
            start = pos.index[0]
            end = pos.index[-1]
            rango = vals.loc[start:end]
            return int((rango <= 0).sum())

        res = df_base.groupby(["rfc", "subarrendatario"]).apply(contar)
        res.name = "count_meses_sin_cobro"
        return res

    count_meses = _count_meses_sin_cobro(detalle)
    # Info de contrato derivada del detalle (sin agregar columnas nuevas al detalle)
    agregados = {}
    if "area" in detalle.columns:
        agregados["area"] = ("area", "first")
    if "fecha" in detalle.columns:
        agregados["fecha_contrato_inicio"] = ("fecha", "min")
        agregados["fecha_contrato_fin"] = ("fecha", "max")
        agregados["duracion_meses_contrato"] = ("fecha", "count")
    if "importe_renta_auditoria" in detalle.columns:
        agregados["renta_mensual_contrato"] = ("importe_renta_auditoria", "first")

    if agregados:
        contrato_info = (
            detalle.groupby(["rfc", "subarrendatario"])
            .agg(**agregados)
            .reset_index()
        )
    else:
        contrato_info = pd.DataFrame(columns=["rfc", "subarrendatario"])

    # Base completa de claves rfc + subarrendatario (no se eliminan duplicados de RFC con nombres distintos)
    claves = detalle[["rfc", "subarrendatario"]].drop_duplicates()

    # Agrupa diferencia solo con cobro, pero mantiene todas las claves (resto queda en 0)
    # La diferencia solo se suma donde importe_renta_c > 0; si no hay cobro, queda 0.
    df_dif = detalle[
        (detalle["importe_renta_c"].fillna(0) > 0)
    ].copy()
    diferencia_sum = (
        df_dif.fillna({"diferencia_base_vs_aud": 0})
        .groupby(["rfc", "subarrendatario"], as_index=False)["diferencia_base_vs_aud"]
        .sum()
        .rename(columns={"diferencia_base_vs_aud": "diferencia_base_vs_aud_sum"})
    )
    diferencia_sum = claves.merge(
        diferencia_sum, on=["rfc", "subarrendatario"], how="left"
    ).fillna({"diferencia_base_vs_aud_sum": 0})
    # Bandera de cobro: si existe al menos un importe_renta_c > 0
    estatus = (
        detalle.fillna({"importe_renta_c": 0})
        .assign(tiene_cobro=lambda d: d["importe_renta_c"] > 0)
        .groupby(["rfc", "subarrendatario"])["tiene_cobro"]
        .any()
        .replace({True: "SI HAY COBRO", False: "NO HAY COBRO"})
    )

    resumen = diferencia_sum
    resumen = resumen.merge(
        primeros_cliente.rename("primer_importe_renta_c"),
        on=["rfc", "subarrendatario"],
        how="left",
    )
    resumen = resumen.merge(
        primeros_auditoria.rename("primer_importe_renta_auditoria"),
        on=["rfc", "subarrendatario"],
        how="left",
    )
    resumen["primer_importe_renta_c"] = resumen[
        "primer_importe_renta_c"
    ].fillna(0)
    resumen["primer_importe_renta_auditoria"] = resumen[
        "primer_importe_renta_auditoria"
    ].fillna(0)
    resumen = resumen.merge(
        estatus.rename("estatus_cobro"), on=["rfc", "subarrendatario"], how="left"
    )
    resumen["estatus_cobro"] = resumen["estatus_cobro"].fillna("NO HAY COBRO")

    resumen = resumen.merge(
        count_meses.rename("count_meses_sin_cobro"),
        on=["rfc", "subarrendatario"],
        how="left",
    )
    resumen["count_meses_sin_cobro"] = resumen["count_meses_sin_cobro"].fillna(0).astype(int)

    resumen = resumen.merge(sumas_cobro, on=["rfc", "subarrendatario"], how="left")
    for c in [
        "importe_mtto_c_cobro_sum",
        "importe_mtto_auditoria_cobro_sum",
        "subtotal_c_cobro_sum",
        "subtotal_a_cobro_sum",
        "total_c_cobro_sum",
        "total_a_cobro_sum",
    ]:
        if c in resumen.columns:
            resumen[c] = resumen[c].fillna(0).round(2)
    if "cuota_mantenimiento_pct" in resumen.columns:
        resumen["cuota_mantenimiento_pct"] = resumen["cuota_mantenimiento_pct"].round(4)

    if not contrato_info.empty:
        resumen = resumen.merge(contrato_info, on=["rfc", "subarrendatario"], how="left")

    # Renombrar columnas a los alias solicitados por el usuario
    column_renames = {
        "fecha_contrato_inicio": "f_contrato_ini",
        "fecha_contrato_fin": "f_contrato_fin",
        "renta_mensual_contrato": "renta_mensual",
        "duracion_meses_contrato": "duracion_meses",
        "cuota_mantenimiento_pct": "cuota_mantenimiento",
        "importe_mtto_c_cobro_sum": "sum_importe_mtto_c",
        "subtotal_c_cobro_sum": "sum_subtotal_c",
        "total_c_cobro_sum": "sum_total_c",
        "primer_importe_renta_auditoria": "primer_importe_renta_a",
        "importe_mtto_auditoria_cobro_sum": "sum_importe_mtto_a",
        "subtotal_a_cobro_sum": "sum_subtotal_a",
        "total_a_cobro_sum": "sum_total_a",
        "diferencia_base_vs_aud_sum": "sum_dif_base_vs_aud",
    }
    resumen = resumen.rename(columns=column_renames)

    # Reordenar columnas segun solicitud
    orden = [
        "rfc",
        "subarrendatario",
        "area",
        "f_contrato_ini",
        "f_contrato_fin",
        "renta_mensual",
        "duracion_meses",
        "cuota_mantenimiento",
        "primer_importe_renta_c",
        "sum_importe_mtto_c",
        "sum_subtotal_c",
        "sum_total_c",
        "primer_importe_renta_a",
        "sum_importe_mtto_a",
        "sum_subtotal_a",
        "sum_total_a",
        "sum_dif_base_vs_aud",
        "estatus_cobro",
        "count_meses_sin_cobro",
    ]
    presentes = [c for c in orden if c in resumen.columns]
    resto = [c for c in resumen.columns if c not in presentes]
    resumen = resumen[presentes + resto]

    return resumen


def generar_resumen_desde_fuente() -> pd.DataFrame:
    """
    Carga clientes/contratos/INPC, genera el detalle completo y devuelve el resumen,
    sin leer/escribir archivos intermedios.
    """
    clientes, contratos, inpc = load_all()
    detalle = generar_detalle_todos(contratos, clientes, inpc)
    return generar_resumen(detalle)


def aplicar_presentacion_resumen(
    ruta: str,
    logo_path: str = "data/input/Picture1.png",
) -> None:
    """Inserta logo y titulos en la parte superior de la hoja."""
    wb = load_workbook(ruta)
    ws = wb.active

    # Reservar espacio arriba del encabezado original
    espacio = 7
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

    titulo = ws.cell(row=4, column=6, value="RESUMEN SUBARRENDATARIOS - DIFERENCIA DE COBRO")
    titulo.font = Font(bold=True, size=14)
    titulo.alignment = Alignment(horizontal="center")
    ws.merge_cells(start_row=4, start_column=6, end_row=4, end_column=13)

    wb.save(ruta)


if __name__ == "__main__":
    resumen = generar_resumen_desde_fuente()
    output_path = "data/output/resumen_subarrendatarios.xlsx"
    resumen.to_excel(output_path, index=False)
    aplicar_presentacion_resumen(output_path)
    print(f"Resumen guardado en {output_path}")
