import sys
import pandas as pd

# Permitir imports desde src
sys.path.append("src")
from loader import load_all
from detalle_builder import generar_detalle_todos


def generar_resumen(detalle: pd.DataFrame) -> pd.DataFrame:
    """
    Genera un resumen agrupando por RFC y cliente.
    - Agrupa por rfc y cliente.
    - Suma diferencia_base_vs_aud.
    - Incluye el primer importe_renta_cliente > 0 y el primer importe_renta_auditoria > 0 (o 0 si no existen).
    - Marca si hay cobro (importe_renta_cliente > 0 en alguna fila) o NO HAY COBRO.
    - No elimina RFC duplicados si el cliente difiere (clave = rfc + cliente).
    """
    if detalle.empty:
        return detalle

    if detalle.empty:
        return pd.DataFrame(
            columns=[
                "rfc",
                "cliente",
                "diferencia_base_vs_aud_sum",
                "primer_importe_renta_cliente",
                "primer_importe_renta_auditoria",
                "estatus_cobro",
            ]
        )

    def _primer_importe(df_base: pd.DataFrame, columna: str) -> pd.Series:
        """
        Devuelve serie con el primer valor > 0 por rfc/cliente para la columna indicada.
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
            .groupby(["rfc", "cliente"])[columna]
            .first()
            .round(2)
        )
        return serie

    base_primeros = detalle.copy()
    primeros_cliente = _primer_importe(base_primeros, "importe_renta_cliente")
    primeros_auditoria = _primer_importe(base_primeros, "importe_renta_auditoria")

    # Base completa de claves rfc + cliente (no se eliminan duplicados de RFC con nombres distintos)
    claves = detalle[["rfc", "cliente"]].drop_duplicates()

    # Agrupa diferencia solo con cobro, pero mantiene todas las claves (resto queda en 0)
    # La diferencia solo se suma donde importe_renta_cliente > 0; si no hay cobro, queda 0.
    df_dif = detalle[
        (detalle["importe_renta_cliente"].fillna(0) > 0)
    ].copy()
    diferencia_sum = (
        df_dif.fillna({"diferencia_base_vs_aud": 0})
        .groupby(["rfc", "cliente"], as_index=False)["diferencia_base_vs_aud"]
        .sum()
        .rename(columns={"diferencia_base_vs_aud": "diferencia_base_vs_aud_sum"})
    )
    diferencia_sum = claves.merge(
        diferencia_sum, on=["rfc", "cliente"], how="left"
    ).fillna({"diferencia_base_vs_aud_sum": 0})
    # Bandera de cobro: si existe al menos un importe_renta_cliente > 0
    estatus = (
        detalle.fillna({"importe_renta_cliente": 0})
        .assign(tiene_cobro=lambda d: d["importe_renta_cliente"] > 0)
        .groupby(["rfc", "cliente"])["tiene_cobro"]
        .any()
        .replace({True: "SI HAY COBRO", False: "NO HAY COBRO"})
    )

    resumen = diferencia_sum
    resumen = resumen.merge(
        primeros_cliente.rename("primer_importe_renta_cliente"),
        on=["rfc", "cliente"],
        how="left",
    )
    resumen = resumen.merge(
        primeros_auditoria.rename("primer_importe_renta_auditoria"),
        on=["rfc", "cliente"],
        how="left",
    )
    resumen["primer_importe_renta_cliente"] = resumen[
        "primer_importe_renta_cliente"
    ].fillna(0)
    resumen["primer_importe_renta_auditoria"] = resumen[
        "primer_importe_renta_auditoria"
    ].fillna(0)
    resumen = resumen.merge(
        estatus.rename("estatus_cobro"), on=["rfc", "cliente"], how="left"
    )
    resumen["estatus_cobro"] = resumen["estatus_cobro"].fillna("NO HAY COBRO")

    return resumen


def generar_resumen_desde_fuente() -> pd.DataFrame:
    """
    Carga clientes/contratos/INPC, genera el detalle completo y devuelve el resumen,
    sin leer/escribir archivos intermedios.
    """
    clientes, contratos, inpc = load_all()
    detalle = generar_detalle_todos(contratos, clientes, inpc)
    return generar_resumen(detalle)


if __name__ == "__main__":
    resumen = generar_resumen_desde_fuente()
    output_path = "data/output/resumen_subarrendatarios.xlsx"
    resumen.to_excel(output_path, index=False)
    print(f"Resumen guardado en {output_path}")
