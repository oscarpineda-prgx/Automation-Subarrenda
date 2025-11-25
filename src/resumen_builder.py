import sys
import pandas as pd

# Permitir imports desde src
sys.path.append("src")
from loader import load_all
from detalle_builder import generar_detalle_todos


def generar_resumen(detalle: pd.DataFrame) -> pd.DataFrame:
    """
    Genera un resumen agrupando por RFC y cliente.
    - Filtra filas con importe_renta_cliente != 0.
    - Agrupa por rfc y cliente.
    - Suma diferencia_base_vs_aud.
    """
    if detalle.empty:
        return detalle

    df = detalle[detalle["importe_renta_cliente"] != 0].copy()
    if df.empty:
        return pd.DataFrame(columns=["rfc", "cliente", "diferencia_base_vs_aud_sum"])

    resumen = (
        df.groupby(["rfc", "cliente"], as_index=False)["diferencia_base_vs_aud"]
        .sum()
        .rename(columns={"diferencia_base_vs_aud": "diferencia_base_vs_aud_sum"})
    )
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
