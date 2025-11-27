from src.loader import load_all
from src.detalle_builder import generar_detalle_todos
from src.resumen_builder import generar_resumen
from src.format_excel import format_workbook
from src.detalle_exporter import exportar_detalles_individuales

def main():
    clientes, contratos, inpc = load_all()

    # Mostrar primeras filas de referencia
    print("Clientes:\n", clientes.head(), "\n")
    print("Contratos:\n", contratos.head(), "\n")
    print("INPC:\n", inpc.head(), "\n")

    # Generar detalle mensual para todos los subarrendatarios
    detalle = generar_detalle_todos(contratos, clientes, inpc)
    print("Detalle (todos los subarrendatarios):\n", detalle.head(), "\n")
    print("Filas totales en detalle:", len(detalle))

    # Generar detalle individual por RFC + cliente (agrega tabla de incrementos por archivo)
    exportar_detalles_individuales(detalle, df_clientes=clientes, df_contratos=contratos)

    # Guardar a Excel en data/output
    output_path_detalle = "data/output/detalle_subarrendatarios.xlsx"
    detalle.to_excel(output_path_detalle, index=False)
    print(f"Detalle guardado en {output_path_detalle}")

    # Generar y guardar resumen
    resumen = generar_resumen(detalle)
    output_path_resumen = "data/output/resumen_subarrendatarios.xlsx"
    resumen.to_excel(output_path_resumen, index=False)
    print(f"Resumen guardado en {output_path_resumen}")

    # Aplicar formato a ambos archivos generados
    for path in (output_path_detalle, output_path_resumen):
        try:
            format_workbook(path)
            print(f"Formato aplicado a {path}")
        except Exception as exc:
            print(f"No se pudo formatear {path}: {exc}")

if __name__ == "__main__":
    main()
