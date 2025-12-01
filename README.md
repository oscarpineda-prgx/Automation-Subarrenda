# Automation Subarrenda

Herramienta en Python para automatizar el calculo y presentacion de reportes de subarrendamiento: genera un detalle mensual por subarrendatario, exporta archivos individuales con tablas auxiliares y produce un resumen ejecutivo listo para entregar.

## Contenido rapido
- Carga bases de clientes, contratos e INPC desde `data/input`.
- Calcula renta auditada ajustada por INPC, mantenimiento, subtotal y total con IVA.
- Compara contra la base del cliente (renta/mtto facturado) y calcula diferencias.
- Exporta:
  - `data/output/detalle_subarrendatarios.xlsx` (consolidado).
  - `data/output/detalle_individual/*.xlsx` (uno por RFC + subarrendatario, con logo y tablas laterales).
  - `data/output/resumen_subarrendatarios.xlsx` (resumen ejecutivo, con logo y titulo).

## Requisitos
- Python 3.x (se usa venv en el repo: `.\venv\Scripts\activate` en Windows).
- Dependencias en `requirements.txt` (pandas, numpy, openpyxl, xlsxwriter, python-dateutil, pyyaml, loguru).
- Para generar documentos Word (opcional, documento tecnico): `python-docx` (instalado en el venv, puedes agregarlo a `requirements.txt` si lo quieres persistir).

## Entradas (data/input)

| Archivo | Uso principal | Columnas clave (normalizadas) |
| --- | --- | --- |
| `base_cliente.xlsx` | Montos facturados y mes FDA | `ano`, `mes2` (mes de cobro), `mes3` (mes incremento FDA), `rfc`, `cliente`/`nombre_del_subarrendatario`, `importe_renta`, `importe_mtto` |
| `base_contratos.xlsx` | Fechas y montos de contrato auditoria | `fecha_de_firma_del_contrato`, `fecha_de_terminacion_del_contrato`, `rfc_del_subarrendatario`, `nombre_del_subarrendatario`, `superficie_del_inmueble`, `direccion_del_inmueble`, `monto_de_renta_mensual`, `cuota_de_mantenimiento` |
| `base_inpc.xlsx` | Porcentajes INPC mensuales | `fecha`, `%` (o columna con porcentaje), `mes`, `anio` |
| `Picture1.png` | Logo para encabezados | Imagen insertada en resumen y detalles individuales |

Rutas actuales en Citrix estan definidas como absolutas en `src/loader.py`; dejando los archivos en `data/input` no necesitas cambiar nada.

## Flujo general (alto nivel)

```
main.py
  ├─ load_all()        # loader.py: lee y normaliza bases
  ├─ generar_detalle_todos()  # detalle_builder.py: calcula detalle mensual + INPC
  ├─ exportar_detalles_individuales()  # detalle_exporter.py: un Excel por RFC+cliente
  ├─ to_excel(detalle/resumen)         # guarda consolidados
  ├─ format_workbook()   # format_excel.py: formatos fecha/moneda/encabezado
  └─ aplicar_presentacion_resumen()  # resumen_builder.py: logo + titulo
```

## Como ejecutarlo

1) Activar entorno:
```
.\venv\Scripts\activate
```
2) Instalar dependencias:
```
pip install -r requirements.txt

```
3) Verifica/actualiza archivos de entrada en `data/input`.
4) Ejecuta:
```
python main.py
```
5) Revisa salidas en `data/output` y `data/output/detalle_individual`.

## Modulos principales

| Modulo | Rol |
| --- | --- |
| `main.py` | Orquestador: carga datos, genera detalle y resumen, exporta archivos, aplica formato y presentacion. |
| `src/loader.py` | Lectura de bases de clientes/contratos/INPC; normaliza nombres (trim, minusculas, guiones bajos, sin acentos) y convierte fechas. |
| `src/detalle_builder.py` | Genera detalle mensual: fechas mes a mes, renta auditada ajustada por INPC (aniversario usa INPC de dos meses antes), mantenimiento, subtotal y total con IVA; cruza con base cliente por ano/mes/RFC y calcula diferencias. |
| `src/detalle_exporter.py` | Crea un Excel por RFC + subarrendatario con logo, total de diferencia, tabla de meses de incremento (FDA vs auditoria) y tabla de fechas (firma vs primer cobro). |
| `src/resumen_builder.py` | Agrupa por RFC + subarrendatario: sumas de mtto/subtotales/totales solo donde hay cobro, primer importe > 0, banderas SI/NO HAY COBRO, necesidad de acta de entrega, meses sin cobro; renombra y ordena columnas y agrega presentacion al resumen. |
| `src/format_excel.py` | Formato uniforme: oculta cuadricula, fechas `mm/dd/yyyy`, formato contable para montos, encabezados azul con texto blanco. |

## Reglas de negocio clave
- Ajuste INPC: la renta auditada sube en cada aniversario (mes de firma), usando el INPC de dos meses antes.
- Mantenimiento: `importe_mtto_auditoria = renta_auditoria * cuota_mantenimiento`; `total_a = subtotal_a * 1.16`.
- Cruce con clientes: busca por `ano` + `mes2` + `rfc`; faltantes se tratan como 0; `diferencia_base_vs_aud = total_c - total_a`.
- Acta de entrega: Necesaria si auditoria sube en aniversario y cliente no sube ese mes; Innecesaria si ambos suben; Desconocida si faltan datos.
- Meses sin cobro: cuenta meses sin importe_renta_c > 0 entre el primer y ultimo cobro.

## Salidas
- `data/output/detalle_subarrendatarios.xlsx`: detalle mensual consolidado.
- `data/output/detalle_individual/*.xlsx`: un archivo por RFC + subarrendatario con tablas laterales y formato.
- `data/output/resumen_subarrendatarios.xlsx`: resumen ejecutivo con logo y titulo.
- (Opcional) `data/output/Documentacion_Proyecto_Subarrenda.docx`: documento tecnico generado con python-docx.

## Solucion de problemas rapidos
- Columnas faltantes o nombres distintos: revisa `src/loader.py` para ajustar normalizacion/renombres.
- Fechas invalidas: se convierten con `errors="coerce"` a NaT; si faltan fechas de contrato, ese subarrendatario se salta.
- Sin datos de INPC: la renta auditada no se ajusta (permanece con ultimo valor).
- Logo no aparece: confirma `data/input/Picture1.png`.

## Extensiones sugeridas
- Agregar validacion de esquema de archivos de entrada antes de procesar.
- Incluir pruebas unitarias para funciones de fecha/INPC/cruces.
- Exponer configuracion de rutas/archivos via YAML para evitar rutas absolutas.
