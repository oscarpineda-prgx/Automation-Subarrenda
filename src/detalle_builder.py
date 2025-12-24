import pandas as pd


def _to_datetime(val):
    """Convierte a datetime con coerción a NaT."""
    return pd.to_datetime(val, errors="coerce")


def generar_fechas_mensuales(fecha_inicio, fecha_fin):
    """
    Genera fechas mes a mes preservando el día del inicio, hasta la fecha fin.
    """
    start = _to_datetime(fecha_inicio)
    end = _to_datetime(fecha_fin)
    if pd.isna(start) or pd.isna(end):
        return []
    # Usa DateOffset mensual para respetar el mismo dia a traves del rango
    return pd.date_range(start=start, end=end, freq=pd.DateOffset(months=1))


def calcular_importe_mtto(df):
    """
    Calcula el importe de mantenimiento de auditoría.
    Asume cuota en proporción (ej. 0.05 -> 5%). Si viniera en porcentaje (5),
    habría que dividir entre 100.
    """
    return df["importe_renta_auditoria"] * df["cuota_mantenimiento"]


def calcular_subtotal_a(df):
    """Calcula subtotal A = renta auditoría + mtto."""
    df["importe_mtto_auditoria"] = calcular_importe_mtto(df)
    df["subtotal_a"] = df["importe_renta_auditoria"] + df["importe_mtto_auditoria"]
    return df


def calcular_total_a(df):
    """Calcula total = subtotal_a * 1.16 (IVA)."""
    df["total_a"] = df["subtotal_a"] * 1.16
    return df


def _mapa_inpc(df_inpc):
    """Devuelve diccionario fecha normalizada -> porcentaje decimal."""
    if df_inpc is None or "fecha" not in df_inpc.columns:
        return {}
    col_pct = "%"
    if col_pct not in df_inpc.columns:
        candidatos = [c for c in df_inpc.columns if "porc" in c]
        col_pct = candidatos[0] if candidatos else None
    if not col_pct:
        return {}
    fechas = pd.to_datetime(df_inpc["fecha"], errors="coerce")
    pct = pd.to_numeric(df_inpc[col_pct], errors="coerce")
    return {f.normalize(): p for f, p in zip(fechas, pct) if pd.notna(f) and pd.notna(p)}


def calcular_renta_auditoria_con_inpc(fechas, renta_inicial, df_inpc):
    """
    Serie de importe_renta_auditoria aplicando ajustes INPC una vez al año,
    cuando el mes coincide con el mes de inicio (Aniversario), tomando el INPC de dos meses antes.
    """
    renta_inicial = 0 if pd.isna(renta_inicial) else renta_inicial
    if len(fechas) == 0:
        return []
    pct_map = _mapa_inpc(df_inpc)
    fechas_dt = pd.to_datetime(fechas, errors="coerce")
    mes_inicio = fechas_dt.iloc[0].month if pd.notna(fechas_dt.iloc[0]) else None
    renta_actual = renta_inicial
    serie = []
    for idx, f in enumerate(fechas_dt):
        if idx > 0 and pd.notna(f) and mes_inicio is not None and f.month == mes_inicio:
            # 2 meses antes de la fecha
            fecha_inpc = (f.normalize() - pd.DateOffset(months=2))
            pct = pct_map.get(fecha_inpc, 0) or 0
            # Ajusta renta en cada aniversario aplicando INPC
            renta_actual = renta_actual * (1 + pct)
        
        # Inserta renta actual en la serie
        serie.append(renta_actual)
    return serie

def _normalizar_clave_merge(df, col):
    """Convierte columna a numérico para merge, devolviendo serie."""
    return pd.to_numeric(df[col], errors="coerce")

def anexar_importe_renta_mtto_clientes(detalle, clientes):
    """
    Cruza contra base de clientes y agrega importe_renta_c y importe_mtto_c.

    Prioriza cruce por `orden` (ano, mes2, orden) cuando exista en ambas bases; si no, usa
    la lógica anterior por (ano, mes2, rfc). Si hay múltiples filas en clientes por la misma clave,
    se agregan (sum).
    Si no hay match, asigna 0.
    """
    det = detalle.copy()
    base_cols = ["ano", "mes2", "rfc", "importe_renta", "importe_mtto"]
    if "orden" in clientes.columns:
        base_cols.insert(2, "orden")
    cli = clientes[[c for c in base_cols if c in clientes.columns]].copy()

    # Claves normalizadas, detalle
    det['ano_key'] = _normalizar_clave_merge(det, 'ano')
    det['mes_key'] = _normalizar_clave_merge(det, 'mes')
    det['rfc_key'] = det['rfc'].astype(str).str.strip().str.upper()
    if "orden" in det.columns:
        det["orden_key"] = _normalizar_clave_merge(det, "orden")

    # Claves normalizadas, clientes
    cli['ano_key'] = _normalizar_clave_merge(cli, 'ano')
    cli['mes_key'] = _normalizar_clave_merge(cli, 'mes2')
    cli['rfc_key'] = cli['rfc'].astype(str).str.strip().str.upper()
    if "orden" in cli.columns:
        cli["orden_key"] = _normalizar_clave_merge(cli, "orden")

    # Decide claves de cruce
    usar_orden = "orden_key" in det.columns and "orden_key" in cli.columns
    if usar_orden:
        join_cols = ["ano_key", "mes_key", "orden_key", "rfc_key"]
    else:
        join_cols = ["ano_key", "mes_key", "rfc_key"]

    # Normalizar montos y agregar si hay duplicados en clientes por clave
    cli_amounts = (
        cli.assign(
            importe_renta=lambda d: pd.to_numeric(d.get("importe_renta", 0), errors="coerce").fillna(0),
            importe_mtto=lambda d: pd.to_numeric(d.get("importe_mtto", 0), errors="coerce").fillna(0),
        )
        .groupby(join_cols, as_index=False)[["importe_renta", "importe_mtto"]]
        .sum()
    )

    # Merge detalle con clientes usando claves normalizadas (ya agregadas)
    merged = det.merge(
        cli_amounts,
        on=join_cols,
        how="left",
    )
    # Montos del cliente; si no hay match se van a 0
    merged['importe_renta_c'] = pd.to_numeric(merged['importe_renta'], errors='coerce').fillna(0).round(2)
    merged['importe_mtto_c'] = pd.to_numeric(merged["importe_mtto"], errors='coerce').fillna(0).round(2)
    merged["subtotal_c"] = (merged["importe_renta_c"] + merged["importe_mtto_c"]).round(2)
    merged["total_c"] = (merged["subtotal_c"] * 1.16).round(2)
    merged["diferencia_base_vs_aud"] = (merged["total_c"] - merged["total_a"]).round(2)
    drop_cols = ['importe_renta', "importe_mtto", 'ano_key', 'mes_key', 'rfc_key']
    if "orden_key" in merged.columns:
        drop_cols.append("orden_key")
    merged = merged.drop(columns=[c for c in drop_cols if c in merged.columns])
    return merged

_MESES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]


def agregar_incremento_constante(detalle, valor="INPC"):
    """Añade columna incremento con un valor constante."""
    det = detalle.copy()
    det["incremento"] = valor
    return det


def generar_detalle_por_proveedor(df_contratos, subarrendatario, rfc_sub=None, df_inpc=None):
    """
    Compatibilidad: Genera el detalle mensual para un subarrendatario usando df_contratos normalizado.

    Nota: Cuando exista `orden` (único por contrato), se recomienda usar `generar_detalle_por_contrato`.
    """
    df_prov = df_contratos[df_contratos["nombre_del_subarrendatario"] == subarrendatario]
    if rfc_sub is not None:
        df_prov = df_prov[df_prov["rfc_del_subarrendatario"] == rfc_sub]
    if df_prov.empty:
        raise ValueError(f"Subarrendatario no encontrado: {subarrendatario} (rfc={rfc_sub})")

    row = df_prov.iloc[0]
    fecha_firma = row.get("fecha_de_firma_del_contrato")
    fecha_terminacion = row.get("fecha_de_terminacion_del_contrato")
    fechas = generar_fechas_mensuales(fecha_firma, fecha_terminacion)
    if len(fechas) == 0:
        raise ValueError(
            f"Fechas inválidas para {subarrendatario}: inicio={fecha_firma}, fin={fecha_terminacion}"
        )

    df_detalle = pd.DataFrame({"fecha": fechas})
    df_detalle["ano"] = df_detalle["fecha"].dt.year
    df_detalle["mes"] = df_detalle["fecha"].dt.month
    if "orden" in row.index and pd.notna(row.get("orden")):
        try:
            df_detalle["orden"] = int(pd.to_numeric(row.get("orden"), errors="coerce"))
        except Exception:
            df_detalle["orden"] = row.get("orden")
    df_detalle["rfc"] = row.get("rfc_del_subarrendatario")
    df_detalle["subarrendatario"] = row.get("nombre_del_subarrendatario")
    df_detalle["area"] = row.get("superficie_del_inmueble")
    df_detalle["direccion_inmueble"] = row.get("direccion_del_inmueble")

    renta = pd.to_numeric(row.get("monto_de_renta_mensual"), errors="coerce")
    cuota = pd.to_numeric(row.get("cuota_de_mantenimiento"), errors="coerce")
    renta = 0 if pd.isna(renta) else renta
    cuota = 0 if pd.isna(cuota) else cuota

    df_detalle["importe_renta_auditoria"] = calcular_renta_auditoria_con_inpc(
        df_detalle["fecha"], renta, df_inpc
    )
    df_detalle["cuota_mantenimiento"] = cuota

    df_detalle = calcular_subtotal_a(df_detalle)
    df_detalle = calcular_total_a(df_detalle)
    for col in ["importe_renta_auditoria", "importe_mtto_auditoria", "subtotal_a", "total_a"]:
        if col in df_detalle.columns:
            df_detalle[col] = df_detalle[col].round(2)
    return df_detalle


def generar_detalle_por_contrato(df_contratos, orden, df_inpc=None):
    """
    Genera el detalle mensual para un contrato usando df_contratos normalizado.
    La clave principal es `orden` (única por contrato).

    Columnas esperadas: fecha_de_firma_del_contrato, fecha_de_terminacion_del_contrato,
    rfc_del_subarrendatario, nombre_del_subarrendatario, superficie_del_inmueble,
    direccion_del_inmueble, monto_de_renta_mensual, cuota_de_mantenimiento.
    """
    if "orden" not in df_contratos.columns:
        raise ValueError("df_contratos no contiene columna 'orden'.")

    df = df_contratos.copy()
    df["orden_key"] = pd.to_numeric(df["orden"], errors="coerce")
    orden_key = pd.to_numeric(orden, errors="coerce")
    if pd.isna(orden_key):
        raise ValueError(f"Orden inválida: {orden}")

    df_con = df[df["orden_key"] == orden_key]
    if df_con.empty:
        raise ValueError(f"Contrato no encontrado para orden={orden}")
    if len(df_con) > 1:
        print(f"Aviso: orden={orden} aparece {len(df_con)} veces en contratos; se usará la primera fila.")

    row = df_con.iloc[0]

    fecha_firma = row.get("fecha_de_firma_del_contrato")
    fecha_terminacion = row.get("fecha_de_terminacion_del_contrato")
    # Crea un arreglo especializado de pandas de fechas, desde la fecha inicial hasta la final
    fechas = generar_fechas_mensuales(fecha_firma, fecha_terminacion)
    if len(fechas) == 0:
        raise ValueError(f"Fechas inválidas para orden={orden}: inicio={fecha_firma}, fin={fecha_terminacion}")

    df_detalle = pd.DataFrame({"fecha": fechas})
    df_detalle["ano"] = df_detalle["fecha"].dt.year
    df_detalle["mes"] = df_detalle["fecha"].dt.month
    df_detalle["orden"] = int(orden_key) if pd.notna(orden_key) else orden
    df_detalle["rfc"] = row.get("rfc_del_subarrendatario")
    df_detalle["subarrendatario"] = row.get("nombre_del_subarrendatario")
    df_detalle["area"] = row.get("superficie_del_inmueble")
    df_detalle["direccion_inmueble"] = row.get("direccion_del_inmueble")
    renta = pd.to_numeric(row.get("monto_de_renta_mensual"), errors="coerce")
    cuota = pd.to_numeric(row.get("cuota_de_mantenimiento"), errors="coerce")
    renta = 0 if pd.isna(renta) else renta
    cuota = 0 if pd.isna(cuota) else cuota
    df_detalle["importe_renta_auditoria"] = calcular_renta_auditoria_con_inpc(
        df_detalle["fecha"], renta, df_inpc
    )
    df_detalle["cuota_mantenimiento"] = cuota

    df_detalle = calcular_subtotal_a(df_detalle)
    df_detalle = calcular_total_a(df_detalle)
    # Redondear montos a 2 decimales (solo las columnas presentes)
    for col in ["importe_renta_auditoria", "importe_mtto_auditoria", "subtotal_a", "total_a"]:
        if col in df_detalle.columns:
            df_detalle[col] = df_detalle[col].round(2)
    return df_detalle


def generar_detalle_todos(df_contratos, df_clientes=None, df_inpc=None):
    """
    Genera el detalle mensual para todos los contratos en df_contratos.
    Concatena los resultados de generar_detalle_por_contrato y,
    si se pasa df_clientes, cruza importe_renta_c y importe_mtto_c.
    """
    detalles = []
    saltados = 0

    if "orden" in df_contratos.columns:
        ordenes = (
            pd.to_numeric(df_contratos["orden"], errors="coerce")
            .dropna()
            .astype("Int64")
            .drop_duplicates()
            .sort_values()
        )
        for orden in ordenes:
            if pd.isna(orden):
                continue
            try:
                detalles.append(
                    generar_detalle_por_contrato(df_contratos, int(orden), df_inpc=df_inpc)
                )
            except ValueError as exc:
                print(f"Saltando orden={orden}: {exc}")
                saltados += 1

        # Fallback para registros sin orden (compatibilidad)
        sin_orden = df_contratos[pd.to_numeric(df_contratos["orden"], errors="coerce").isna()]
        if not sin_orden.empty:
            claves = (
                sin_orden[["nombre_del_subarrendatario", "rfc_del_subarrendatario"]]
                .dropna(subset=["nombre_del_subarrendatario"])
                .drop_duplicates()
            )
            for _, row in claves.iterrows():
                sub = row["nombre_del_subarrendatario"]
                rfc = row["rfc_del_subarrendatario"]
                try:
                    detalles.append(
                        generar_detalle_por_proveedor(df_contratos, sub, rfc_sub=rfc, df_inpc=df_inpc)
                    )
                except ValueError as exc:
                    print(f"Saltando {sub}: {exc}")
                    saltados += 1
    else:
        # Compatibilidad: clave combinada nombre + rfc
        claves = (
            df_contratos[["nombre_del_subarrendatario", "rfc_del_subarrendatario"]]
            .dropna(subset=["nombre_del_subarrendatario"])
            .drop_duplicates()
        )
        for _, row in claves.iterrows():
            sub = row["nombre_del_subarrendatario"]
            rfc = row["rfc_del_subarrendatario"]
            try:
                detalles.append(
                    generar_detalle_por_proveedor(df_contratos, sub, rfc_sub=rfc, df_inpc=df_inpc)
                )
            except ValueError as exc:
                print(f"Saltando {sub}: {exc}")
                saltados += 1
    if not detalles:
        return pd.DataFrame()
    detalle = pd.concat(detalles, ignore_index=True)
    if df_clientes is not None:
        detalle = anexar_importe_renta_mtto_clientes(detalle, df_clientes)
    detalle = agregar_incremento_constante(detalle, valor="INPC")
    # Reordenar columnas finales
    orden = [
        "orden",
        "fecha",
        "ano",
        "mes",
        "rfc",
        "subarrendatario",
        "area",
        "direccion_inmueble",
        "cuota_mantenimiento",
        "importe_renta_c",
        "importe_mtto_c",
        "subtotal_c",
        "total_c",
        "importe_renta_auditoria",
        "importe_mtto_auditoria",
        "subtotal_a",
        "total_a",
        "diferencia_base_vs_aud",
        "incremento",
    ]
    presentes = [c for c in orden if c in detalle.columns]
    resto = [c for c in detalle.columns if c not in presentes]
    detalle = detalle[presentes + resto]
    if saltados:
        print(f"Total de contratos saltados por fechas inválidas o datos faltantes: {saltados}")
    return detalle
