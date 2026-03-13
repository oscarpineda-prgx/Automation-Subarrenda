import pandas as pd

_MAX_FECHA_REVISION = pd.Timestamp(year=2025, month=12, day=31)


def _to_datetime(val):
    """Convierte a datetime con coercion a NaT."""
    return pd.to_datetime(val, errors="coerce")


def generar_fechas_mensuales(fecha_inicio, fecha_fin, fecha_max=None):
    """
    Genera fechas mes a mes preservando el dia del inicio, hasta la fecha fin.
    Si fecha_max existe, limita el rango al menor entre fecha_fin y fecha_max.
    """
    start = _to_datetime(fecha_inicio)
    end = _to_datetime(fecha_fin)
    if fecha_max is not None:
        max_dt = _to_datetime(fecha_max)
        if pd.notna(max_dt) and pd.notna(end):
            end = min(end, max_dt)
        elif pd.notna(max_dt) and pd.isna(end):
            end = max_dt
    if pd.isna(start) or pd.isna(end) or start > end:
        return []
    return pd.date_range(start=start, end=end, freq=pd.DateOffset(months=1))


def calcular_importe_mtto(df):
    """
    Calcula el importe de mantenimiento de auditoria.
    Asume cuota en proporcion (ej. 0.05 -> 5%).
    """
    return df["importe_renta_auditoria"] * df["cuota_mantenimiento"]


def calcular_subtotal_a(df):
    """Calcula subtotal A = renta auditoria + mtto."""
    df["importe_mtto_auditoria"] = calcular_importe_mtto(df)
    df["subtotal_a"] = df["importe_renta_auditoria"] + df["importe_mtto_auditoria"]
    return df


def calcular_total_a(df):
    """Calcula total = subtotal_a * 1.16 (IVA)."""
    df["total_a"] = df["subtotal_a"] * 1.16
    return df


def _mapa_inpc(df_inpc):
    """Devuelve diccionario (ano, mes) -> porcentaje decimal."""
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
    return {(f.year, f.month): p for f, p in zip(fechas, pct) if pd.notna(f) and pd.notna(p)}


def _normalizar_mes_incremento(val):
    if pd.isna(val):
        return None
    if isinstance(val, (int, float)):
        num = int(val)
        return num if 1 <= num <= 12 else None
    txt = str(val).strip().upper()
    if not txt:
        return None
    if txt.isdigit():
        num = int(txt)
        return num if 1 <= num <= 12 else None
    meses = {
        "ENERO": 1,
        "FEBRERO": 2,
        "MARZO": 3,
        "ABRIL": 4,
        "MAYO": 5,
        "JUNIO": 6,
        "JULIO": 7,
        "AGOSTO": 8,
        "SEPTIEMBRE": 9,
        "OCTUBRE": 10,
        "NOVIEMBRE": 11,
        "DICIEMBRE": 12,
    }
    if txt in meses:
        return meses[txt]
    for nombre, num in meses.items():
        if nombre.startswith(txt[:3]):
            return num
    return None


def _mapa_mes_incremento_por_orden(df_clientes):
    """Mapea ORDEN -> mes de incremento (columna MES3 de base_cliente)."""
    if df_clientes is None or df_clientes.empty:
        return {}
    if "orden" not in df_clientes.columns or "mes3" not in df_clientes.columns:
        return {}

    cli = df_clientes.copy()
    cli["orden_key"] = pd.to_numeric(cli["orden"], errors="coerce")
    cli = cli.dropna(subset=["orden_key"])
    if cli.empty:
        return {}

    cli["mes_incremento"] = cli["mes3"].map(_normalizar_mes_incremento)
    cli = cli.dropna(subset=["mes_incremento"])
    if cli.empty:
        return {}

    return (
        cli.groupby("orden_key", as_index=False)["mes_incremento"]
        .first()
        .set_index("orden_key")["mes_incremento"]
        .to_dict()
    )


def _primer_texto_no_vacio(serie: pd.Series):
    """Devuelve el primer texto no vacio de una serie; si no existe, regresa None."""
    for val in serie:
        if pd.isna(val):
            continue
        txt = str(val).strip()
        if txt:
            return txt
    return None


def _mapa_identidad_cliente_por_orden(df_clientes: pd.DataFrame) -> pd.DataFrame:
    """
    Construye un mapa de identidad por contrato desde base_cliente:
    ORDEN -> RFC + CLIENTE.
    """
    if df_clientes is None or df_clientes.empty or "orden" not in df_clientes.columns:
        return pd.DataFrame(columns=["orden_key", "rfc_cliente", "cliente"])

    cli = df_clientes.copy()
    cli["orden_key"] = pd.to_numeric(cli["orden"], errors="coerce")
    cli = cli.dropna(subset=["orden_key"])
    if cli.empty:
        return pd.DataFrame(columns=["orden_key", "rfc_cliente", "cliente"])

    if "rfc" not in cli.columns:
        cli["rfc"] = None
    if "cliente" not in cli.columns:
        cli["cliente"] = None

    identidad = (
        cli.groupby("orden_key", as_index=False)
        .agg(
            rfc_cliente=("rfc", _primer_texto_no_vacio),
            cliente=("cliente", _primer_texto_no_vacio),
        )
    )
    return identidad


def aplicar_identidad_cliente_por_orden(detalle: pd.DataFrame, df_clientes: pd.DataFrame) -> pd.DataFrame:
    """
    Sobrescribe identificadores del detalle con datos de base_cliente por ORDEN.
    Prioriza: ORDEN + RFC + CLIENTE, con fallback a valores existentes si faltan datos.
    """
    if detalle is None or detalle.empty:
        return detalle
    if df_clientes is None or df_clientes.empty:
        return detalle
    if "orden" not in detalle.columns:
        return detalle

    identidad = _mapa_identidad_cliente_por_orden(df_clientes)
    if identidad.empty:
        return detalle

    det = detalle.copy()
    det["orden_key"] = pd.to_numeric(det["orden"], errors="coerce")
    det = det.merge(identidad, on="orden_key", how="left")

    rfc_cliente = det["rfc_cliente"].map(lambda v: "" if pd.isna(v) else str(v).strip())
    cli_cliente = det["cliente"].map(lambda v: "" if pd.isna(v) else str(v).strip())

    if "rfc" not in det.columns:
        det["rfc"] = None
    if "subarrendatario" not in det.columns:
        det["subarrendatario"] = None

    det["rfc"] = det["rfc"].where(rfc_cliente == "", rfc_cliente)
    det["subarrendatario"] = det["subarrendatario"].where(cli_cliente == "", cli_cliente)

    det = det.drop(columns=["orden_key", "rfc_cliente", "cliente"])
    return det


def calcular_renta_auditoria_con_inpc(fechas, renta_inicial, df_inpc, mes_incremento=None):
    """
    Serie de importe_renta_auditoria aplicando ajustes INPC una vez al ano,
    cuando el mes coincide con el mes de incremento configurado.
    Se usa el INPC de dos meses antes.
    """
    renta_inicial = 0 if pd.isna(renta_inicial) else renta_inicial
    if len(fechas) == 0:
        return []

    pct_map = _mapa_inpc(df_inpc)
    fechas_dt = pd.to_datetime(fechas, errors="coerce")

    mes_incremento = _normalizar_mes_incremento(mes_incremento)

    renta_actual = renta_inicial
    serie = []
    for idx, f in enumerate(fechas_dt):
        if idx > 0 and pd.notna(f) and mes_incremento is not None and f.month == mes_incremento:
            fecha_inpc = f.normalize() - pd.DateOffset(months=2)
            pct = pct_map.get((fecha_inpc.year, fecha_inpc.month), 0) or 0
            renta_actual = renta_actual * (1 + pct)
        serie.append(renta_actual)
    return serie


def _normalizar_clave_merge(df, col):
    """Convierte columna a numerico para merge, devolviendo serie."""
    return pd.to_numeric(df[col], errors="coerce")


def anexar_importe_renta_mtto_clientes(detalle, clientes):
    """
    Cruza contra base de clientes y agrega importe_renta_c e importe_mtto_c.
    Regla: el cruce se hace unicamente por ORDEN + periodo (ano, mes).
    """
    det = detalle.copy()
    base_cols = ["ano", "mes2", "mes", "orden", "importe_renta", "importe_mtto", "incremento"]
    cli = clientes[[c for c in base_cols if c in clientes.columns]].copy()

    if "incremento" not in cli.columns:
        cli["incremento"] = None

    det["ano_key"] = _normalizar_clave_merge(det, "ano")
    det["mes_key"] = _normalizar_clave_merge(det, "mes")
    det["orden_key"] = _normalizar_clave_merge(det, "orden") if "orden" in det.columns else pd.NA

    cli["ano_key"] = _normalizar_clave_merge(cli, "ano")
    mes_col_cliente = "mes2" if "mes2" in cli.columns else "mes"
    cli["mes_key"] = _normalizar_clave_merge(cli, mes_col_cliente)
    cli["orden_key"] = _normalizar_clave_merge(cli, "orden") if "orden" in cli.columns else pd.NA

    join_cols = ["ano_key", "mes_key", "orden_key"]
    cli = cli.dropna(subset=["orden_key"])
    det = det.dropna(subset=["orden_key"])

    cli_amounts = (
        cli.assign(
            importe_renta=lambda d: pd.to_numeric(d.get("importe_renta", 0), errors="coerce").fillna(0),
            importe_mtto=lambda d: pd.to_numeric(d.get("importe_mtto", 0), errors="coerce").fillna(0),
        )
        .groupby(join_cols, as_index=False)
        .agg(
            importe_renta=("importe_renta", "sum"),
            importe_mtto=("importe_mtto", "sum"),
            incremento=("incremento", "first"),
        )
    )

    inc_lookup = (
        cli[["orden_key", "incremento"]]
        .dropna(subset=["orden_key"])
        .groupby("orden_key", as_index=False)["incremento"]
        .first()
        .rename(columns={"incremento": "incremento_lookup"})
    )

    merged = det.merge(cli_amounts, on=join_cols, how="left")
    merged["importe_renta_c"] = pd.to_numeric(merged["importe_renta"], errors="coerce").fillna(0).round(2)
    merged["importe_mtto_c"] = pd.to_numeric(merged["importe_mtto"], errors="coerce").fillna(0).round(2)
    merged["subtotal_c"] = (merged["importe_renta_c"] + merged["importe_mtto_c"]).round(2)
    merged["total_c"] = (merged["subtotal_c"] * 1.16).round(2)
    merged["diferencia_base_vs_aud"] = (merged["total_c"] - merged["total_a"]).round(2)

    merged = merged.merge(inc_lookup, on="orden_key", how="left")
    merged["incremento"] = merged["incremento"].fillna(merged.get("incremento_lookup"))
    merged = merged.drop(columns=[c for c in ["incremento_lookup"] if c in merged.columns])
    merged["incremento"] = merged["incremento"].fillna("INPC")

    drop_cols = ["importe_renta", "importe_mtto", "ano_key", "mes_key", "orden_key"]
    merged = merged.drop(columns=[c for c in drop_cols if c in merged.columns])
    return merged


_MESES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]


def agregar_incremento_constante(detalle, valor="INPC"):
    """Anade columna incremento con un valor constante."""
    det = detalle.copy()
    if "incremento" not in det.columns:
        det["incremento"] = valor
    else:
        det["incremento"] = det["incremento"].fillna(valor)
    return det


def generar_detalle_por_proveedor(df_contratos, subarrendatario, rfc_sub=None, df_inpc=None, mes_incremento=None):
    """
    Compatibilidad: Genera detalle mensual para un subarrendatario.
    """
    df_prov = df_contratos[df_contratos["nombre_del_subarrendatario"] == subarrendatario]
    if rfc_sub is not None:
        df_prov = df_prov[df_prov["rfc_del_subarrendatario"] == rfc_sub]
    if df_prov.empty:
        raise ValueError(f"Subarrendatario no encontrado: {subarrendatario} (rfc={rfc_sub})")

    row = df_prov.iloc[0]
    fecha_firma = row.get("fecha_de_firma_del_contrato")
    fecha_terminacion = row.get("fecha_de_terminacion_del_contrato")
    fechas = generar_fechas_mensuales(fecha_firma, fecha_terminacion, fecha_max=_MAX_FECHA_REVISION)
    if len(fechas) == 0:
        raise ValueError(
            f"Fechas invalidas para {subarrendatario}: inicio={fecha_firma}, fin={fecha_terminacion}"
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
    df_detalle["mes_incremento"] = mes_incremento

    renta = pd.to_numeric(row.get("monto_de_renta_mensual"), errors="coerce")
    cuota = pd.to_numeric(row.get("cuota_de_mantenimiento"), errors="coerce")
    renta = 0 if pd.isna(renta) else renta
    cuota = 0 if pd.isna(cuota) else cuota

    df_detalle["importe_renta_auditoria"] = calcular_renta_auditoria_con_inpc(
        df_detalle["fecha"], renta, df_inpc, mes_incremento=mes_incremento
    )
    df_detalle["cuota_mantenimiento"] = cuota

    df_detalle = calcular_subtotal_a(df_detalle)
    df_detalle = calcular_total_a(df_detalle)
    for col in ["importe_renta_auditoria", "importe_mtto_auditoria", "subtotal_a", "total_a"]:
        if col in df_detalle.columns:
            df_detalle[col] = df_detalle[col].round(2)
    return df_detalle


def generar_detalle_por_contrato(df_contratos, orden, df_inpc=None, mes_incremento=None):
    """
    Genera el detalle mensual para un contrato usando la clave orden.
    """
    if "orden" not in df_contratos.columns:
        raise ValueError("df_contratos no contiene columna 'orden'.")

    df = df_contratos.copy()
    df["orden_key"] = pd.to_numeric(df["orden"], errors="coerce")
    orden_key = pd.to_numeric(orden, errors="coerce")
    if pd.isna(orden_key):
        raise ValueError(f"Orden invalida: {orden}")

    df_con = df[df["orden_key"] == orden_key]
    if df_con.empty:
        raise ValueError(f"Contrato no encontrado para orden={orden}")
    if len(df_con) > 1:
        print(f"Aviso: orden={orden} aparece {len(df_con)} veces en contratos; se usara la primera fila.")

    row = df_con.iloc[0]

    fecha_firma = row.get("fecha_de_firma_del_contrato")
    fecha_terminacion = row.get("fecha_de_terminacion_del_contrato")
    fechas = generar_fechas_mensuales(fecha_firma, fecha_terminacion, fecha_max=_MAX_FECHA_REVISION)
    if len(fechas) == 0:
        raise ValueError(f"Fechas invalidas para orden={orden}: inicio={fecha_firma}, fin={fecha_terminacion}")

    df_detalle = pd.DataFrame({"fecha": fechas})
    df_detalle["ano"] = df_detalle["fecha"].dt.year
    df_detalle["mes"] = df_detalle["fecha"].dt.month
    df_detalle["orden"] = int(orden_key) if pd.notna(orden_key) else orden
    df_detalle["rfc"] = row.get("rfc_del_subarrendatario")
    df_detalle["subarrendatario"] = row.get("nombre_del_subarrendatario")
    df_detalle["area"] = row.get("superficie_del_inmueble")
    df_detalle["direccion_inmueble"] = row.get("direccion_del_inmueble")
    df_detalle["mes_incremento"] = mes_incremento

    renta = pd.to_numeric(row.get("monto_de_renta_mensual"), errors="coerce")
    cuota = pd.to_numeric(row.get("cuota_de_mantenimiento"), errors="coerce")
    renta = 0 if pd.isna(renta) else renta
    cuota = 0 if pd.isna(cuota) else cuota

    df_detalle["importe_renta_auditoria"] = calcular_renta_auditoria_con_inpc(
        df_detalle["fecha"], renta, df_inpc, mes_incremento=mes_incremento
    )
    df_detalle["cuota_mantenimiento"] = cuota

    df_detalle = calcular_subtotal_a(df_detalle)
    df_detalle = calcular_total_a(df_detalle)
    for col in ["importe_renta_auditoria", "importe_mtto_auditoria", "subtotal_a", "total_a"]:
        if col in df_detalle.columns:
            df_detalle[col] = df_detalle[col].round(2)
    return df_detalle


def generar_detalle_todos(df_contratos, df_clientes=None, df_inpc=None):
    """
    Genera detalle mensual para todos los contratos en df_contratos.
    Si se pasa df_clientes, cruza importes del cliente por ORDEN.
    """
    detalles = []
    saltados = 0
    mes_incremento_por_orden = _mapa_mes_incremento_por_orden(df_clientes)

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
                    generar_detalle_por_contrato(
                        df_contratos,
                        int(orden),
                        df_inpc=df_inpc,
                        mes_incremento=mes_incremento_por_orden.get(float(orden)),
                    )
                )
            except ValueError as exc:
                print(f"Saltando orden={orden}: {exc}")
                saltados += 1

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
                        generar_detalle_por_proveedor(
                            df_contratos,
                            sub,
                            rfc_sub=rfc,
                            df_inpc=df_inpc,
                            mes_incremento=None,
                        )
                    )
                except ValueError as exc:
                    print(f"Saltando {sub}: {exc}")
                    saltados += 1
    else:
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
                    generar_detalle_por_proveedor(
                        df_contratos,
                        sub,
                        rfc_sub=rfc,
                        df_inpc=df_inpc,
                        mes_incremento=None,
                    )
                )
            except ValueError as exc:
                print(f"Saltando {sub}: {exc}")
                saltados += 1

    if not detalles:
        return pd.DataFrame()

    detalle = pd.concat(detalles, ignore_index=True)
    if df_clientes is not None:
        detalle = anexar_importe_renta_mtto_clientes(detalle, df_clientes)
        detalle = aplicar_identidad_cliente_por_orden(detalle, df_clientes)
    detalle = agregar_incremento_constante(detalle, valor="INPC")

    orden = [
        "orden",
        "fecha",
        "ano",
        "mes",
        "mes_incremento",
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
        print(f"Total de contratos saltados por fechas invalidas o datos faltantes: {saltados}")
    return detalle
