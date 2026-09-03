"""
ETL — Incident Management (vista "Incidencias", solo perfil hortifrut)

Separado del /publish principal (etl_logic.py) para que subir el Excel de
Horizon no tenga que además parsear la hoja "Incidencias" y cruzarla contra
todo HORIZON_FORECAST en la misma petición — esa pasada extra era peso
adicional (memoria + tiempo) en cada publish del Excel principal, incluso
para quienes no tocan Incidencias ese día. Ahora tiene su propio botón /
endpoint (/publish-incidencias), sobre el MISMO archivo Excel de Horizon,
pero solo corre este procesamiento cuando alguien lo pide explícitamente.

Fuente: hoja "Incidencias" del Excel de Horizon, cruzada por ID_PO (puede
venir como Instruction estilo "RH-0133T" o como PO estilo "PO00000051363")
contra HORIZON_FORECAST para traer Mode / Pack Plan / Packing / FCL /
Pallets / Line reales.

Replica la lógica de los 4 dashboards de Tableau que ya existían (Incidents,
General Compliance, Marítimo, Aéreo):
  Incumplimiento Marítimo % = COUNTD(Instruction con incidencia) / COUNTD(Instruction total)
  Incumplimiento Aéreo %    = COUNTD(PO con incidencia) / COUNTD(PO total)
  General Compliance %     = Σ FCL con incidencia / Σ FCL total
"""
import pandas as pd


def _dstr_ddmmyyyy(x):
    if pd.isna(x): return ""
    try:
        return pd.Timestamp(x).strftime('%d/%m/%Y')
    except Exception:
        return ""


def build_incidencias_data(xlsm_path):
    xls = pd.ExcelFile(xlsm_path, engine='openpyxl')
    df = pd.read_excel(xls, sheet_name='HORIZON_FORECAST')
    df = df[~df['Status'].isin(['Maquinaria', 'Traslado'])].copy()
    df['FCL'] = pd.to_numeric(df['FCL'], errors='coerce').fillna(0)
    df['Pack Plan'] = pd.to_numeric(df['Pack Plan'], errors='coerce')
    df = df[df['Pack Plan'].notna()].copy()
    df['Pack Plan'] = df['Pack Plan'].astype(int)
    df_fc = df[df['Status'].isin(['Confirmado', 'Proyectado', 'Cancelado'])]

    INCIDENCIAS = []
    UNIVERSO_COMPLIANCE = {"marítimo": {"instructions": 0, "fcl_total": 0.0, "pallets_total": 0.0},
                            "aéreo": {"pos": 0, "fcl_total": 0.0, "pallets_total": 0.0},
                            "terrestre": {"instructions": 0, "fcl_total": 0.0, "pallets_total": 0.0}}

    inc_df = pd.read_excel(xls, sheet_name='Incidencias', header=0)
    # Solo las primeras 14 columnas son datos reales de incidencias (el resto,
    # a partir de la col. 14, son tablas auxiliares/listas de referencia que
    # comparten la misma hoja). Nos quedamos solo con las que usamos, por
    # NOMBRE — inmune a que se reordenen o se agreguen columnas nuevas.
    inc_cols = ['ID_PO', 'PERSONA', 'ID_Incident', 'Dispatch Date', 'T_Supplier',
                'Supplier', 'Incident', 'Criticidad', 'Value', 'Comentario', 'Filial']
    missing = [c for c in inc_cols if c not in inc_df.columns]
    if missing:
        raise KeyError(f"Faltan columnas esperadas en la hoja Incidencias: {missing}")
    inc_df = inc_df[inc_df['ID_Incident'].notna() & inc_df['Filial'].notna()].copy()

    # Un solo recorrido sobre HORIZON_FORECAST: construye el índice de cruce
    # (por Instruction y por PO) Y todos los agregados de universo a la vez,
    # en vez de 5 loops separados.
    fc_index = {}
    marit_ids, aereo_pos, terr_ids = set(), set(), set()
    marit_by_pp, aereo_by_pp, terr_by_pp = {}, {}, {}
    # by_line — universo (denominador) de "Fulfillment of the Assignment" por proveedor:
    # cuántas Instructions/PO distintas tiene asignadas cada Line (naviera/aerolínea/
    # transportista), para comparar contra cuántas de esas tuvieron incidencia.
    marit_by_line, aereo_by_line, terr_by_line = {}, {}, {}
    marit_fcl_total, aereo_fcl_total, terr_fcl_total = 0.0, 0.0, 0.0
    marit_pallets_total, aereo_pallets_total, terr_pallets_total = 0.0, 0.0, 0.0
    gen_instr_by_pp, gen_po_by_pp, gen_fcl_by_pp = {}, {}, {}
    gen_instr_all, gen_po_all = set(), set()
    gen_fcl_total = 0.0

    for _, r in df_fc.iterrows():
        instr = str(r['Instruction']) if pd.notna(r['Instruction']) else None
        po = str(r['PO']) if pd.notna(r['PO']) else None
        if instr: fc_index[instr] = r
        if po: fc_index[po] = r

        mode = r['Mode'] if pd.notna(r['Mode']) else None
        line = str(r['Line']) if pd.notna(r['Line']) else None
        fcl = float(r['FCL']) if pd.notna(r['FCL']) else 0.0
        pallets = float(r['Pallets']) if pd.notna(r['Pallets']) else 0.0
        pp = int(r['Pack Plan']) if pd.notna(r['Pack Plan']) else None

        if mode == 'Marítimo':
            if instr: marit_ids.add(instr)
            marit_fcl_total += fcl
            marit_pallets_total += pallets
            if pp is not None and instr:
                marit_by_pp.setdefault(pp, set()).add(instr)
            if line and instr:
                marit_by_line.setdefault(line, set()).add(instr)
        elif mode == 'Aéreo':
            if po: aereo_pos.add(po)
            aereo_fcl_total += fcl
            aereo_pallets_total += pallets
            if pp is not None and po:
                aereo_by_pp.setdefault(pp, set()).add(po)
            if line and po:
                aereo_by_line.setdefault(line, set()).add(po)
        elif mode == 'Terrestre':
            # Igual criterio que Marítimo: 1 Instruction = 1 camión = 1 pedido.
            if instr: terr_ids.add(instr)
            terr_fcl_total += fcl
            terr_pallets_total += pallets
            if pp is not None and instr:
                terr_by_pp.setdefault(pp, set()).add(instr)
            if line and instr:
                terr_by_line.setdefault(line, set()).add(instr)

        if instr: gen_instr_all.add(instr)
        if po: gen_po_all.add(po)
        gen_fcl_total += fcl
        if pp is not None:
            if instr: gen_instr_by_pp.setdefault(pp, set()).add(instr)
            if po: gen_po_by_pp.setdefault(pp, set()).add(po)
            gen_fcl_by_pp[pp] = gen_fcl_by_pp.get(pp, 0.0) + fcl

    for _, r in inc_df.iterrows():
        id_po = r['ID_PO']
        id_po_str = str(id_po).strip() if pd.notna(id_po) and str(id_po).strip() not in ('', 'Total') else None
        match = fc_index.get(id_po_str) if id_po_str else None
        if match is None:
            # No cruza con HORIZON_FORECAST de la temporada actual — es un
            # ID_PO de una temporada pasada (el PO es la clave primaria
            # contra Horizon). Se descarta por completo, no se muestra.
            continue

        criticidad_raw = r['Criticidad']
        try:
            criticidad = int(criticidad_raw) if pd.notna(criticidad_raw) and str(criticidad_raw).strip() != '' else None
        except (ValueError, TypeError):
            criticidad = None

        INCIDENCIAS.append({
            "id_incident": int(r['ID_Incident']),
            # str(...) explícito en estos campos: a diferencia de las columnas
            # estructuradas de HORIZON_FORECAST, la hoja "Incidencias" se llena
            # a mano y puede traer un número o una fecha en una columna de texto
            # (p.ej. "Value" o "PERSONA"). Sin el cast, ese valor llega como
            # numpy.int64/Timestamp, que requests.post(json=...) no sabe serializar
            # y revienta la publicación entera a Supabase con un 500 no controlado.
            "persona": str(r['PERSONA']) if pd.notna(r['PERSONA']) else "",
            "id_po": id_po_str or "",
            "dispatch_date": _dstr_ddmmyyyy(r['Dispatch Date']) if pd.notna(r['Dispatch Date']) else "",
            "t_supplier": str(r['T_Supplier']) if pd.notna(r['T_Supplier']) else "",
            "supplier": str(r['Supplier']) if pd.notna(r['Supplier']) else "",
            "incident": str(r['Incident']) if pd.notna(r['Incident']) else "",
            "criticidad": criticidad,
            "value_range": str(r['Value']) if pd.notna(r['Value']) else "",
            "comentario": str(r['Comentario']) if pd.notna(r['Comentario']) else "",
            "mode": (match['Mode'] if pd.notna(match['Mode']) else None),
            "line": (str(match['Line']) if pd.notna(match['Line']) else None),
            "instruction": (str(match['Instruction']) if pd.notna(match['Instruction']) else None),
            "po": (str(match['PO']) if pd.notna(match['PO']) else None),
            "pack_plan": (int(match['Pack Plan']) if pd.notna(match['Pack Plan']) else None),
            "packing": (match['Packing'] if pd.notna(match['Packing']) else None),
            "shipper": (match['Shipper'] if pd.notna(match['Shipper']) else None),
            "pod": (match['POD'] if pd.notna(match['POD']) else None),
            "fcl": (round(float(match['FCL']), 2) if pd.notna(match['FCL']) else None),
            "pallets": (round(float(match['Pallets']), 1) if pd.notna(match['Pallets']) else None),
            "matched": True,
        })

    UNIVERSO_COMPLIANCE["marítimo"] = {
        "instructions": len(marit_ids),
        "fcl_total": round(marit_fcl_total, 2),
        "pallets_total": round(marit_pallets_total, 1),
        "by_pack_plan": {pp: len(s) for pp, s in marit_by_pp.items()},
        "by_line": {line: len(s) for line, s in marit_by_line.items()},
    }
    UNIVERSO_COMPLIANCE["aéreo"] = {
        "pos": len(aereo_pos),
        "fcl_total": round(aereo_fcl_total, 2),
        "pallets_total": round(aereo_pallets_total, 1),
        "by_pack_plan": {pp: len(s) for pp, s in aereo_by_pp.items()},
        "by_line": {line: len(s) for line, s in aereo_by_line.items()},
    }
    UNIVERSO_COMPLIANCE["terrestre"] = {
        "instructions": len(terr_ids),
        "fcl_total": round(terr_fcl_total, 2),
        "pallets_total": round(terr_pallets_total, 1),
        "by_pack_plan": {pp: len(s) for pp, s in terr_by_pp.items()},
        "by_line": {line: len(s) for line, s in terr_by_line.items()},
    }
    UNIVERSO_COMPLIANCE["general"] = {
        "by_pack_plan_instruction": {pp: len(s) for pp, s in gen_instr_by_pp.items()},
        "by_pack_plan_po": {pp: len(s) for pp, s in gen_po_by_pp.items()},
        "by_pack_plan_fcl": {pp: round(v, 2) for pp, v in gen_fcl_by_pp.items()},
        "total_instruction": len(gen_instr_all),
        "total_po": len(gen_po_all),
        "total_fcl": round(gen_fcl_total, 2),
    }

    # HORIZON_RANGE — todos los Pack Plan y meses (Dispatch Date) que existen
    # en HORIZON_FORECAST (Confirmado/Proyectado/Cancelado), sin importar si
    # tienen o no una incidencia asociada. Dibuja el eje X completo de la
    # vista de Incidencias (todos los Pack Plan / todos los meses), no solo
    # los que tuvieron un evento.
    all_pack_plans = sorted(int(x) for x in df_fc['Pack Plan'].dropna().unique())
    dispatch_dates_valid = df_fc['Dispatch Date'].dropna()
    all_months = sorted(set(pd.Timestamp(d).strftime('%Y-%m') for d in dispatch_dates_valid if pd.notna(d)))
    HORIZON_RANGE = {"pack_plans": all_pack_plans, "months": all_months}

    xls.close()
    import gc
    gc.collect()

    return dict(INCIDENCIAS=INCIDENCIAS, UNIVERSO_COMPLIANCE=UNIVERSO_COMPLIANCE, HORIZON_RANGE=HORIZON_RANGE)
