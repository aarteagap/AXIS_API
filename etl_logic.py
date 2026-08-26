import pandas as pd, numpy as np, json, re

# ══════════════════════════════════════════════════════════════
# Bandera temporal: en False, la vista de Incidencias queda
# completamente desactivada (no se lee la hoja "Incidencias" del
# Excel, no se publica INCIDENCIAS/UNIVERSO_COMPLIANCE/HORIZON_RANGE
# a Supabase/GitHub). Se aísla así mientras se investiga el error
# "Failed to fetch" en /publish, para no arriesgar la carga de
# quienes ya usan el dashboard hoy. Cuando se confirme que Incidencias
# no era la causa (o se corrija lo que sea), cambiar a True.
# ══════════════════════════════════════════════════════════════
ENABLE_INCIDENCIAS = True


def build_dashboard_data(xlsm_path):
    import pandas as pd, numpy as np, json, re
    df = pd.read_excel(xlsm_path, sheet_name='HORIZON_FORECAST', engine='openpyxl')
    df = df[~df['Status'].isin(['Maquinaria', 'Traslado'])].copy()
    df['FCL'] = pd.to_numeric(df['FCL'], errors='coerce').fillna(0)
    df['Pack Plan'] = pd.to_numeric(df['Pack Plan'], errors='coerce')
    df = df[df['Pack Plan'].notna()].copy()
    df['Pack Plan'] = df['Pack Plan'].astype(int)

    def dstr(x):
        if pd.isna(x): return ""
        try:
            return pd.Timestamp(x).strftime('%Y-%m-%d')
        except Exception:
            return ""

    def tstr(x):
        if pd.isna(x): return ""
        try:
            if hasattr(x, 'strftime'): return x.strftime('%H:%M')
            total = int(x.total_seconds())
            return f"{total//3600:02d}:{(total%3600)//60:02d}"
        except Exception:
            return ""

    def dstr_ddmmyyyy(x):
        if pd.isna(x): return ""
        try:
            return pd.Timestamp(x).strftime('%d/%m/%Y')
        except Exception:
            return ""

    conf = df[df['Status'] == 'Confirmado'].copy()
    proy = df[df['Status'] == 'Proyectado'].copy()
    canc = df[df['Status'] == 'Cancelado'].copy()

    # ══════════════════════════ SENASA ══════════════════════════
    sen = conf.copy()
    SENASA = []
    for _, r in sen.iterrows():
        SENASA.append({
            "pack_plan": int(r['Pack Plan']),
            "packing": r['Packing'] if pd.notna(r['Packing']) else "",
            "mode": r['Mode'] if pd.notna(r['Mode']) else "",
            "pod": r['POD'] if pd.notna(r['POD']) else "",
            "senasa_att": r['SENASA Attention'] if pd.notna(r['SENASA Attention']) else "",
            "inspector": r['Assigned Inspector'] if pd.notna(r['Assigned Inspector']) else "",
            "insp_date": dstr(r['Inspection Date']),
            "insp_time": tstr(r['Inspection Time']),
            "fcl": round(float(r['FCL']), 2),
            "po": str(r['PO']) if pd.notna(r['PO']) else "",
            "dispatch_date": dstr_ddmmyyyy(r['Dispatch Date']),
        })

    # ══════════════════════════ WROWS (weekly export program) ══════════════════════════
    wr = df[df['Status'] == 'Confirmado'].copy()
    WROWS = []
    for _, r in wr.iterrows():
        WROWS.append({
            "pack_plan": int(r['Pack Plan']),
            "packing": r['Packing'] if pd.notna(r['Packing']) else "",
            "mode": r['Mode'] if pd.notna(r['Mode']) else "",
            "shipper": r['Shipper'] if pd.notna(r['Shipper']) else "",
            "pod": r['POD'] if pd.notna(r['POD']) else "",
            "date": dstr(r['Loading Date']),
            "fcl": round(float(r['FCL']), 2),
            "pallets": round(float(r['Pallets']), 1) if pd.notna(r['Pallets']) else 0.0,
        })

    # ══════════════════════════ AIR (aereo transport detail, confirmado) ══════════════════════════
    air = conf[conf['Mode'] == 'Aéreo'].copy()
    AIR = []
    for _, r in air.iterrows():
        wk = r['Pack Plan']
        AIR.append({
            "wk": f"S{int(wk)}",
            "pack_plan": int(wk),
            "packing": r['Packing'] if pd.notna(r['Packing']) else "",
            "po": str(r['PO']) if pd.notna(r['PO']) else "",
            "shipper": r['Shipper'] if pd.notna(r['Shipper']) else "",
            "instruction": r['Instruction'] if pd.notna(r['Instruction']) else "",
            "log_op": r['Logistics Operator'] if pd.notna(r['Logistics Operator']) else "",
            "dest": r['POD'] if pd.notna(r['POD']) else "",
            "dispatch": dstr_ddmmyyyy(r['Dispatch Date']),
            "postime": tstr(r['Positioning Time']),
            "gatein": r['Gate-In Storage'] if pd.notna(r['Gate-In Storage']) else "",
            "pallets": int(r['Pallets']) if pd.notna(r['Pallets']) else 0,
            "fcl": round(float(r['FCL']), 2) if pd.notna(r['FCL']) else 0.0,
            "awb": str(r['Booking|AWB|CRT']) if pd.notna(r['Booking|AWB|CRT']) else "",
            "file": str(r['File']) if pd.notna(r['File']) else "",
            "line": r['Line'] if pd.notna(r['Line']) else "",
            "transp_op": r['Transport Operator'] if pd.notna(r['Transport Operator']) else "",
            "basc": r['Journey Status BASC'] if pd.notna(r['Journey Status BASC']) else "",
            "dam": str(r['DAM']) if pd.notna(r['DAM']) else "",
            "driver": r['Driver'] if pd.notna(r['Driver']) else "",
            "license": r['License'] if pd.notna(r['License']) else "",
            "tractor": r['Tractor'] if pd.notna(r['Tractor']) else "",
            "trailer": r['Trailer'] if pd.notna(r['Trailer']) else "",
            "phone": str(r['Driver Phone Number']) if pd.notna(r['Driver Phone Number']) else "",
        })

    # ══════════════════════════ AIRLINES (PO diferenciados por línea, mode Aereo) ══════════════════════════
    air_po = air[air['PO'].notna()]
    al = air_po.groupby('Line')['PO'].nunique().sort_values(ascending=False)
    palette = ["#E8A33D","#2E8FB0","#163B54","#7A6FD0","#C0392B","#2E8B57","#C77D1E","#9CA3AF","#6B7280","#1C2940"]
    AIRLINES = [{"name": str(k), "count": int(v), "color": palette[i % len(palette)]} for i, (k, v) in enumerate(al.items())]

    # ══════════════════════════ FWKS / WKL ══════════════════════════
    # 6W field formats vary (e.g. "2227ADD", "PRE27", "2224PRE", bare "EXADD") — so instead of
    # strict positional parsing we detect each category by substring containment, checked in
    # priority order so overlapping tags (AIRADD/EXADD both contain "ADD") don't get double-counted.
    df['_6w_str'] = df['6W'].apply(lambda v: '' if pd.isna(v) else str(v).upper())
    mask_airadd = df['_6w_str'].str.contains('AIRADD', na=False)
    mask_exadd = df['_6w_str'].str.contains('EXPOADD', na=False) & ~mask_airadd
    mask_expo = df['_6w_str'].str.contains('EXPO', na=False) & ~mask_airadd & ~mask_exadd
    mask_add_plain = df['_6w_str'].str.contains('ADD', na=False) & ~mask_airadd & ~mask_exadd & ~mask_expo
    mask_pre = df['_6w_str'].str.contains('PRE', na=False)
    mask_status_cc = df['Status'].isin(['Confirmado', 'Cancelado'])
    # Orange projection line = FCL tagged PRE (anywhere in the 6W code) OR already Confirmado/Cancelado
    mask_proj = mask_pre | mask_status_cc

    weeks = sorted(df['Pack Plan'].dropna().unique().tolist())
    weeks = [int(w) for w in weeks]

    df['Pallets'] = pd.to_numeric(df['Pallets'], errors='coerce').fillna(0)

    FWKS, WKL = [], []
    for wk in weeks:
        dwk = df[df['Pack Plan'] == wk]
        m_wk = df['Pack Plan'] == wk
        c = dwk[dwk['Status'] == 'Confirmado']['FCL'].sum()
        cp = dwk[dwk['Status'] == 'Confirmado']['Pallets'].sum()
        canc_fcl = dwk[dwk['Status'] == 'Cancelado']['FCL'].sum()
        canc_fcl_p = dwk[dwk['Status'] == 'Cancelado']['Pallets'].sum()

        # All of the following are GENERAL counters (all statuses), not limited to Confirmado —
        # only "fcl_conf" above is status-filtered.
        p = df[m_wk & mask_proj]['FCL'].sum()
        fadd = df[m_wk & mask_add_plain]['FCL'].sum()
        fexpo = df[m_wk & mask_expo]['FCL'].sum()
        fexadd = df[m_wk & mask_exadd]['FCL'].sum()
        fairadd = df[m_wk & mask_airadd]['FCL'].sum()
        a = fadd + fexadd  # badge/total-projected addon total = ADD + EXADD (EXPO and AIRADD tracked separately)

        pp = df[m_wk & mask_proj]['Pallets'].sum()
        padd = df[m_wk & mask_add_plain]['Pallets'].sum()
        pexpo = df[m_wk & mask_expo]['Pallets'].sum()
        pexadd = df[m_wk & mask_exadd]['Pallets'].sum()
        pairadd = df[m_wk & mask_airadd]['Pallets'].sum()
        pa = padd + pexadd

        FWKS.append({"wk": wk, "label": f"S{wk}", "fcl_conf": round(float(c), 1),
                     "fcl_proj": round(float(p), 1), "fcl_addon": round(float(a), 1),
                     "fcl_add": round(float(fadd), 1), "fcl_expo": round(float(fexpo), 1), "fcl_exadd": round(float(fexadd), 1),
                     "fcl_airadd": round(float(fairadd), 1), "fcl_cancelado": round(float(canc_fcl), 1),
                     "pal_conf": round(float(cp), 1), "pal_proj": round(float(pp), 1), "pal_addon": round(float(pa), 1),
                     "pal_add": round(float(padd), 1), "pal_expo": round(float(pexpo), 1), "pal_exadd": round(float(pexadd), 1),
                     "pal_airadd": round(float(pairadd), 1), "pal_cancelado": round(float(canc_fcl_p), 1)})

        cm = dwk[(dwk['Status']=='Confirmado') & (dwk['Mode']=='Marítimo')]['FCL'].sum()
        ca = dwk[(dwk['Status']=='Confirmado') & (dwk['Mode']=='Aéreo')]['FCL'].sum()
        ct = dwk[(dwk['Status']=='Confirmado') & (dwk['Mode']=='Terrestre')]['FCL'].sum()
        WKL.append({"wk": wk, "label": f"S{wk}", "conf_mar": round(float(cm),1), "conf_aer": round(float(ca),1),
                    "conf_ter": round(float(ct),1), "tot_conf": round(float(cm+ca+ct),1), "tot_proy": round(float(p+a),1)})

    # ══════════════════════════ KPI ══════════════════════════
    total_recs = len(df)
    n_conf, n_proy, n_canc = len(conf), len(proy), len(canc)
    fcl_conf_total = conf['FCL'].sum()
    share_mar = conf[conf['Mode']=='Marítimo']['FCL'].sum() / fcl_conf_total * 100 if fcl_conf_total else 0
    share_aer = conf[conf['Mode']=='Aéreo']['FCL'].sum() / fcl_conf_total * 100 if fcl_conf_total else 0
    cob_dam = conf['DAM'].notna().sum() / n_conf * 100 if n_conf else 0
    KPI = {
        "tasa_conf": round(n_conf/total_recs*100, 1),
        "tasa_canc": round(n_canc/total_recs*100, 1),
        "share_mar": round(share_mar, 1),
        "share_aer": round(share_aer, 1),
        "cob_dam": round(cob_dam, 1),
        "total_recs": total_recs, "confirmados": n_conf, "proyectados": n_proy, "cancelados": n_canc,
    }

    # ══════════════════════════ EX (fcl + DAM coverage) ══════════════════════════
    fcl_aereo = conf[conf['Mode']=='Aéreo']['FCL'].sum()
    fcl_maritimo = conf[conf['Mode']=='Marítimo']['FCL'].sum()
    fcl_terrestre = conf[conf['Mode']=='Terrestre']['FCL'].sum()
    dam_unique = conf['DAM'].dropna().nunique()
    dam_missing = conf['DAM'].isna().sum()
    dam_by_mode = []
    for m in ['Aéreo','Marítimo','Terrestre']:
        sub = conf[conf['Mode']==m]
        w = sub['DAM'].notna().sum(); wo = sub['DAM'].isna().sum(); tot = len(sub)
        dam_by_mode.append({"mode": m, "dam_with": int(w), "dam_without": int(wo), "total": int(tot),
                             "pct": round(w/tot*100,1) if tot else 0.0})
    dam_by_pp = []
    for wk in weeks:
        sub = conf[conf['Pack Plan']==wk]
        if len(sub)==0: continue
        du = sub['DAM'].dropna().nunique(); sd = sub['DAM'].isna().sum()
        dam_by_pp.append({"label": f"S{int(wk)}", "dam_unicas": int(du), "sin_dam": int(sd), "total": int(du+sd)})
    EX = {
        "total_fcl": round(float(fcl_conf_total),1), "fcl_aereo": round(float(fcl_aereo),1),
        "fcl_maritimo": round(float(fcl_maritimo),1), "fcl_terrestre": round(float(fcl_terrestre),1),
        "dam_unique": int(dam_unique), "dam_missing": int(dam_missing),
        "dam_by_mode": dam_by_mode, "dam_by_pp": dam_by_pp,
    }

    # ══════════════════════════ DEST (top 10 destinos confirmado) ══════════════════════════
    dst = conf.groupby('POD')['FCL'].sum().sort_values(ascending=False).head(10)
    DEST = [{"port": str(k), "fcl": round(float(v),1)} for k, v in dst.items()]

    # ══════════════════════════ LINEAS ══════════════════════════
    lin = conf.groupby('Line')['FCL'].sum().sort_values(ascending=False).head(8)
    LINEAS = [{"line": str(k), "fcl": round(float(v),1)} for k, v in lin.items()]

    # ══════════════════════════ PMI_PORTS (Gate-Out Port = origen Peru) ══════════════════════════
    PMI_PORTS = []
    colors = {"Lima":"#E8A33D","Callao":"#2E8FB0","Chancay":"#163B54","Paita":"#7A6FD0"}
    for port in ["Lima","Callao","Chancay","Paita"]:
        sub = df[df['Gate-Out Port']==port]
        if len(sub)==0: continue
        c = (sub['Status']=='Confirmado').sum(); p = (sub['Status']=='Proyectado').sum(); k = (sub['Status']=='Cancelado').sum()
        tot = len(sub)
        cumpl = round(c/tot*100,1) if tot else 0.0
        fcl = sub[sub['Status']=='Confirmado']['FCL'].sum()
        modes = sub['Mode'].value_counts()
        mode = modes.idxmax() if len(modes) else ""
        PMI_PORTS.append({"port": port, "conf": int(c), "proy": int(p), "canc": int(k), "total": int(tot),
                           "cumpl": cumpl, "fcl": round(float(fcl),1), "mode": mode, "color": colors.get(port,"#888"),
                           "delta_conf": 0, "delta_cumpl": 0.0})

    # ══════════════════════════════════════════════════════════════
    # FWKS_MODE — same as FWKS but split by transport Mode, for the 3 extra forecast charts
    # ══════════════════════════════════════════════════════════════
    FWKS_MODE = {}
    for mode in ['Aéreo', 'Marítimo', 'Terrestre']:
        dmode = df[df['Mode'] == mode]
        m_mode = df['Mode'] == mode
        rows = []
        for wk in weeks:
            dwk = dmode[dmode['Pack Plan'] == wk]
            m_wk_mode = m_mode & (df['Pack Plan'] == wk)
            c = dwk[dwk['Status'] == 'Confirmado']['FCL'].sum()
            cp = dwk[dwk['Status'] == 'Confirmado']['Pallets'].sum()
            canc_fcl = dwk[dwk['Status'] == 'Cancelado']['FCL'].sum()
            canc_fcl_p = dwk[dwk['Status'] == 'Cancelado']['Pallets'].sum()
            p = df[m_wk_mode & mask_proj]['FCL'].sum()
            fadd = df[m_wk_mode & mask_add_plain]['FCL'].sum()
            fexpo = df[m_wk_mode & mask_expo]['FCL'].sum()
            fexadd = df[m_wk_mode & mask_exadd]['FCL'].sum()
            fairadd = df[m_wk_mode & mask_airadd]['FCL'].sum()
            a = fadd + fexadd
            pp = df[m_wk_mode & mask_proj]['Pallets'].sum()
            padd = df[m_wk_mode & mask_add_plain]['Pallets'].sum()
            pexpo = df[m_wk_mode & mask_expo]['Pallets'].sum()
            pexadd = df[m_wk_mode & mask_exadd]['Pallets'].sum()
            pairadd = df[m_wk_mode & mask_airadd]['Pallets'].sum()
            pa = padd + pexadd
            rows.append({"wk": wk, "label": f"S{wk}", "fcl_conf": round(float(c), 1),
                         "fcl_proj": round(float(p), 1), "fcl_addon": round(float(a), 1),
                         "fcl_add": round(float(fadd), 1), "fcl_expo": round(float(fexpo), 1), "fcl_exadd": round(float(fexadd), 1),
                         "fcl_airadd": round(float(fairadd), 1), "fcl_cancelado": round(float(canc_fcl), 1),
                         "pal_conf": round(float(cp), 1), "pal_proj": round(float(pp), 1), "pal_addon": round(float(pa), 1),
                         "pal_add": round(float(padd), 1), "pal_expo": round(float(pexpo), 1), "pal_exadd": round(float(pexadd), 1),
                         "pal_airadd": round(float(pairadd), 1), "pal_cancelado": round(float(canc_fcl_p), 1)})
        FWKS_MODE[mode] = rows

    # ══════════════════════════════════════════════════════════════
    # EMBARQUES — "Control de Embarques" (PMI tab), Status=Confirmado only
    # ══════════════════════════════════════════════════════════════
    def tstr2(x):
        if pd.isna(x): return ""
        try:
            if hasattr(x, 'strftime'): return x.strftime('%H:%M')
            total = int(x.total_seconds())
            return f"{total//3600:02d}:{(total%3600)//60:02d}"
        except Exception:
            return ""

    emb = conf.copy()
    EMBARQUES = []
    for _, r in emb.iterrows():
        EMBARQUES.append({
            "round": r['Dispatch Round'] if pd.notna(r['Dispatch Round']) else "(sin ronda)",
            "pack_plan": int(r['Pack Plan']),
            "load_date": dstr(r['Loading Date']),
            "load_time": tstr2(r['Loading Time']),
            "pack_est": (r['Packing Stage Estimated Arrival Time'].strftime('%d/%m/%Y %H:%M') if hasattr(r['Packing Stage Estimated Arrival Time'], 'strftime') else str(r['Packing Stage Estimated Arrival Time'])) if pd.notna(r['Packing Stage Estimated Arrival Time']) else "",
            "basc": r['Journey Status BASC'] if pd.notna(r['Journey Status BASC']) else "",
            "load_status": r['Packing Stage Load Status'] if pd.notna(r['Packing Stage Load Status']) else "",
            "instruction": r['Instruction'] if pd.notna(r['Instruction']) else "",
            "po": str(r['PO']) if pd.notna(r['PO']) else "",
            "line": r['Line'] if pd.notna(r['Line']) else "",
            "log_op": r['Logistics Operator'] if pd.notna(r['Logistics Operator']) else "",
            "transp_op": r['Transport Operator'] if pd.notna(r['Transport Operator']) else "",
            "booking": str(r['Booking|AWB|CRT']) if pd.notna(r['Booking|AWB|CRT']) else "",
            "container": str(r['Container']) if pd.notna(r['Container']) else "",
            "pod": r['POD'] if pd.notna(r['POD']) else "",
            "packing": r['Packing'] if pd.notna(r['Packing']) else "",
            "gatein": r['Gate-In Storage'] if pd.notna(r['Gate-In Storage']) else "",
            "gatein_address": r['Gate-In Storage Address'] if pd.notna(r['Gate-In Storage Address']) else "",
            "postime": tstr2(r['Positioning Time']),
            "dispatch_date": dstr_ddmmyyyy(r['Dispatch Date']),
            "mode": r['Mode'] if pd.notna(r['Mode']) else "",
            "dam": str(r['DAM']) if pd.notna(r['DAM']) else "",
            "fcl": round(float(r['FCL']), 2) if pd.notna(r['FCL']) else 0.0,
            "retiro_cita": bool(pd.notna(r['Gate-Out Stage Appointment'])),
        })

    # ══════════════════════════════════════════════════════════════
    # TR_PROGRAM — "Programa de Técnicos Reefer" (PMI tab), Status=Confirmado
    # ══════════════════════════════════════════════════════════════
    tr = conf[conf['T.R. Time'].notna() | conf['Assigned T.R.'].notna()].copy()
    TR_PROGRAM = []
    for _, r in tr.iterrows():
        TR_PROGRAM.append({
            "packing": r['Packing'] if pd.notna(r['Packing']) else "",
            "pack_plan": int(r['Pack Plan']),
            "mode": r['Mode'] if pd.notna(r['Mode']) else "",
            "dispatch_date": dstr_ddmmyyyy(r['Dispatch Date']),
            "log_op": r['Logistics Operator'] if pd.notna(r['Logistics Operator']) else "",
            "senasa_att": r['SENASA Attention'] if pd.notna(r['SENASA Attention']) else "",
            "line": r['Line'] if pd.notna(r['Line']) else "",
            "insp_time": tstr2(r['Inspection Time']),
            "tr_time": tstr2(r['T.R. Time']),
            "assigned_tr": r['Assigned T.R.'] if pd.notna(r['Assigned T.R.']) else "",
            "booking": str(r['Booking|AWB|CRT']) if pd.notna(r['Booking|AWB|CRT']) else "",
            "inspector": r['Assigned Inspector'] if pd.notna(r['Assigned Inspector']) else "",
        })

    # ══════════════════════════ PORTS_RAW (row-level, for dynamic Status/Pack Plan filters) ══════════════════════════
    # NOTE: uses POL (Port of Loading), not Gate-Out Port — POL has much better fill rate and is the
    # field that correctly represents the origin gateway for BOTH Aéreo (airport) and Marítimo (seaport)
    # shipments alike; Gate-Out Port is populated almost exclusively for Marítimo.
    PORTS_RAW = []
    for _, r in df.iterrows():
        if pd.isna(r['Status']):
            continue
        port = r['POL'] if pd.notna(r['POL']) else ""
        PORTS_RAW.append({
            "port": port,
            "status": r['Status'],
            "pack_plan": int(r['Pack Plan']),
            "fcl": round(float(r['FCL']), 2),
            "mode": r['Mode'] if pd.notna(r['Mode']) else "",
            "po": str(r['PO']) if pd.notna(r['PO']) else "",
            "line": r['Line'] if pd.notna(r['Line']) else "",
        })

    # ══════════════════════════════════════════════════════════════
    # FORECAST_RAW — row-level data so the Forecast tab can filter by Shipper/Modo/Pack Plan
    # on the client, then re-aggregate into weekly totals on the fly.
    # ══════════════════════════════════════════════════════════════
    FORECAST_RAW = []
    df_fc = df[df['Status'].isin(['Confirmado', 'Proyectado', 'Cancelado'])]
    for _, r in df_fc.iterrows():
        s6w = '' if pd.isna(r['6W']) else str(r['6W']).upper()
        is_airadd = 'AIRADD' in s6w
        is_exadd = ('EXPOADD' in s6w) and not is_airadd
        is_expo = ('EXPO' in s6w) and not is_airadd and not is_exadd
        is_add = ('ADD' in s6w) and not is_airadd and not is_exadd and not is_expo
        is_pre = 'PRE' in s6w
        status = r['Status']
        FORECAST_RAW.append({
            "pack_plan": int(r['Pack Plan']),
            "mode": r['Mode'] if pd.notna(r['Mode']) else "",
            "shipper": r['Shipper'] if pd.notna(r['Shipper']) else "",
            "pod": r['POD'] if pd.notna(r['POD']) else "",
            "status": status,
            "fcl": round(float(r['FCL']), 2) if pd.notna(r['FCL']) else 0.0,
            "pallets": round(float(r['Pallets']), 1) if pd.notna(r['Pallets']) else 0.0,
            "is_pre": bool(is_pre), "is_add": bool(is_add), "is_expo": bool(is_expo),
            "is_exadd": bool(is_exadd), "is_airadd": bool(is_airadd),
        })

    # ══════════════════════════════════════════════════════════════
    # META — publish metadata (when the source Excel was last modified)
    # ══════════════════════════════════════════════════════════════
    try:
        from openpyxl import load_workbook as _load_wb_meta
        from datetime import timezone as _tz
        _wb_props = _load_wb_meta(xlsm_path, read_only=True).properties
        if _wb_props.modified:
            _dt = _wb_props.modified
            # OOXML core properties store dates in UTC; openpyxl returns a naive datetime for
            # this value, so without explicitly marking it as UTC, JS in the browser would
            # misinterpret it as already being local time (shifting it by the UTC offset).
            if _dt.tzinfo is None:
                _dt = _dt.replace(tzinfo=_tz.utc)
            _excel_modified = _dt.isoformat()
        else:
            _excel_modified = None
    except Exception:
        _excel_modified = None

    META = {"excel_modified_iso": _excel_modified}

    # ══════════════════════════════════════════════════════════════
    # INCIDENCIAS — vista nueva de Incident Management (solo perfil
    # hortifrut). Fuente: hoja "Incidencias" del mismo Excel, cruzada
    # por ID_PO (que puede venir como Instruction estilo "RH-0133T" o
    # como PO estilo "PO00000051363") contra HORIZON_FORECAST para
    # traer Mode / Pack Plan / Packing / FCL / Pallets reales.
    # Replica la lógica de los 4 dashboards de Tableau que ya existían
    # (Incidents, General Compliance, Marítimo, Aéreo):
    #   Incumplimiento Marítimo % = COUNTD(Instruction con incidencia) / COUNTD(Instruction total)
    #   Incumplimiento Aéreo %    = COUNTD(PO con incidencia) / COUNTD(PO total)
    #   General Compliance %     = Σ FCL con incidencia / Σ FCL total
    # ══════════════════════════════════════════════════════════════
    # INCIDENCIAS — vista de Incident Management (solo perfil hortifrut).
    # Fuente: hoja "Incidencias" del mismo Excel, cruzada por ID_PO
    # (puede venir como Instruction estilo "RH-0133T" o como PO estilo
    # "PO00000051363") contra HORIZON_FORECAST para traer Mode / Pack
    # Plan / Packing / FCL / Pallets reales.
    #
    # Nota de robustez (v2): reescrito para leer columnas POR NOMBRE
    # (no por posición .iloc[N]) — la hoja "Incidencias" tiene 60
    # columnas, muchas de ellas tablas auxiliares repetidas más allá
    # de la columna 13 (un segundo "Criticidad" en la 47, un segundo
    # "Incident" en la 45, etc.). Leer por posición es fráril: si se
    # inserta o reordena una sola columna en el Excel, los índices se
    # corren en silencio. También se unificaron los 5 loops separados
    # sobre HORIZON_FORECAST en una sola pasada, más liviano para el
    # servidor de Render.
    #
    # Replica la lógica de los 4 dashboards de Tableau que ya existían
    # (Incidents, General Compliance, Marítimo, Aéreo):
    #   Incumplimiento Marítimo % = COUNTD(Instruction con incidencia) / COUNTD(Instruction total)
    #   Incumplimiento Aéreo %    = COUNTD(PO con incidencia) / COUNTD(PO total)
    #   General Compliance %     = Σ FCL con incidencia / Σ FCL total
    # ══════════════════════════════════════════════════════════════
    INCIDENCIAS = []
    UNIVERSO_COMPLIANCE = {"marítimo": {"instructions": 0, "fcl_total": 0.0, "pallets_total": 0.0},
                            "aéreo": {"pos": 0, "fcl_total": 0.0, "pallets_total": 0.0},
                            "terrestre": {"instructions": 0, "fcl_total": 0.0, "pallets_total": 0.0}}
    if ENABLE_INCIDENCIAS:
      try:
        inc_df = pd.read_excel(xlsm_path, sheet_name='Incidencias', engine='openpyxl', header=0)
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
            fcl = float(r['FCL']) if pd.notna(r['FCL']) else 0.0
            pallets = float(r['Pallets']) if pd.notna(r['Pallets']) else 0.0
            pp = int(r['Pack Plan']) if pd.notna(r['Pack Plan']) else None

            if mode == 'Marítimo':
                if instr: marit_ids.add(instr)
                marit_fcl_total += fcl
                marit_pallets_total += pallets
                if pp is not None and instr:
                    marit_by_pp.setdefault(pp, set()).add(instr)
            elif mode == 'Aéreo':
                if po: aereo_pos.add(po)
                aereo_fcl_total += fcl
                aereo_pallets_total += pallets
                if pp is not None and po:
                    aereo_by_pp.setdefault(pp, set()).add(po)
            elif mode == 'Terrestre':
                # Igual criterio que Marítimo: 1 Instruction = 1 camión = 1 pedido.
                if instr: terr_ids.add(instr)
                terr_fcl_total += fcl
                terr_pallets_total += pallets
                if pp is not None and instr:
                    terr_by_pp.setdefault(pp, set()).add(instr)

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
                "persona": r['PERSONA'] if pd.notna(r['PERSONA']) else "",
                "id_po": id_po_str or "",
                "dispatch_date": dstr_ddmmyyyy(r['Dispatch Date']) if pd.notna(r['Dispatch Date']) else "",
                "t_supplier": r['T_Supplier'] if pd.notna(r['T_Supplier']) else "",
                "supplier": r['Supplier'] if pd.notna(r['Supplier']) else "",
                "incident": r['Incident'] if pd.notna(r['Incident']) else "",
                "criticidad": criticidad,
                "value_range": r['Value'] if pd.notna(r['Value']) else "",
                "comentario": r['Comentario'] if pd.notna(r['Comentario']) else "",
                "mode": (match['Mode'] if pd.notna(match['Mode']) else None),
                "instruction": (str(match['Instruction']) if pd.notna(match['Instruction']) else None),
                "po": (str(match['PO']) if pd.notna(match['PO']) else None),
                "pack_plan": (int(match['Pack Plan']) if pd.notna(match['Pack Plan']) else None),
                "packing": (match['Packing'] if pd.notna(match['Packing']) else None),
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
        }
        UNIVERSO_COMPLIANCE["aéreo"] = {
            "pos": len(aereo_pos),
            "fcl_total": round(aereo_fcl_total, 2),
            "pallets_total": round(aereo_pallets_total, 1),
            "by_pack_plan": {pp: len(s) for pp, s in aereo_by_pp.items()},
        }
        UNIVERSO_COMPLIANCE["terrestre"] = {
            "instructions": len(terr_ids),
            "fcl_total": round(terr_fcl_total, 2),
            "pallets_total": round(terr_pallets_total, 1),
            "by_pack_plan": {pp: len(s) for pp, s in terr_by_pp.items()},
        }
        UNIVERSO_COMPLIANCE["general"] = {
            "by_pack_plan_instruction": {pp: len(s) for pp, s in gen_instr_by_pp.items()},
            "by_pack_plan_po": {pp: len(s) for pp, s in gen_po_by_pp.items()},
            "by_pack_plan_fcl": {pp: round(v, 2) for pp, v in gen_fcl_by_pp.items()},
            "total_instruction": len(gen_instr_all),
            "total_po": len(gen_po_all),
            "total_fcl": round(gen_fcl_total, 2),
        }
      except Exception:
        # Si la hoja "Incidencias" no existe en este Excel, o le faltan
        # columnas esperadas (algunos snapshots de prueba no la traen tal
        # cual), la vista de Incidencias del dashboard queda simplemente
        # vacía en vez de romper la carga completa de datos.
        pass

    # EXECUTE_RAW — row-level Confirmado data for the Execute tab's global filters,
    # KPI cards (Cold treatments / Senasa Attention / Booking counts) and the
    # PO-diferenciado-by-mode line chart.
    # ══════════════════════════════════════════════════════════════
    EXECUTE_RAW = []
    for _, r in conf.iterrows():
        EXECUTE_RAW.append({
            "pack_plan": int(r['Pack Plan']),
            "mode": r['Mode'] if pd.notna(r['Mode']) else "",
            "shipper": r['Shipper'] if pd.notna(r['Shipper']) else "",
            "dispatch_date": dstr_ddmmyyyy(r['Dispatch Date']),
            "po": str(r['PO']) if pd.notna(r['PO']) else "",
            "booking": str(r['Booking|AWB|CRT']) if pd.notna(r['Booking|AWB|CRT']) else "",
            "has_tr_time": bool(pd.notna(r['T.R. Time'])),
            "senasa_att": r['SENASA Attention'] if pd.notna(r['SENASA Attention']) else "",
        })

    # ══════════════════════════════════════════════════════════════
    # HORIZON_RANGE — todos los Pack Plan y meses (Dispatch Date) que
    # existen en HORIZON_FORECAST (Confirmado/Proyectado/Cancelado),
    # sin importar si tienen o no una incidencia asociada. Lo usa la
    # vista de Incidencias para dibujar el eje X completo (todos los
    # Pack Plan / todos los meses), no solo los que tuvieron un evento.
    # ══════════════════════════════════════════════════════════════
    all_pack_plans = sorted(int(x) for x in df_fc['Pack Plan'].dropna().unique())
    dispatch_dates_valid = df_fc['Dispatch Date'].dropna()
    all_months = sorted(set(pd.Timestamp(d).strftime('%Y-%m') for d in dispatch_dates_valid if pd.notna(d)))
    HORIZON_RANGE = {"pack_plans": all_pack_plans, "months": all_months}

    out = dict(SENASA=SENASA, WROWS=WROWS, AIR=AIR, AIRLINES=AIRLINES, FWKS=FWKS, WKL=WKL,
               KPI=KPI, EX=EX, DEST=DEST, LINEAS=LINEAS, PMI_PORTS=PMI_PORTS, FWKS_MODE=FWKS_MODE,
               EMBARQUES=EMBARQUES, TR_PROGRAM=TR_PROGRAM, PORTS_RAW=PORTS_RAW,
               FORECAST_RAW=FORECAST_RAW, META=META, EXECUTE_RAW=EXECUTE_RAW,
               INCIDENCIAS=INCIDENCIAS, UNIVERSO_COMPLIANCE=UNIVERSO_COMPLIANCE, HORIZON_RANGE=HORIZON_RANGE)

    return out
