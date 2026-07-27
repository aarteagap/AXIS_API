import pandas as pd, numpy as np, json, re


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
        return pd.Timestamp(x).strftime('%Y-%m-%d')

    def tstr(x):
        if pd.isna(x): return ""
        if hasattr(x, 'strftime'): return x.strftime('%H:%M')
        try:
            total = int(x.total_seconds())
            return f"{total//3600:02d}:{(total%3600)//60:02d}"
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
            "dispatch_date": pd.Timestamp(r['Dispatch Date']).strftime('%d/%m/%Y') if pd.notna(r['Dispatch Date']) else "",
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
            "dispatch": pd.Timestamp(r['Dispatch Date']).strftime('%d/%m/%Y') if pd.notna(r['Dispatch Date']) else "",
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
    mask_exadd = df['_6w_str'].str.contains('EXADD', na=False) & ~mask_airadd
    mask_expo = df['_6w_str'].str.contains('EXPO', na=False) & ~mask_airadd & ~mask_exadd
    mask_add_plain = df['_6w_str'].str.contains('ADD', na=False) & ~mask_airadd & ~mask_exadd
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
        if hasattr(x, 'strftime'): return x.strftime('%H:%M')
        try:
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
            "postime": tstr2(r['Positioning Time']),
            "dispatch_date": pd.Timestamp(r['Dispatch Date']).strftime('%d/%m/%Y') if pd.notna(r['Dispatch Date']) else "",
            "mode": r['Mode'] if pd.notna(r['Mode']) else "",
            "dam": str(r['DAM']) if pd.notna(r['DAM']) else "",
            "fcl": round(float(r['FCL']), 2) if pd.notna(r['FCL']) else 0.0,
            "retiro_cita": bool(pd.notna(r['Gate-Out Stage Appointment'])),
        })

    # ══════════════════════════════════════════════════════════════
    # TR_PROGRAM — "Programa de Técnicos Reefer" (PMI tab), Status=Confirmado
    # ══════════════════════════════════════════════════════════════
    tr = conf[conf['Assigned T.R.'].notna()].copy()
    TR_PROGRAM = []
    for _, r in tr.iterrows():
        TR_PROGRAM.append({
            "packing": r['Packing'] if pd.notna(r['Packing']) else "",
            "pack_plan": int(r['Pack Plan']),
            "mode": r['Mode'] if pd.notna(r['Mode']) else "",
            "dispatch_date": pd.Timestamp(r['Dispatch Date']).strftime('%d/%m/%Y') if pd.notna(r['Dispatch Date']) else "",
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
        })

    # ══════════════════════════════════════════════════════════════
    # FORECAST_RAW — row-level data so the Forecast tab can filter by Shipper/Modo/Pack Plan
    # on the client, then re-aggregate into weekly totals on the fly.
    # ══════════════════════════════════════════════════════════════
    FORECAST_RAW = []
    for _, r in df.iterrows():
        s6w = '' if pd.isna(r['6W']) else str(r['6W']).upper()
        is_airadd = 'AIRADD' in s6w
        is_exadd = ('EXADD' in s6w) and not is_airadd
        is_expo = ('EXPO' in s6w) and not is_airadd and not is_exadd
        is_add = ('ADD' in s6w) and not is_airadd and not is_exadd
        is_pre = 'PRE' in s6w
        status = r['Status']
        FORECAST_RAW.append({
            "pack_plan": int(r['Pack Plan']),
            "mode": r['Mode'] if pd.notna(r['Mode']) else "",
            "shipper": r['Shipper'] if pd.notna(r['Shipper']) else "",
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

    out = dict(SENASA=SENASA, WROWS=WROWS, AIR=AIR, AIRLINES=AIRLINES, FWKS=FWKS, WKL=WKL,
               KPI=KPI, EX=EX, DEST=DEST, LINEAS=LINEAS, PMI_PORTS=PMI_PORTS, FWKS_MODE=FWKS_MODE,
               EMBARQUES=EMBARQUES, TR_PROGRAM=TR_PROGRAM, PORTS_RAW=PORTS_RAW,
               FORECAST_RAW=FORECAST_RAW, META=META)

    return out
