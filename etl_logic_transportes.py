"""
ETL — Control de Transportes (TABLEAU_CONSOLIDADO_DE_TRANSPORTE_2627.xlsx)

Lee la hoja "CONSOLIDADO 2526", filtra filas de plantilla vacías (pp nulo),
y produce:
  - TRANSPORT_RAW: lista de filas a nivel de registro, con los MISMOS nombres
    de campo que ya consume el frontend (control_tower_transporte__4_.html /
    dashboard.html): pp, sede, transportista, naviera, almacenRetiro,
    almacenIngreso, puertoRetiro, puertoIngreso, gateOutCompliance,
    packingCompliance, gateInCompliance, detentionGateOut, detentionPacking,
    detentionGateIn, dispatchDelay, preLoadingWait, emptyOutboundH,
    routeStartDelayH, tiempoCargaMin, checkInDelayH, entryDelayH,
    fullInboundH, detGateOutH, detPackingH, positioningDeltaMin.
  - WEEKLY, SUMMARY, BY_PORT_PICKUP, BY_PORT_RECEIVE, BY_LINE, BY_CARRIER:
    agregados de conveniencia (no los usa el dashboard actual, que agrega
    todo client-side desde TRANSPORT_RAW, pero se publican igual por si se
    necesitan en otra vista o para análisis directo en Supabase).

⚠️ Nota sobre 'detPackingH': el Excel no trae una columna explícita de horas
para "Detention at Packing" (solo la categoría NO DETENTION/HIGH/CRITICAL).
Se aproxima aquí como las horas entre 'Ingreso Packing' y 'Salida Packing'.
Si tu definición real es otra, ajusta la función `_det_packing_hours()`.

Uso:
    python3 etl_logic_transportes.py /ruta/al/TABLEAU_CONSOLIDADO_DE_TRANSPORTE_2627.xlsx
    → imprime un JSON {dataset_name: payload, ...} por stdout, listo para
      hacer upsert en la tabla transport_data (ver schema_transportes.sql
      y el endpoint /publish-transportes en app.py).
"""
import sys
import json
from datetime import datetime, time, timedelta
from collections import defaultdict
import openpyxl

SHEET_NAME = 'CONSOLIDADO 2526'

def _to_hours(v):
    """Convierte time/timedelta de Excel a horas (float). None si no aplica."""
    if v is None:
        return None
    if isinstance(v, timedelta):
        return round(v.total_seconds() / 3600, 2)
    if isinstance(v, time):
        return round((v.hour * 3600 + v.minute * 60 + v.second) / 3600, 2)
    return None

def _to_minutes(v):
    if v is None:
        return None
    if isinstance(v, timedelta):
        return round(v.total_seconds() / 60, 1)
    if isinstance(v, time):
        return round((v.hour * 3600 + v.minute * 60 + v.second) / 60, 1)
    return None

def _minutes_between(a, b):
    """Minutos entre dos datetimes (b - a). None si falta alguno."""
    if not isinstance(a, datetime) or not isinstance(b, datetime):
        return None
    return round((b - a).total_seconds() / 60, 1)

def _clean_str(v):
    if v is None:
        return None
    if isinstance(v, str):
        v = v.strip()
        return v if v and v.upper() != '#VALUE!' else None
    return v

def load_rows(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb[SHEET_NAME]
    rows = list(ws.iter_rows(min_row=1, values_only=True))
    headers = rows[0]
    idx = {h: i for i, h in enumerate(headers) if h}

    def val(r, name):
        i = idx.get(name)
        return r[i] if i is not None else None

    out = []
    for r in rows[1:]:
        pp = val(r, 'pp')
        if pp is None:
            continue  # filas de plantilla vacías, per handoff

        sede_raw = _clean_str(val(r, 'Sede'))
        sede = sede_raw.title() if sede_raw else sede_raw  # normaliza "CHAO" -> "Chao"

        ingreso_packing = val(r, 'Ingreso Packing')
        salida_packing = val(r, 'Salida Packing')
        cita_pos = val(r, 'Cita - Hora de Posicionamiento')  # string "DD/MM/YYYY HH:MM"
        llegada_packing = val(r, 'Llegada a packing')
        positioning_delta = None
        if isinstance(cita_pos, str) and isinstance(llegada_packing, datetime):
            try:
                cita_dt = datetime.strptime(cita_pos.strip(), '%d/%m/%Y %H:%M')
                positioning_delta = _minutes_between(cita_dt, llegada_packing)
            except ValueError:
                positioning_delta = None

        row = {
            'id': val(r, 'ID'),
            'instruction': _clean_str(val(r, 'Instrucción')),
            'pp': int(pp),
            'sede': sede,
            'transportista': _clean_str(val(r, 'Transportista')),
            'naviera': _clean_str(val(r, 'LINEA NAVIERA')),
            'almacenRetiro': _clean_str(val(r, 'Almacén retiro')),
            'almacenIngreso': _clean_str(val(r, 'Almacén de Ingreso')),
            'puertoRetiro': _clean_str(val(r, 'Puerto retiro')),
            'puertoIngreso': _clean_str(val(r, 'Puerto ingreso')),
            'fcl': val(r, 'FCL'),
            'estatus': _clean_str(val(r, 'ESTATUS')),

            'gateOutCompliance': _clean_str(val(r, 'Gate-Out Appointment Compliance')),
            'packingCompliance': _clean_str(val(r, 'Packing Appointment Compliance')),
            'gateInCompliance': _clean_str(val(r, 'Gate-In Appointment Compliance')),

            'detentionGateOut': _clean_str(val(r, 'Detention at Gate-Out Storage')),
            'detentionPacking': _clean_str(val(r, 'Detention at Packing')),
            'detentionGateIn': _clean_str(val(r, 'Detention at Gate-In Storage')),

            'dispatchDelay': _clean_str(val(r, 'Dispatch Delay')),
            'preLoadingWait': _clean_str(val(r, 'Pre-Loading Wait Time')),

            'emptyOutboundH': _to_hours(val(r, 'Empty Outbound Time')),
            'routeStartDelayH': _to_hours(val(r, 'Route Start Delay')),
            'checkInDelayH': _to_hours(val(r, 'Check-In Delay')),
            'entryDelayH': _to_hours(val(r, 'Entry Delay')),
            'fullInboundH': _to_hours(val(r, 'Full Inbound Time')),
            'detGateOutH': _to_hours(val(r, 'Gate out Detention')),
            'tiempoCargaMin': _to_minutes(val(r, 'Tiempo de carga')),
            # Ver nota de cabecera: aproximación, no viene explícita en el Excel.
            'detPackingH': _minutes_between(ingreso_packing, salida_packing) and round(_minutes_between(ingreso_packing, salida_packing) / 60, 2),
            'positioningDeltaMin': positioning_delta,
        }
        out.append(row)
    return out


# ── Reglas de calidad de datos por campo ────────────────────
# (campo, mínimo válido, máximo válido, motivo legible en inglés — igual
# estilo/idioma que ya usa el resto de labels de Transporte en el dashboard)
DATA_QUALITY_RULES = [
    ('emptyOutboundH', 0, 96,  'Arrival at plant occurs before route start, or transit exceeds 4 days'),
    ('routeStartDelayH', 0, 48, 'Route start occurs before warehouse departure, or delay exceeds 2 days'),
    ('checkInDelayH', 0, 24,   'Check-in delay is negative or exceeds 24 hours'),
    ('entryDelayH', 0, 24,     'Entry delay is negative or exceeds 24 hours'),
    ('fullInboundH', 0, 96,    'Arrival at port occurs before departure from plant, or transit exceeds 4 days'),
    ('detGateOutH', 0, 48,     'Detention time at gate-out is negative or exceeds 2 days'),
    ('tiempoCargaMin', 0, 720, 'Loading time is negative or exceeds 12 hours'),
    ('detPackingH', 0, 48,     'Detention time at packing is negative or exceeds 2 days'),
]

def data_quality_report(rows):
    """Para cada fila, revisa los campos de duración contra un rango válido
    de negocio. Devuelve la lista de excepciones (una por campo problemático,
    una fila puede aportar más de una) para mostrar en el banner de calidad
    de datos, y el set de índices de fila afectados."""
    issues = []
    bad_row_idx = set()
    for i, r in enumerate(rows):
        for field, lo, hi, reason in DATA_QUALITY_RULES:
            v = r.get(field)
            if v is None:
                continue
            if v < lo or v > hi:
                issues.append({
                    'instruction': r.get('instruction') or f"ID-{r.get('id')}" if r.get('id') is not None else '—',
                    'pp': r.get('pp'),
                    'sede': r.get('sede'),
                    'field': field,
                    'value': v,
                    'reason': reason,
                })
                bad_row_idx.add(i)
    return issues, bad_row_idx


def avg_excluding_quality_issues(rows, field):
    """Promedio de un campo, excluyendo los valores marcados por
    DATA_QUALITY_RULES para ese campo específico (no toda la fila)."""
    rule = next((ru for ru in DATA_QUALITY_RULES if ru[0] == field), None)
    if not rule:
        vals = [r[field] for r in rows if r.get(field) is not None]
        return sum(vals) / len(vals) if vals else None
    _, lo, hi, _ = rule
    vals = [r[field] for r in rows if r.get(field) is not None and lo <= r[field] <= hi]
    return sum(vals) / len(vals) if vals else None


def _count_by(rows, field):
    o = defaultdict(int)
    for r in rows:
        v = r.get(field)
        if v is not None:
            o[v] += 1
    return dict(o)

def build_datasets(rows):
    """Agregados de conveniencia (no requeridos por el dashboard actual,
    que agrega todo client-side desde TRANSPORT_RAW, pero se publican para
    consultas directas / futuras vistas)."""
    weekly = defaultdict(int)
    for r in rows:
        weekly[f"S{r['pp']}"] += 1

    by_port_pickup = defaultdict(lambda: {'total': 0, 'on_time': 0})
    for r in rows:
        if not r['puertoRetiro']:
            continue
        by_port_pickup[r['puertoRetiro']]['total'] += 1
        if r['gateOutCompliance'] == 'ON TIME':
            by_port_pickup[r['puertoRetiro']]['on_time'] += 1

    by_port_receive = defaultdict(lambda: {'total': 0, 'on_time': 0})
    for r in rows:
        if not r['puertoIngreso']:
            continue
        by_port_receive[r['puertoIngreso']]['total'] += 1
        if r['gateInCompliance'] == 'ON TIME':
            by_port_receive[r['puertoIngreso']]['on_time'] += 1

    by_line = defaultdict(lambda: {'total': 0, 'on_time': 0})
    for r in rows:
        if not r['naviera']:
            continue
        by_line[r['naviera']]['total'] += 1
        if r['packingCompliance'] == 'ON TIME':
            by_line[r['naviera']]['on_time'] += 1

    by_carrier = defaultdict(lambda: {'total': 0, 'on_time': 0})
    for r in rows:
        if not r['transportista']:
            continue
        by_carrier[r['transportista']]['total'] += 1
        if r['gateInCompliance'] == 'ON TIME':
            by_carrier[r['transportista']]['on_time'] += 1

    summary = {
        'total_rows': len(rows),
        'gate_out_compliance': _count_by(rows, 'gateOutCompliance'),
        'packing_compliance': _count_by(rows, 'packingCompliance'),
        'gate_in_compliance': _count_by(rows, 'gateInCompliance'),
        'weeks': sorted({r['pp'] for r in rows}),
    }

    dq_issues, dq_bad_rows = data_quality_report(rows)

    return {
        'TRANSPORT_RAW': rows,
        'WEEKLY': dict(sorted(weekly.items(), key=lambda kv: int(kv[0][1:]))),
        'SUMMARY': summary,
        'BY_PORT_PICKUP': dict(by_port_pickup),
        'BY_PORT_RECEIVE': dict(by_port_receive),
        'BY_LINE': dict(by_line),
        'BY_CARRIER': dict(by_carrier),
        'DATA_QUALITY': {'issues': dq_issues, 'affected_rows': len(dq_bad_rows), 'total_rows': len(rows)},
    }


def process(xlsx_path):
    rows = load_rows(xlsx_path)
    return build_datasets(rows)


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Uso: python3 etl_logic_transportes.py <ruta al xlsx>', file=sys.stderr)
        sys.exit(1)
    datasets = process(sys.argv[1])
    print(json.dumps(datasets, ensure_ascii=False, default=str, indent=2))
