# ============================================================
# app_publish_transportes_snippet.py
#
# Agrega ESTE bloque a tu app.py existente en el repo AXIS_API
# (junto al endpoint /publish que ya está en producción). Reutiliza
# las mismas variables de entorno que ya tienes configuradas en Render
# (API_KEY, GITHUB_TOKEN, GITHUB_REPO, SUPABASE_URL, SUPABASE_SERVICE_KEY)
# — no crea infraestructura nueva, solo un endpoint más en el mismo
# servicio Flask.
#
# Requiere que ya hayas ejecutado schema_transportes.sql y que
# etl_logic_transportes.py esté en el mismo repo (junto a app.py).
# ============================================================

import os
import base64
import requests
from flask import request, jsonify
from etl_logic_transportes import process as process_transportes

# Reutiliza las mismas env vars que ya usa /publish — no dupliques secretos.
API_KEY = os.environ["API_KEY"]
GITHUB_TOKEN = os.environ["GITHUB_TOKEN"]
GITHUB_REPO = os.environ["GITHUB_REPO"]  # ej. "aarteagap/AXIS_API"
SUPABASE_URL = os.environ["SUPABASE_URL"]
SUPABASE_SERVICE_KEY = os.environ["SUPABASE_SERVICE_KEY"]


def _require_api_key():
    key = request.headers.get("X-API-Key")
    if key != API_KEY:
        return jsonify({"error": "API key inválida"}), 401
    return None


def _upsert_transport_data(datasets: dict, updated_by: str = "publish-transportes"):
    """Hace upsert de cada dataset en la tabla transport_data usando el
    service_role key (evita RLS, igual que hace /publish con dashboard_data)."""
    url = f"{SUPABASE_URL}/rest/v1/transport_data"
    headers = {
        "apikey": SUPABASE_SERVICE_KEY,
        "Authorization": f"Bearer {SUPABASE_SERVICE_KEY}",
        "Content-Type": "application/json",
        "Prefer": "resolution=merge-duplicates",
    }
    rows = [
        {
            "dataset_name": name,
            "payload": payload,
            "supplier_visible": True,
            "updated_by": updated_by,
        }
        for name, payload in datasets.items()
    ]
    resp = requests.post(url, headers=headers, json=rows, timeout=30)
    resp.raise_for_status()
    return resp


def _fetch_line_by_instruction():
    """Trae el mapa Instruction -> Line (naviera) publicado por /publish
    (el AXIS principal, fuente autoritativa: así se creó el requerimiento
    en Horizon) desde la tabla dashboard_data. Si Supabase no responde o
    el AXIS principal no se ha publicado todavía, devuelve {} — el ETL de
    Transportes cae de vuelta a lo que haya en la columna 'Línea Naviera'
    del propio Excel de transportes (llenada a mano por el proveedor)."""
    if not SUPABASE_URL or not SUPABASE_SERVICE_KEY:
        return {}
    url = f"{SUPABASE_URL}/rest/v1/dashboard_data"
    headers = {"apikey": SUPABASE_SERVICE_KEY, "Authorization": f"Bearer {SUPABASE_SERVICE_KEY}"}
    try:
        r = requests.get(url, headers=headers, params={"dataset_name": "eq.LINE_BY_INSTRUCTION", "select": "payload"}, timeout=15)
        r.raise_for_status()
        rows = r.json()
        return rows[0]["payload"] if rows else {}
    except Exception:
        return {}


def _commit_excel_to_github(file_bytes: bytes, filename: str = "TABLEAU_CONSOLIDADO_DE_TRANSPORTE_2627.xlsx"):
    """Sube el Excel al repo del dashboard, igual que hace /publish con el
    Excel de AXIS — deja rastro de qué archivo generó qué datos."""
    api_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/data/{filename}"
    headers = {"Authorization": f"token {GITHUB_TOKEN}"}

    # Necesita el sha actual si el archivo ya existe (para actualizar en vez de crear)
    existing = requests.get(api_url, headers=headers, timeout=15)
    sha = existing.json().get("sha") if existing.status_code == 200 else None

    payload = {
        "message": f"Actualiza {filename} vía /publish-transportes",
        "content": base64.b64encode(file_bytes).decode("utf-8"),
    }
    if sha:
        payload["sha"] = sha

    resp = requests.put(api_url, headers=headers, json=payload, timeout=30)
    resp.raise_for_status()
    return resp.json().get("commit", {}).get("sha")


# ── Registrar en tu app Flask existente ─────────────────────
# Si tu app.py usa el patrón `@app.route(...)`, agrega esto tal cual.
# Si usa Blueprints, adáptalo al blueprint que ya tengas para /publish.

def register_publish_transportes(app):
    @app.route("/publish-transportes", methods=["POST"])
    def publish_transportes():
        auth_error = _require_api_key()
        if auth_error:
            return auth_error

        if "file" not in request.files:
            return jsonify({"error": "Falta el archivo (campo 'file')"}), 400

        f = request.files["file"]
        file_bytes = f.read()

        # Guarda temporalmente para que openpyxl pueda leerlo desde disco
        tmp_path = "/tmp/_transportes_upload.xlsx"
        with open(tmp_path, "wb") as out:
            out.write(file_bytes)

        line_by_instruction = _fetch_line_by_instruction()
        try:
            datasets = process_transportes(tmp_path, line_by_instruction=line_by_instruction)
        except Exception as e:
            return jsonify({"error": f"Error procesando el Excel: {e}"}), 422

        try:
            _upsert_transport_data(datasets)
        except Exception as e:
            return jsonify({"error": f"Error publicando a Supabase: {e}"}), 502

        commit_sha = None
        try:
            commit_sha = _commit_excel_to_github(file_bytes)
        except Exception as e:
            # No es fatal: los datos ya están en Supabase aunque el commit falle.
            app.logger.warning(f"No se pudo commitear el Excel a GitHub: {e}")

        return jsonify({
            "ok": True,
            "rows_processed": len(datasets.get("TRANSPORT_RAW", [])),
            "datasets": list(datasets.keys()),
            "commit_sha": commit_sha,
        })


# En tu app.py, después de crear `app = Flask(__name__)`, agrega:
#     from app_publish_transportes_snippet import register_publish_transportes
#     register_publish_transportes(app)
