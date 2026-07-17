import os
import tempfile
import base64
import json
from datetime import datetime, timezone
import requests
from flask import Flask, request, jsonify

app = Flask(__name__)

# Enable CORS so the dashboard (hosted on GitHub Pages, a different origin) can call this API.
@app.after_request
def add_cors_headers(resp):
    resp.headers['Access-Control-Allow-Origin'] = '*'
    resp.headers['Access-Control-Allow-Headers'] = 'Content-Type, X-API-Key'
    resp.headers['Access-Control-Allow-Methods'] = 'GET, POST, OPTIONS'
    return resp

@app.route("/publish", methods=["OPTIONS"])
@app.route("/convert", methods=["OPTIONS"])
def cors_preflight():
    return ("", 204)

from etl_logic import build_dashboard_data

# Shared secret so random people on the internet can't hit this endpoint.
# Set this as an environment variable in Render (see deployment instructions).
API_KEY = os.environ.get("API_KEY", "change-me")

# GitHub settings for the /publish endpoint. Set these as environment
# variables in Render — never hardcode a real token here.
GITHUB_TOKEN = os.environ.get("GITHUB_TOKEN", "")
GITHUB_REPO = os.environ.get("GITHUB_REPO", "aarteagap/ATHENA_dashboard")
GITHUB_FILE_PATH = os.environ.get("GITHUB_FILE_PATH", "data.json")
GITHUB_BRANCH = os.environ.get("GITHUB_BRANCH", "main")


def github_headers():
    return {
        "Authorization": f"token {GITHUB_TOKEN}",
        "Accept": "application/vnd.github+json",
        "User-Agent": "athena-etl-api",
    }


def push_to_github(data_dict, updater_name=""):
    """Fetch the current SHA of data.json, then PUT the new content. Returns (ok, info)."""
    if not GITHUB_TOKEN:
        return False, "GITHUB_TOKEN is not configured on the server."

    api_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_FILE_PATH}"

    # 1. Get current SHA (required by GitHub to update an existing file)
    r = requests.get(api_url, headers=github_headers(), params={"ref": GITHUB_BRANCH})
    if r.status_code != 200:
        return False, f"Could not read current file (status {r.status_code}): {r.text[:300]}"
    sha = r.json().get("sha")

    # 2. PUT the new content
    who = updater_name.strip() if updater_name and updater_name.strip() else "Usuario sin identificar"
    timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    content_b64 = base64.b64encode(json.dumps(data_dict, ensure_ascii=False).encode("utf-8")).decode("ascii")
    payload = {
        "message": f"Actualización de datos por {who} — {timestamp}",
        "content": content_b64,
        "sha": sha,
        "branch": GITHUB_BRANCH,
        "committer": {"name": who, "email": "athena-dashboard@no-reply.local"},
        "author": {"name": who, "email": "athena-dashboard@no-reply.local"},
    }
    r2 = requests.put(api_url, headers=github_headers(), json=payload)
    if r2.status_code not in (200, 201):
        return False, f"GitHub rejected the update (status {r2.status_code}): {r2.text[:300]}"
    return True, r2.json().get("commit", {}).get("sha", "")


@app.route("/", methods=["GET"])
def health():
    return jsonify({"status": "ok", "service": "horizon-etl-api"})


def _load_excel_from_request():
    """Shared logic: read the uploaded Excel from either multipart or base64 JSON body."""
    if "file" in request.files:
        f = request.files["file"]
        with tempfile.NamedTemporaryFile(suffix=".xlsm", delete=False) as tmp:
            f.save(tmp.name)
            return tmp.name
    payload = request.get_json(silent=True) or {}
    b64 = payload.get("file_base64")
    if not b64:
        return None
    raw = base64.b64decode(b64)
    with tempfile.NamedTemporaryFile(suffix=".xlsm", delete=False) as tmp:
        tmp.write(raw)
        return tmp.name


@app.route("/convert", methods=["POST"])
def convert():
    """Convert an uploaded Excel to the dashboard JSON structure. Does NOT touch GitHub."""
    if request.headers.get("X-API-Key") != API_KEY:
        return jsonify({"error": "unauthorized"}), 401

    tmp_path = _load_excel_from_request()
    if not tmp_path:
        return jsonify({"error": "no file uploaded. Send multipart field 'file' or JSON {file_base64}"}), 400

    try:
        data = build_dashboard_data(tmp_path)
        return jsonify(data)
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)


@app.route("/publish", methods=["POST"])
def publish():
    """Convert an uploaded Excel AND push the result to GitHub as data.json."""
    if request.headers.get("X-API-Key") != API_KEY:
        return jsonify({"error": "unauthorized"}), 401

    updater_name = request.form.get("updater_name") or (request.get_json(silent=True) or {}).get("updater_name", "")

    tmp_path = _load_excel_from_request()
    if not tmp_path:
        return jsonify({"error": "no file uploaded. Send multipart field 'file' or JSON {file_base64}"}), 400

    try:
        data = build_dashboard_data(tmp_path)
    except Exception as e:
        return jsonify({"error": f"conversion failed: {e}"}), 500
    finally:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)

    ok, info = push_to_github(data, updater_name)
    if not ok:
        return jsonify({"error": info}), 502

    return jsonify({"status": "ok", "commit_sha": info, "updated_by": updater_name, "rows": {k: len(v) if isinstance(v, list) else None for k, v in data.items()}})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
