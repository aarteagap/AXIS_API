import os
import tempfile
import base64
from flask import Flask, request, jsonify
from etl_logic import build_dashboard_data

app = Flask(__name__)

# Shared secret so random people on the internet can't hit this endpoint.
# Set this as an environment variable in Render (see deployment instructions).
API_KEY = os.environ.get("API_KEY", "change-me")


@app.route("/", methods=["GET"])
def health():
    return jsonify({"status": "ok", "service": "horizon-etl-api"})


@app.route("/convert", methods=["POST"])
def convert():
    # Simple auth: header "X-API-Key: <API_KEY>"
    if request.headers.get("X-API-Key") != API_KEY:
        return jsonify({"error": "unauthorized"}), 401

    tmp_path = None
    try:
        if "file" in request.files:
            # multipart/form-data upload
            f = request.files["file"]
            with tempfile.NamedTemporaryFile(suffix=".xlsm", delete=False) as tmp:
                f.save(tmp.name)
                tmp_path = tmp.name
        else:
            # JSON body: { "file_base64": "..." }  (used by Power Automate)
            payload = request.get_json(silent=True) or {}
            b64 = payload.get("file_base64")
            if not b64:
                return jsonify({"error": "no file uploaded. Send multipart field 'file' or JSON {file_base64}"}), 400
            raw = base64.b64decode(b64)
            with tempfile.NamedTemporaryFile(suffix=".xlsm", delete=False) as tmp:
                tmp.write(raw)
                tmp_path = tmp.name

        data = build_dashboard_data(tmp_path)
        return jsonify(data)
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        if tmp_path and os.path.exists(tmp_path):
            os.unlink(tmp_path)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
