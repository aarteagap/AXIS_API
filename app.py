import os
import tempfile
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

    if "file" not in request.files:
        return jsonify({"error": "no file uploaded, expected multipart field 'file'"}), 400

    f = request.files["file"]

    with tempfile.NamedTemporaryFile(suffix=".xlsm", delete=False) as tmp:
        f.save(tmp.name)
        tmp_path = tmp.name

    try:
        data = build_dashboard_data(tmp_path)
        return jsonify(data)
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        os.unlink(tmp_path)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
