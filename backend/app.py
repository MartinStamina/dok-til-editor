"""dok-til-editor backend — Flask service that wraps convert.py.

Accepts POST /convert with multipart file + JWT bearer token.
Returns JSON: { html, warnings, images: [{name, dataUri, size}] }.
"""
from __future__ import annotations

import base64
import mimetypes
import os
import subprocess
import sys
import tempfile
from pathlib import Path

import jwt
from flask import Flask, jsonify, request

app = Flask(__name__)

JWT_SECRET = os.environ.get("JWT_SECRET", "")
ALLOWED_ORIGINS = {
    "https://protonova.avonova-apps.com",
    "https://protonova.netlify.app",
    "http://localhost:8888",
}
HERE = Path(__file__).parent
CONVERT_PY = HERE / "convert.py"


@app.after_request
def cors(resp):
    origin = request.headers.get("Origin", "")
    if origin in ALLOWED_ORIGINS:
        resp.headers["Access-Control-Allow-Origin"] = origin
    resp.headers["Access-Control-Allow-Headers"] = "Content-Type, Authorization"
    resp.headers["Access-Control-Allow-Methods"] = "POST, OPTIONS"
    return resp


def verify_token() -> dict | None:
    auth = request.headers.get("Authorization", "")
    if not auth.startswith("Bearer "):
        return None
    token = auth[7:]
    try:
        return jwt.decode(token, JWT_SECRET, algorithms=["HS256"])
    except Exception:
        return None


@app.route("/health")
def health():
    return jsonify({"ok": True})


@app.route("/convert", methods=["POST", "OPTIONS"])
def convert():
    if request.method == "OPTIONS":
        return ("", 204)

    payload = verify_token()
    if not payload:
        return jsonify({"error": "Invalid or missing auth token"}), 401

    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    uploaded = request.files["file"]
    embed = request.form.get("embed_images", "0") == "1"

    ext = Path(uploaded.filename or "").suffix.lower()
    if ext not in (".pdf", ".docx"):
        return jsonify({"error": "Only PDF and DOCX files are supported"}), 400

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        input_path = tmp / (uploaded.filename or f"input{ext}")
        uploaded.save(input_path)
        output_path = tmp / f"{input_path.stem}.html"

        cmd = [sys.executable, str(CONVERT_PY), str(input_path), "-o", str(output_path)]
        if embed:
            cmd.append("--embed-images")

        try:
            result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", timeout=180)
        except subprocess.TimeoutExpired:
            return jsonify({"error": "Conversion exceeded 3 minutes"}), 504

        if result.returncode != 0 or not output_path.exists():
            return jsonify({"error": "Conversion failed", "details": (result.stderr or "")[-500:]}), 500

        html = output_path.read_text(encoding="utf-8")
        warnings = [line.replace("[warn] ", "") for line in (result.stderr or "").splitlines() if "[warn]" in line]

        image_dir = tmp / f"{input_path.stem}_bilder"
        images = []
        if image_dir.exists():
            for img in sorted(image_dir.iterdir()):
                if not img.is_file():
                    continue
                mime, _ = mimetypes.guess_type(str(img))
                data = base64.b64encode(img.read_bytes()).decode("ascii")
                images.append({
                    "name": img.name,
                    "dataUri": f"data:{mime or 'application/octet-stream'};base64,{data}",
                    "size": img.stat().st_size,
                })

        return jsonify({"html": html, "warnings": warnings, "images": images})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
