"""dok-til-editor web — Streamlit app with magic-link auth for Avonova advisors.

Run:
    cd .claude/skills/dok-til-editor/web
    python -m pip install -r requirements.txt
    python -m streamlit run app.py

Environment variables (optional):
    SENDGRID_API_KEY  SendGrid API key for sending magic links via email.
                      Without it: token shown directly in UI (MVP/dev mode).
    SENDGRID_FROM     From-address for magic link emails (default: noreply@avonova.fi).
    APP_BASE_URL      Public URL used to build magic links (default: http://localhost:8501).
    ALLOWED_DOMAINS   Comma-separated approved email domains (default: avonova.no,avonova.fi).
"""
from __future__ import annotations

import io
import os
import re
import secrets
import sqlite3
import subprocess
import sys
import tempfile
import time
import zipfile
from pathlib import Path

import streamlit as st

WEB_DIR = Path(__file__).parent
_local_convert = WEB_DIR / "convert.py"
_skill_convert = WEB_DIR.parent / "convert.py"
CONVERT_PY = _local_convert if _local_convert.exists() else _skill_convert
DB_PATH = WEB_DIR / "tokens.db"

TOKEN_VALID_MIN = 15
SESSION_VALID_HOURS = 24
ALLOWED_DOMAINS = [d.strip() for d in os.environ.get("ALLOWED_DOMAINS", "avonova.no,avonova.fi").split(",") if d.strip()]
APP_BASE_URL = os.environ.get("APP_BASE_URL", "http://localhost:8501")
SENDGRID_API_KEY = os.environ.get("SENDGRID_API_KEY", "")
SENDGRID_FROM = os.environ.get("SENDGRID_FROM", "noreply@avonova.fi")

EMAIL_PATTERN = re.compile(r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$")


def db() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""CREATE TABLE IF NOT EXISTS tokens (
        token TEXT PRIMARY KEY, email TEXT, created_at REAL, used INTEGER DEFAULT 0
    )""")
    return conn


def create_token(email: str) -> str:
    token = secrets.token_urlsafe(24)
    with db() as conn:
        conn.execute("INSERT INTO tokens VALUES (?, ?, ?, 0)", (token, email, time.time()))
    return token


def consume_token(token: str) -> str | None:
    with db() as conn:
        row = conn.execute("SELECT email, created_at, used FROM tokens WHERE token = ?", (token,)).fetchone()
        if not row:
            return None
        email, created, used = row
        if used or time.time() - created > TOKEN_VALID_MIN * 60:
            return None
        conn.execute("UPDATE tokens SET used = 1 WHERE token = ?", (token,))
    return email


def send_magic_link(email: str, login_url: str) -> tuple[bool, str]:
    """Send magic link via SendGrid. Returns (success, message)."""
    if not SENDGRID_API_KEY:
        return False, "SENDGRID_API_KEY not configured — showing link directly below."
    try:
        import requests
        resp = requests.post(
            "https://api.sendgrid.com/v3/mail/send",
            headers={
                "Authorization": f"Bearer {SENDGRID_API_KEY}",
                "Content-Type": "application/json",
            },
            json={
                "personalizations": [{"to": [{"email": email}]}],
                "from": {"email": SENDGRID_FROM, "name": "dok-til-editor"},
                "subject": "Your sign-in link — dok-til-editor",
                "content": [{
                    "type": "text/html",
                    "value": (
                        f"<p>Click the link below to sign in to dok-til-editor. "
                        f"This link is valid for {TOKEN_VALID_MIN} minutes and can be used once.</p>"
                        f'<p><a href="{login_url}">Sign in</a></p>'
                        f"<p>If you did not request this link, you can ignore this email.</p>"
                    ),
                }],
            },
            timeout=10,
        )
        if resp.status_code in (200, 201, 202):
            return True, "Magic link sent to your email."
        return False, f"SendGrid error: {resp.status_code} {resp.text[:200]}"
    except Exception as e:
        return False, f"Send failed: {e}"


def is_logged_in() -> bool:
    if "user_email" not in st.session_state:
        return False
    if time.time() - st.session_state.get("login_at", 0) > SESSION_VALID_HOURS * 3600:
        st.session_state.pop("user_email", None)
        return False
    return True


def render_login():
    st.title("dok-til-editor")
    st.caption("Convert PDF/Word documents to CKEditor-friendly HTML")

    domain_list = " or ".join(f"@{d}" for d in ALLOWED_DOMAINS)
    st.markdown(f"Sign in with a {domain_list} email to continue.")

    with st.form("login"):
        email = st.text_input("Email address", placeholder="name@avonova.fi")
        submitted = st.form_submit_button("Send sign-in link")

    if submitted:
        email = (email or "").strip().lower()
        if not EMAIL_PATTERN.match(email):
            st.error("Invalid email address")
            return
        domain = email.split("@")[-1]
        if domain not in ALLOWED_DOMAINS:
            st.error(f"Only {domain_list} addresses are accepted")
            return

        token = create_token(email)
        login_url = f"{APP_BASE_URL}?token={token}"
        ok, msg = send_magic_link(email, login_url)
        if ok:
            st.success(msg)
        else:
            st.warning(msg)
            st.info("Dev mode — click the link below to sign in:")
            st.markdown(f"[Sign in now →]({login_url})")
            st.code(login_url, language=None)


def render_converter():
    user = st.session_state.user_email
    with st.sidebar:
        st.markdown(f"**Signed in as**\n{user}")
        if st.button("Sign out"):
            st.session_state.clear()
            st.rerun()

    st.title("Document converter")
    st.markdown("Upload a PDF or Word file. Get HTML ready to paste into the Mitt Avonova CKEditor.")

    uploaded = st.file_uploader("Choose a document", type=["pdf", "docx"])
    embed_images = st.checkbox(
        "Embed images in HTML (base64)",
        value=False,
        help="Makes the HTML large (potentially several MB). Recommended only for short instructions.",
    )

    if not uploaded:
        return

    if not st.button("Convert", type="primary"):
        return

    with st.spinner("Converting document..."):
        tmpdir = Path(tempfile.mkdtemp(prefix="dok-til-editor-"))
        input_path = tmpdir / uploaded.name
        input_path.write_bytes(uploaded.getbuffer())
        stem = input_path.stem
        output_path = tmpdir / f"{stem}.html"

        cmd = [sys.executable, str(CONVERT_PY), str(input_path), "-o", str(output_path)]
        if embed_images:
            cmd.append("--embed-images")

        try:
            result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", timeout=180)
        except subprocess.TimeoutExpired:
            st.error("Conversion took too long (>3 min).")
            return

        if result.returncode != 0 or not output_path.exists():
            st.error("Conversion failed")
            st.code(result.stderr or "(no error message)")
            return

        html = output_path.read_text(encoding="utf-8")
        warnings = [line.replace("[warn] ", "") for line in (result.stderr or "").splitlines() if "[warn]" in line]
        image_dir = tmpdir / f"{stem}_bilder"
        attached = sorted(image_dir.iterdir()) if image_dir.exists() else []

    st.success(f"Converted: {len(html):,} characters of HTML, {len(attached)} image attachment(s)")

    if warnings:
        with st.expander(f"{len(warnings)} notice(s)"):
            for w in warnings:
                st.text(w)

    tab_html, tab_preview, tab_download = st.tabs(["HTML source", "Preview", "Download"])

    with tab_html:
        st.caption("Select all and copy (Ctrl+A, Ctrl+C), or use the button in the top-right corner.")
        st.code(html, language="html")

    with tab_preview:
        st.caption("How the result will appear in CKEditor 4.")
        st.html(
            "<style>"
            ".ck-preview {background:#fff;padding:24px;border:1px solid #d8dce0;border-radius:4px;"
            "font-family:Arial,sans-serif;font-size:14px;line-height:1.6;color:#333;}"
            ".ck-preview h2{font-size:1.5em;margin:1em 0 0.4em;}"
            ".ck-preview h3{font-size:1.25em;margin:0.9em 0 0.3em;}"
            ".ck-preview h4{font-size:1.1em;margin:0.7em 0 0.2em;}"
            ".ck-preview ul,.ck-preview ol{padding-left:2em;margin:0.5em 0;}"
            ".ck-preview img{max-width:100%;height:auto;}"
            ".ck-preview table{border-collapse:collapse;margin:0.5em 0;}"
            ".ck-preview td,.ck-preview th{border:1px solid #d1d1d1;padding:4px 8px;}"
            "</style>"
            f'<div class="ck-preview">{html}</div>'
        )

    with tab_download:
        if attached:
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                zf.writestr(output_path.name, html)
                for img in attached:
                    zf.write(img, f"{image_dir.name}/{img.name}")
            st.download_button(
                "Download ZIP (HTML + images)",
                data=buf.getvalue(),
                file_name=f"{stem}.zip",
                mime="application/zip",
            )
        st.download_button(
            "Download HTML only",
            data=html.encode("utf-8"),
            file_name=output_path.name,
            mime="text/html",
        )


def main() -> None:
    st.set_page_config(page_title="dok-til-editor", page_icon="📄", layout="wide")

    params = st.query_params
    incoming_token = params.get("token")
    if incoming_token and not is_logged_in():
        email = consume_token(incoming_token)
        st.query_params.clear()
        if email:
            st.session_state.user_email = email
            st.session_state.login_at = time.time()
            st.rerun()
        else:
            st.error("Token is invalid or expired. Please request a new link.")

    if is_logged_in():
        render_converter()
    else:
        render_login()


if __name__ == "__main__":
    main()
