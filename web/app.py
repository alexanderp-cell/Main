#!/usr/bin/env python3
"""Browser UI for Gmail P/N quote search."""

from __future__ import annotations

import os
import sys
from functools import wraps
from pathlib import Path

from flask import Flask, jsonify, render_template, request, session

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from pn_quote_search import (  # noqa: E402
    TUZ_O_AA_HEADERS,
    hits_to_combined_tsv,
    load_env,
    parse_pn_text,
    search_quotes,
)

app = Flask(__name__)
app.secret_key = os.environ.get("QUOTE_DESK_SECRET", "quote-desk-dev-secret-change-me")


def _desk_password() -> str:
    env_path = ROOT / ".secrets" / "gmail.env"
    if env_path.exists():
        env = load_env(env_path)
        pw = (env.get("QUOTE_DESK_PASSWORD") or "").strip()
        if pw:
            return pw
    return (os.environ.get("QUOTE_DESK_PASSWORD") or "").strip()


def _authed() -> bool:
    required = _desk_password()
    if not required:
        # No password configured → allow (local-only use)
        return True
    return bool(session.get("authed"))


def require_auth(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if not _authed():
            return jsonify({"ok": False, "error": "Нужен пароль доступа."}), 401
        return view(*args, **kwargs)

    return wrapped


@app.get("/")
def index():
    return render_template(
        "index.html",
        headers=TUZ_O_AA_HEADERS,
        needs_password=bool(_desk_password()) and not _authed(),
    )


@app.post("/api/login")
def api_login():
    data = request.get_json(silent=True) or {}
    password = str(data.get("password") or "").strip()
    required = _desk_password()
    if not required:
        session["authed"] = True
        return jsonify({"ok": True})
    if password and password == required:
        session["authed"] = True
        return jsonify({"ok": True})
    return jsonify({"ok": False, "error": "Неверный пароль."}), 403


@app.get("/api/session")
def api_session():
    return jsonify(
        {
            "ok": True,
            "authed": _authed(),
            "needs_password": bool(_desk_password()) and not _authed(),
        }
    )


@app.post("/api/search")
@require_auth
def api_search():
    data = request.get_json(silent=True) or {}
    text = data.get("pns") or request.form.get("pns") or ""
    try:
        limit = int(data.get("limit") or request.form.get("limit") or 40)
    except (TypeError, ValueError):
        limit = 40
    limit = max(5, min(limit, 80))

    pns = parse_pn_text(str(text))
    if not pns:
        return jsonify({"ok": False, "error": "Введите хотя бы один партийный номер."}), 400

    try:
        results = search_quotes(pns, limit=limit)
    except Exception as exc:  # noqa: BLE001 — surface to UI
        return jsonify({"ok": False, "error": str(exc)}), 500

    return jsonify(
        {
            "ok": True,
            "headers": TUZ_O_AA_HEADERS,
            "results": results,
            "tsv": hits_to_combined_tsv(results),
            "total_quotes": sum(r["quote_count"] for r in results),
        }
    )


def main() -> None:
    host = "0.0.0.0"
    port = int(os.environ.get("PORT", "5055"))
    print(f"Quote search UI: http://127.0.0.1:{port}")
    if _desk_password():
        print("Password protection: ON")
    app.run(host=host, port=port, debug=False)


if __name__ == "__main__":
    main()
