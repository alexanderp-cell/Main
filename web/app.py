#!/usr/bin/env python3
"""Browser UI for Gmail P/N quote search."""

from __future__ import annotations

import sys
from pathlib import Path

from flask import Flask, jsonify, render_template, request

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from pn_quote_search import (  # noqa: E402
    TUZ_O_AA_HEADERS,
    hits_to_combined_tsv,
    parse_pn_text,
    search_quotes,
)

app = Flask(__name__)


@app.get("/")
def index():
    return render_template("index.html", headers=TUZ_O_AA_HEADERS)


@app.post("/api/search")
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
    port = 5055
    print(f"Quote search UI: http://127.0.0.1:{port}")
    app.run(host=host, port=port, debug=False)


if __name__ == "__main__":
    main()
