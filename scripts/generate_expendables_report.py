#!/usr/bin/env python3
"""FASTAIR — EXPENDABLE (вкл. consumable): поставщики ↔ клиенты.

Обёртка над generate_category_suppliers_clients_report.py (--bucket EXPENDABLE).
"""

from __future__ import annotations

import argparse
from pathlib import Path

from generate_category_suppliers_clients_report import CAT_EXPENDABLE, generate


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--taz",
        type=Path,
        default=Path("/tmp/taz_history/ТАЗ 09.10.2026.xlsx"),
    )
    parser.add_argument("--out-dir", type=Path, default=Path("output"))
    parser.add_argument("--stem", default="expendables_suppliers_clients")
    args = parser.parse_args()
    generate(args.taz, args.out_dir, CAT_EXPENDABLE, args.stem)


if __name__ == "__main__":
    main()
