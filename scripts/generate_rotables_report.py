#!/usr/bin/env python3
"""FASTAIR — ROTABLE (вкл. GSE/TOOL): поставщики ↔ клиенты.

Обёртка над generate_category_suppliers_clients_report.py (--bucket ROTABLE).
"""

from __future__ import annotations

import argparse
from pathlib import Path

from generate_category_suppliers_clients_report import CAT_ROTABLE, generate


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--taz",
        type=Path,
        default=Path("/tmp/taz_history/ТАЗ 23.10.2026.xlsx"),
    )
    parser.add_argument("--out-dir", type=Path, default=Path("output"))
    parser.add_argument("--stem", default="rotables_suppliers_clients_2310")
    args = parser.parse_args()
    generate(args.taz, args.out_dir, CAT_ROTABLE, args.stem)


if __name__ == "__main__":
    main()
