#!/usr/bin/env python3
"""FASTAIR — ROTABLE (вкл. GSE/TOOL и KIT): поставщики ↔ клиенты.

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
        default=Path("/tmp/taz_history/ТАЗ 30.10.2026.xlsx"),
    )
    parser.add_argument("--out-dir", type=Path, default=Path("output"))
    parser.add_argument("--stem", default="rotables_suppliers_clients_3010")
    parser.add_argument(
        "--month-end",
        action="store_true",
        help="Считать дату среза концом месяца",
    )
    args = parser.parse_args()
    generate(
        args.taz,
        args.out_dir,
        CAT_ROTABLE,
        args.stem,
        month_end=args.month_end,
    )


if __name__ == "__main__":
    main()
