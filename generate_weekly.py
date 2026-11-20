#!/usr/bin/env python3
"""Еженедельная рутина: Utair + Белавиа статусы + отчёт «без движения / ожидают оплаты».

Пример:
  python3 generate_weekly.py \\
    --taz "/path/ТАЗ 06.11.2026.xlsx" \\
    --date 06.11.2026 \\
    --previous-date 30.10.2026 \\
    --output-dir output
"""

from __future__ import annotations

import argparse
import re
from datetime import date, timedelta
from pathlib import Path

from generate_report import generate_report, parse_report_date
from generate_stale_orders_report import generate_stale_report


def _infer_date(path: Path, explicit: date | None) -> date:
    if explicit is not None:
        return explicit
    m = re.search(r"(\d{2})\.(\d{2})\.(\d{4})", path.name)
    if not m:
        raise SystemExit("Cannot infer date from TAZ filename; pass --date")
    d, mo, y = map(int, m.groups())
    return date(y, mo, d)


def _find_previous_report(output_dir: Path, client: str, before: date) -> Path | None:
    best: tuple[date, Path] | None = None
    for path in output_dir.glob(f"{client} статус *.xlsx"):
        m = re.search(r"(\d{2})\.(\d{2})\.(\d{4})", path.name)
        if not m:
            continue
        d, mo, y = map(int, m.groups())
        try:
            snap = date(y, mo, d)
        except ValueError:
            continue
        if snap >= before:
            continue
        if best is None or snap > best[0]:
            best = (snap, path)
    return best[1] if best else None


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--taz", type=Path, required=True, help="Current TAZ .xlsx")
    parser.add_argument("--date", type=parse_report_date, default=None)
    parser.add_argument(
        "--previous-date",
        type=parse_report_date,
        default=None,
        help="Start of weekly summary (default: report date − 7 days)",
    )
    parser.add_argument(
        "--previous-taz",
        type=Path,
        default=None,
        help="Previous TAZ for weekly summary (optional)",
    )
    parser.add_argument("--output-dir", type=Path, default=Path("output"))
    parser.add_argument("--stale-days", type=int, default=14)
    parser.add_argument(
        "--skip-stale",
        action="store_true",
        help="Only build Utair/Belavia status files",
    )
    args = parser.parse_args()

    report_date = _infer_date(args.taz, args.date)
    previous_date = args.previous_date or (report_date - timedelta(days=7))
    args.output_dir.mkdir(parents=True, exist_ok=True)

    jobs = [
        ("Utair", False),
        ("Белавиа", True),  # include Invoicer ФЭ
    ]

    written: list[Path] = []
    for client, include_fe in jobs:
        out = args.output_dir / f"{client} статус {report_date.strftime('%d.%m.%Y')}.xlsx"
        prev_report = _find_previous_report(args.output_dir, client, report_date)
        print(f"\n=== {client} → {out.name} ===")
        if prev_report:
            print(f"previous report: {prev_report.name}")
        generate_report(
            args.taz,
            out,
            client=client,
            report_date=report_date,
            previous_input=args.previous_taz,
            previous_report=prev_report,
            previous_date=previous_date,
            week_days=7,
            excluded_invoicers=set() if include_fe else None,
        )
        written.append(out)
        print(f"saved {out}")

    if not args.skip_stale:
        print("\n=== без движения / ожидают оплаты ===")
        stale_path = generate_stale_report(
            args.taz,
            args.output_dir,
            report_date,
            stale_days=args.stale_days,
        )
        written.append(stale_path)

    print("\nWeekly bundle done:")
    for p in written:
        print(f"  - {p}")


if __name__ == "__main__":
    main()
