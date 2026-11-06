#!/usr/bin/env python3
"""Highlight Utair / Belavia orders whose Status has not changed for >14 days.

Compares the current TAZ «в работе» (+ unpaid shipped) against previous weekly
status Excel snapshots in output/.
"""

from __future__ import annotations

import argparse
import re
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from generate_report import (
    COL_BALANCE,
    COL_COMMENT,
    COL_CUSTOMER,
    COL_DESC,
    COL_INVOICE,
    COL_ORDER_DATE,
    COL_PN,
    COL_QTY,
    COL_SALE,
    COL_STATUS,
    COL_UNIT_PRICE,
    EXCLUDED_INVOICERS,
    add_row_key,
    filter_client,
    filter_in_work,
    filter_shipped,
    load_taz,
    normalize_invoice,
    parse_date,
    parse_numeric,
    row_qty,
    row_sale,
)

STALE_DAYS = 14
CLIENTS = ("Utair", "Белавиа")

FILL_HEADER = PatternFill("solid", fgColor="022F40")
FILL_STALE = PatternFill("solid", fgColor="F8D7DA")
FILL_WARN = PatternFill("solid", fgColor="FFF3CD")
FILL_ZEBRA = PatternFill("solid", fgColor="F5FAFB")
FONT_HEADER = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
FONT_TITLE = Font(bold=True, color="022F40", name="Calibri", size=16)
FONT_SUB = Font(color="5A5A5A", name="Calibri", size=11)
FONT_CELL = Font(name="Calibri", size=11)
THIN = Border(
    left=Side(style="thin", color="C4C4C4"),
    right=Side(style="thin", color="C4C4C4"),
    top=Side(style="thin", color="C4C4C4"),
    bottom=Side(style="thin", color="C4C4C4"),
)

HEADERS = [
    "Клиент",
    "№ счета",
    "P/N",
    "DESCRIPTION",
    "Status",
    "QTY",
    "Сумма, USD",
    "Остаток к оплате, USD",
    "Дата взятия в работу",
    "Статус без изменений с",
    "Дней без смены статуса",
    "Срез, где статус уже такой",
    "Комментарий",
]


def parse_report_date_from_name(path: Path) -> date | None:
    m = re.search(r"(\d{2})\.(\d{2})\.(\d{4})", path.name)
    if not m:
        return None
    d, mo, y = map(int, m.groups())
    try:
        return date(y, mo, d)
    except ValueError:
        return None


def load_status_snapshot(path: Path) -> pd.DataFrame:
    """Load «В работе» + «Отгружено» from a weekly status Excel."""
    frames: list[pd.DataFrame] = []
    for sheet in ("В работе", "Отгружено"):
        try:
            df = pd.read_excel(path, sheet_name=sheet)
        except ValueError:
            continue
        if df.empty:
            continue
        df = df[df.get("Status", pd.Series(dtype=object)).astype(str) != "ИТОГО"]
        rename = {}
        for src, dst in (
            ("№ счета", COL_INVOICE),
            ("P/N", COL_PN),
            ("DESCRIPTION", COL_DESC),
            ("Customer", COL_CUSTOMER),
            ("Комментарий", COL_COMMENT),
            ("Остаток к оплате", COL_BALANCE),
            ("Продажная, итого (без НДС)", COL_SALE),
            ("ЗАКАЗ ВЗЯТ В РАБОТУ", COL_ORDER_DATE),
        ):
            if src in df.columns:
                rename[src] = dst
        df = df.rename(columns=rename)
        if COL_INVOICE not in df.columns or COL_STATUS not in df.columns:
            continue
        if COL_PN not in df.columns:
            df[COL_PN] = ""
        if COL_DESC not in df.columns:
            df[COL_DESC] = ""
        frames.append(df)
    if not frames:
        return pd.DataFrame()
    out = pd.concat(frames, ignore_index=True)
    out = add_row_key(out)
    out["_status_norm"] = out[COL_STATUS].astype(str).str.strip()
    return out.drop_duplicates(subset="_key", keep="last")


def collect_history(output_dir: Path, clients: tuple[str, ...]) -> list[tuple[date, str, pd.DataFrame]]:
    """Return [(snap_date, client, snapshot_df), ...] sorted by date ascending."""
    items: list[tuple[date, str, pd.DataFrame]] = []
    for client in clients:
        for path in sorted(output_dir.glob(f"{client} статус *.xlsx")):
            snap_date = parse_report_date_from_name(path)
            if snap_date is None:
                continue
            snap = load_status_snapshot(path)
            if snap.empty:
                continue
            items.append((snap_date, client, snap))
    items.sort(key=lambda x: (x[1], x[0]))
    return items


def earliest_same_status(
    key: str,
    status: str,
    client: str,
    history: list[tuple[date, str, pd.DataFrame]],
    report_date: date,
) -> tuple[date | None, str]:
    """Oldest snapshot (for this client) where key had the same status.

    Walks from newest-before-today backwards; if status differs, stop.
    Returns (earliest_date_with_same_status, source_label).
    """
    client_snaps = [(d, df) for d, c, df in history if c == client and d < report_date]
    if not client_snaps:
        return None, ""

    # newest first
    client_snaps.sort(key=lambda x: x[0], reverse=True)
    earliest: date | None = None
    source = ""
    for snap_date, df in client_snaps:
        hit = df.loc[df["_key"] == key]
        if hit.empty:
            break
        prev_status = str(hit.iloc[0]["_status_norm"])
        if prev_status != status:
            break
        earliest = snap_date
        source = f"{client} статус {snap_date.strftime('%d.%m.%Y')}.xlsx"
    return earliest, source


def current_open_rows(taz: pd.DataFrame, client: str) -> pd.DataFrame:
    in_work = filter_in_work(taz, client)
    shipped = filter_shipped(taz, client)
    parts = [p for p in (in_work, shipped) if not p.empty]
    if not parts:
        return pd.DataFrame()
    out = pd.concat(parts, ignore_index=True)
    out = add_row_key(out)
    out["_status_norm"] = out[COL_STATUS].astype(str).str.strip()
    out["_client"] = client
    return out.drop_duplicates(subset="_key", keep="last")


def build_stale_table(
    taz: pd.DataFrame,
    history: list[tuple[date, str, pd.DataFrame]],
    report_date: date,
    stale_days: int = STALE_DAYS,
) -> pd.DataFrame:
    rows: list[dict] = []
    cutoff = report_date - timedelta(days=stale_days)
    for client in CLIENTS:
        cur = current_open_rows(filter_client(taz, client), client)
        if cur.empty:
            continue
        for _, row in cur.iterrows():
            key = row["_key"]
            status = row["_status_norm"]
            since, source = earliest_same_status(key, status, client, history, report_date)
            if since is None:
                continue
            if since > cutoff:
                continue
            days = (report_date - since).days
            rows.append(
                {
                    "Клиент": client,
                    "№ счета": row.get(COL_INVOICE),
                    "P/N": row.get(COL_PN),
                    "DESCRIPTION": row.get(COL_DESC),
                    "Status": status,
                    "QTY": row_qty(row),
                    "Сумма, USD": row_sale(row),
                    "Остаток к оплате, USD": parse_numeric(row.get(COL_BALANCE)),
                    "Дата взятия в работу": parse_date(row.get(COL_ORDER_DATE)),
                    "Статус без изменений с": since,
                    "Дней без смены статуса": days,
                    "Срез, где статус уже такой": source,
                    "Комментарий": row.get(COL_COMMENT, ""),
                    "_days": days,
                }
            )
    if not rows:
        return pd.DataFrame(columns=HEADERS)
    table = pd.DataFrame(rows)
    table = table.sort_values(["_days", "Клиент", "Сумма, USD"], ascending=[False, True, False])
    return table.drop(columns=["_days"]).reset_index(drop=True)


def write_excel(table: pd.DataFrame, out_path: Path, report_date: date, stale_days: int) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Без движения"
    ws.sheet_view.showGridLines = False

    widths = {
        "A": 12,
        "B": 14,
        "C": 18,
        "D": 28,
        "E": 14,
        "F": 8,
        "G": 14,
        "H": 16,
        "I": 16,
        "J": 16,
        "K": 14,
        "L": 36,
        "M": 40,
    }
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    ws.merge_cells("A1:M1")
    ws["A1"] = "FASTAIR — заказы без движения (Utair + Белавиа)"
    ws["A1"].font = FONT_TITLE

    ws.merge_cells("A2:M2")
    ws["A2"] = (
        f"Срез ТАЗ {report_date.strftime('%d.%m.%Y')} · "
        f"подсветка: статус не менялся ≥ {stale_days} дн. "
        f"(сравнение с еженедельными статусами в output/)"
    )
    ws["A2"].font = FONT_SUB

    ws.merge_cells("A3:M3")
    ws["A3"] = (
        f"Найдено позиций: {len(table)} · "
        "розовый = ≥28 дн. без смены статуса, жёлтый = 14–27 дн."
    )
    ws["A3"].font = FONT_SUB

    header_row = 5
    for col_idx, name in enumerate(HEADERS, start=1):
        cell = ws.cell(header_row, col_idx, name)
        cell.fill = FILL_HEADER
        cell.font = FONT_HEADER
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        cell.border = THIN

    for offset, rec in enumerate(table.to_dict(orient="records")):
        r_idx = header_row + 1 + offset
        days = int(rec["Дней без смены статуса"] or 0)
        row_fill = FILL_STALE if days >= 28 else FILL_WARN
        for c_idx, name in enumerate(HEADERS, start=1):
            val = rec.get(name)
            cell = ws.cell(r_idx, c_idx, val)
            cell.font = FONT_CELL
            cell.border = THIN
            cell.fill = row_fill
            if name in {"Сумма, USD", "Остаток к оплате, USD"} and isinstance(val, (int, float)):
                cell.number_format = '#,##0.00'
                cell.alignment = Alignment(horizontal="right")
            if name in {"Дата взятия в работу", "Статус без изменений с"} and val:
                cell.number_format = "DD.MM.YYYY"
            if name == "Дней без смены статуса":
                cell.alignment = Alignment(horizontal="right")

    ws.auto_filter.ref = f"A{header_row}:M{header_row + max(len(table), 1)}"
    ws.freeze_panes = f"A{header_row + 1}"

    # Summary sheet
    sum_ws = wb.create_sheet("Сводка", 0)
    sum_ws["A1"] = "Сводка: без движения ≥ 2 недель"
    sum_ws["A1"].font = FONT_TITLE
    sum_ws["A2"] = f"Дата отчёта: {report_date.strftime('%d.%m.%Y')}"
    sum_ws["A2"].font = FONT_SUB
    sum_ws.column_dimensions["A"].width = 28
    sum_ws.column_dimensions["B"].width = 14
    sum_ws.column_dimensions["C"].width = 18

    sum_ws["A4"] = "Клиент"
    sum_ws["B4"] = "Позиций"
    sum_ws["C4"] = "Сумма, USD"
    for col in range(1, 4):
        sum_ws.cell(4, col).fill = FILL_HEADER
        sum_ws.cell(4, col).font = FONT_HEADER

    if table.empty:
        sum_ws["A5"] = "Нет позиций"
    else:
        grouped = (
            table.groupby("Клиент", dropna=False)
            .agg(n=("P/N", "size"), sale=("Сумма, USD", "sum"))
            .reset_index()
        )
        for i, g in enumerate(grouped.itertuples(index=False), start=5):
            sum_ws.cell(i, 1, g[0])
            sum_ws.cell(i, 2, int(g[1]))
            cell = sum_ws.cell(i, 3, float(g[2]))
            cell.number_format = '#,##0.00'
        total_row = 5 + len(grouped)
        sum_ws.cell(total_row, 1, "ИТОГО").font = Font(bold=True)
        sum_ws.cell(total_row, 2, int(len(table))).font = Font(bold=True)
        tcell = sum_ws.cell(total_row, 3, float(table["Сумма, USD"].sum()))
        tcell.number_format = '#,##0.00'
        tcell.font = Font(bold=True)

    sum_ws["A10"] = "Правило"
    sum_ws["A10"].font = Font(bold=True, color="022F40")
    sum_ws["A11"] = (
        "Берём открытые позиции (в работе + отгружено с остатком к оплате) по Utair и Белавиа "
        "из текущего ТАЗ. Сверяем Status с прошлыми еженедельными Excel-срезами. "
        f"Если тот же статус уже был в срезе ≥ {stale_days} дн. назад — строка в отчёте."
    )
    sum_ws.merge_cells("A11:C13")
    sum_ws["A11"].alignment = Alignment(wrap_text=True, vertical="top")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--taz", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, default=Path("output"))
    parser.add_argument("--date", type=str, default=None, help="dd.mm.yyyy")
    parser.add_argument("--stale-days", type=int, default=STALE_DAYS)
    parser.add_argument(
        "--stem",
        default=None,
        help="output filename stem (default: stale_orders_DD.MM.YYYY)",
    )
    args = parser.parse_args()

    if args.date:
        d, m, y = map(int, args.date.split("."))
        report_date = date(y, m, d)
    else:
        m = re.search(r"(\d{2})\.(\d{2})\.(\d{4})", args.taz.name)
        if not m:
            raise SystemExit("Cannot infer date from TAZ filename; pass --date")
        d, mo, y = map(int, m.groups())
        report_date = date(y, mo, d)

    taz = load_taz(args.taz, excluded_invoicers=set())  # include ФЭ for Belavia
    # Still exclude SAMPLE etc. via load_taz; keep ФЭ
    history = collect_history(args.output_dir, CLIENTS)
    print(f"history snapshots: {len(history)}")
    for d, c, df in history:
        print(f"  {c} {d.isoformat()} n={len(df)}")

    table = build_stale_table(taz, history, report_date, stale_days=args.stale_days)
    stem = args.stem or f"stale_orders_{report_date.strftime('%d.%m.%Y')}"
    out = args.output_dir / f"{stem}.xlsx"
    write_excel(table, out, report_date, args.stale_days)
    print(f"stale rows: {len(table)}")
    if not table.empty:
        print(table.groupby("Клиент").size().to_string())
        print("top days:", int(table["Дней без смены статуса"].max()))
    print(f"wrote {out}")


if __name__ == "__main__":
    main()
