#!/usr/bin/env python3
"""Utair / Belavia: orders awaiting payment vs orders with no real movement.

Two sheets:
  • «ожидают оплаты» — Status = 4 FINISHED, unpaid client balance, FINISHED ≥ N days
  • «без движения» — not FINISHED / not CANCELLED; Status and fulfilment fields
    unchanged for ≥ N days across weekly status Excels (supplier, supplier pay,
    BA–BD logistics, comments, fact delivery date, delivery forecast).
    Client payment fields (AR–AU) are ignored.
"""

from __future__ import annotations

import argparse
import re
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from generate_report import (
    COL_BALANCE,
    COL_COMMENT,
    COL_CUSTOMER,
    COL_DELIVERY_ACTUAL,
    COL_DESC,
    COL_EXPECTED_ARRIVAL,
    COL_INVOICE,
    COL_ORDER_DATE,
    COL_PN,
    COL_QTY,
    COL_SALE,
    COL_STATUS,
    add_row_key,
    filter_client,
    filter_in_work,
    filter_shipped,
    load_taz,
    parse_date,
    parse_numeric,
    row_qty,
    row_sale,
)

STALE_DAYS = 14
CLIENTS = ("Utair", "Белавиа")
STATUS_FINISHED = "4 FINISHED"
STATUS_CANCELLED = "5 CANCELLED"

# Movement fields (TAZ names). Client payment AR–AU intentionally omitted.
COL_SUPPLIER = "Поставщик"
COL_ROOT_SUPPLIER = "Root supplier"
COL_SUP_PAY_DATE = "Дата оплаты поставщику"
COL_SUP_PAY_AMT = "Сумма оплаты поставщику"
COL_SUP_PAY_BAL = "Остаток к оплате поставщику"
COL_MOVE_START = "Дата начала движения "
COL_AWB = "НОМЕР AWB"
COL_CUSTOMS_INV = "НОМЕР ИНВОЙСА ДЛЯ ТАМОЖНИ"
COL_DAYS_TO_DELIVER = "Дней на поставку (ЧИСЛО)"
COL_LEAD_TIME = "Lead time"
COL_DEADLINE = "КРАЙНЯЯ ДАТА ПОСТАВКИ"

# Weekly Excel may rename some headers
SNAPSHOT_RENAME = {
    "№ счета": COL_INVOICE,
    "счет": COL_INVOICE,
    "P/N": COL_PN,
    "DESCRIPTION": COL_DESC,
    "Customer": COL_CUSTOMER,
    "Комментарий": COL_COMMENT,
    "Комментарии": COL_COMMENT,
    "Остаток к оплате": COL_BALANCE,
    "Продажная, итого (без НДС)": COL_SALE,
    "Продажная, итого": COL_SALE,
    "ЗАКАЗ ВЗЯТ В РАБОТУ": COL_ORDER_DATE,
    "Дата начала движения": COL_MOVE_START,  # without trailing space
    "Дней на поставку": COL_DAYS_TO_DELIVER,
}

MOVEMENT_FIELDS: list[tuple[str, str]] = [
    ("Status", COL_STATUS),
    ("Поставщик", COL_SUPPLIER),
    ("Root supplier", COL_ROOT_SUPPLIER),
    ("Оплата поставщику (дата)", COL_SUP_PAY_DATE),
    ("Оплата поставщику (сумма)", COL_SUP_PAY_AMT),
    ("Остаток поставщику", COL_SUP_PAY_BAL),
    ("Дата начала движения", COL_MOVE_START),
    ("AWB", COL_AWB),
    ("Ожидаемая дата прихода", COL_EXPECTED_ARRIVAL),
    ("Инвойс для таможни", COL_CUSTOMS_INV),
    ("Факт. дата поставки", COL_DELIVERY_ACTUAL),
    ("Дней на поставку", COL_DAYS_TO_DELIVER),
    ("Lead time", COL_LEAD_TIME),
    ("Крайняя дата поставки", COL_DEADLINE),
    ("Комментарий", COL_COMMENT),
]

FILL_HEADER = PatternFill("solid", fgColor="022F40")
FILL_STALE = PatternFill("solid", fgColor="F8D7DA")
FILL_WARN = PatternFill("solid", fgColor="FFF3CD")
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

HEADERS_AWAITING = [
    "Клиент",
    "№ счета",
    "P/N",
    "DESCRIPTION",
    "Status",
    "QTY",
    "Сумма, USD",
    "Остаток к оплате, USD",
    "Факт. дата поставки",
    "FINISHED с",
    "Дней в FINISHED",
    "Срез",
    "Комментарий",
]

HEADERS_STALE = [
    "Клиент",
    "№ счета",
    "P/N",
    "DESCRIPTION",
    "Status",
    "QTY",
    "Сумма, USD",
    "Остаток к оплате, USD",
    "Дата взятия в работу",
    "Без движения с",
    "Дней без движения",
    "Поставщик",
    "Вывод",
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


def _is_empty(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and pd.isna(value):
        return True
    if isinstance(value, str) and not value.strip():
        return True
    if value is pd.NaT:
        return True
    return False


def norm_text(value: Any) -> str:
    if _is_empty(value):
        return ""
    text = str(value).strip()
    text = re.sub(r"\s+", " ", text)
    if text.lower() in {"nan", "nat", "none", "-"}:
        return ""
    return text


def norm_date(value: Any) -> str:
    if _is_empty(value):
        return ""
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    parsed = parse_date(value)
    if parsed is not None:
        return parsed.isoformat()
    # Excel serial / odd values → keep as cleaned text so changes still show
    return norm_text(value)


def norm_num(value: Any) -> str:
    if _is_empty(value):
        return ""
    num = parse_numeric(value)
    if num == 0.0 and norm_text(value) in {"", "0", "0.0"}:
        return "0"
    # preserve distinguishability of tiny differences
    return f"{num:.4f}".rstrip("0").rstrip(".")


def normalize_snapshot_columns(df: pd.DataFrame) -> pd.DataFrame:
    rename = {src: dst for src, dst in SNAPSHOT_RENAME.items() if src in df.columns}
    # expected arrival long header may differ slightly
    for col in list(df.columns):
        cl = str(col)
        if cl.startswith("ОЖИДАЕМАЯ ДАТА ПРИХОДА"):
            rename[col] = COL_EXPECTED_ARRIVAL
        elif cl.startswith("ФАКТИЧЕСКАЯ ДАТА ПОСТАВКИ"):
            rename[col] = COL_DELIVERY_ACTUAL
        elif cl.startswith("Дата начала движения"):
            rename[col] = COL_MOVE_START
        elif cl.startswith("ЗАКАЗ ВЗЯТ В РАБОТУ"):
            rename[col] = COL_ORDER_DATE
    out = df.rename(columns=rename)
    # ensure movement cols exist
    for _, col in MOVEMENT_FIELDS:
        if col not in out.columns:
            out[col] = ""
    if COL_PN not in out.columns:
        out[COL_PN] = ""
    if COL_DESC not in out.columns:
        out[COL_DESC] = ""
    if COL_BALANCE not in out.columns:
        out[COL_BALANCE] = None
    if COL_SALE not in out.columns:
        out[COL_SALE] = None
    if COL_ORDER_DATE not in out.columns:
        out[COL_ORDER_DATE] = None
    return out


def movement_signature(row: pd.Series) -> tuple[str, ...]:
    parts: list[str] = []
    for _, col in MOVEMENT_FIELDS:
        val = row.get(col)
        if col in {
            COL_SUP_PAY_DATE,
            COL_MOVE_START,
            COL_EXPECTED_ARRIVAL,
            COL_DELIVERY_ACTUAL,
            COL_DEADLINE,
        }:
            parts.append(norm_date(val))
        elif col in {COL_SUP_PAY_AMT, COL_SUP_PAY_BAL, COL_DAYS_TO_DELIVER}:
            parts.append(norm_num(val))
        else:
            parts.append(norm_text(val))
    return tuple(parts)


def movement_map(row: pd.Series) -> dict[str, str]:
    sig = movement_signature(row)
    return {label: sig[i] for i, (label, _) in enumerate(MOVEMENT_FIELDS)}


def diff_labels(cur: dict[str, str], prev: dict[str, str]) -> list[str]:
    return [k for k in cur if cur.get(k, "") != prev.get(k, "")]


def load_status_snapshot(path: Path) -> pd.DataFrame:
    frames: list[pd.DataFrame] = []
    for sheet in ("В работе", "Отгружено"):
        try:
            df = pd.read_excel(path, sheet_name=sheet)
        except ValueError:
            continue
        if df.empty:
            continue
        status_col = "Status" if "Status" in df.columns else COL_STATUS
        df = df[df[status_col].astype(str) != "ИТОГО"]
        df = normalize_snapshot_columns(df)
        if COL_INVOICE not in df.columns or COL_STATUS not in df.columns:
            continue
        frames.append(df)
    if not frames:
        return pd.DataFrame()
    out = pd.concat(frames, ignore_index=True)
    out = add_row_key(out)
    out["_status_norm"] = out[COL_STATUS].astype(str).str.strip()
    out["_sig"] = out.apply(movement_signature, axis=1)
    out["_sig_map"] = out.apply(movement_map, axis=1)
    return out.drop_duplicates(subset="_key", keep="last")


def collect_history(output_dir: Path) -> list[tuple[date, str, pd.DataFrame]]:
    items: list[tuple[date, str, pd.DataFrame]] = []
    for client in CLIENTS:
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
    client_snaps = [(d, df) for d, c, df in history if c == client and d < report_date]
    client_snaps.sort(key=lambda x: x[0], reverse=True)
    earliest: date | None = None
    source = ""
    for snap_date, df in client_snaps:
        hit = df.loc[df["_key"] == key]
        if hit.empty:
            break
        if str(hit.iloc[0]["_status_norm"]) != status:
            break
        earliest = snap_date
        source = f"{client} статус {snap_date.strftime('%d.%m.%Y')}.xlsx"
    return earliest, source


def earliest_same_movement(
    key: str,
    signature: tuple[str, ...],
    client: str,
    history: list[tuple[date, str, pd.DataFrame]],
    report_date: date,
) -> tuple[date | None, str, list[str]]:
    """Oldest snap where the full movement signature still matches.

    Returns (since, source_label, fields_that_changed_just_before).
    """
    client_snaps = [(d, df) for d, c, df in history if c == client and d < report_date]
    client_snaps.sort(key=lambda x: x[0], reverse=True)
    earliest: date | None = None
    source = ""
    changed_before: list[str] = []
    cur_map = {label: signature[i] for i, (label, _) in enumerate(MOVEMENT_FIELDS)}
    for snap_date, df in client_snaps:
        hit = df.loc[df["_key"] == key]
        if hit.empty:
            break
        prev_sig = hit.iloc[0]["_sig"]
        if prev_sig != signature:
            prev_map = hit.iloc[0]["_sig_map"]
            changed_before = diff_labels(cur_map, prev_map)
            break
        earliest = snap_date
        source = f"{client} статус {snap_date.strftime('%d.%m.%Y')}.xlsx"
    return earliest, source, changed_before


def _format_awb(value: Any) -> str:
    """AWB is usually text; Excel sometimes stores a date there."""
    if _is_empty(value):
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d.%m.%Y")
    if isinstance(value, date):
        return value.strftime("%d.%m.%Y")
    text = norm_text(value)
    if re.match(r"^\d{4}-\d{2}-\d{2}", text):
        try:
            return date.fromisoformat(text[:10]).strftime("%d.%m.%Y")
        except ValueError:
            return text
    return text


def diagnose_pause(row: pd.Series, report_date: date, days_frozen: int) -> str:
    """Human-readable reason why fulfilment looks stuck."""
    status = norm_text(row.get(COL_STATUS)) or "?"
    order_date = parse_date(row.get(COL_ORDER_DATE))
    days_open = (report_date - order_date).days if order_date else None

    supplier = norm_text(row.get(COL_SUPPLIER)) or norm_text(row.get(COL_ROOT_SUPPLIER))
    pay_date = norm_date(row.get(COL_SUP_PAY_DATE))
    pay_amt = norm_num(row.get(COL_SUP_PAY_AMT))
    pay_bal = parse_numeric(row.get(COL_SUP_PAY_BAL))
    has_pay = bool(pay_date or pay_amt)

    move = norm_date(row.get(COL_MOVE_START))
    awb = _format_awb(row.get(COL_AWB))
    eta = parse_date(row.get(COL_EXPECTED_ARRIVAL))
    fact = parse_date(row.get(COL_DELIVERY_ACTUAL))
    comment = norm_text(row.get(COL_COMMENT))
    awb_u = awb.upper()
    age = days_open if days_open is not None else days_frozen
    has_logistics = bool(move or awb or eta is not None)

    def _pay_label() -> str:
        bits: list[str] = []
        if pay_date:
            bits.append(date.fromisoformat(pay_date).strftime("%d.%m.%Y"))
        if pay_amt:
            bits.append(f"${pay_amt}")
        return ", ".join(bits) if bits else "да"

    def _logistics_bits() -> str:
        parts: list[str] = []
        if move:
            parts.append(f"движение с {date.fromisoformat(move).strftime('%d.%m.%Y')}")
        if awb:
            parts.append(f"AWB {awb}")
        if eta is not None:
            parts.append(f"ETA {eta.strftime('%d.%m.%Y')}")
        return "; ".join(parts)

    # 1) Fact delivery present → we already shipped the part
    if fact is not None:
        return (
            f"Факт. дата поставки {fact.strftime('%d.%m.%Y')} уже стоит, "
            f"но статус всё ещё «{status}» — отгрузили, а карточку не закрыли "
            f"(не FINISHED / хвосты). Поля без обновлений {days_frozen} дн."
        )

    # 2) No supplier assigned
    if not supplier:
        age_bit = f" В работе уже {days_open} дн." if days_open is not None else ""
        return (
            f"Поставщик не выбран — закупка не стартовала, пауза на этапе назначения.{age_bit}"
        )

    unpaid_note = ""
    if not has_pay and age >= 7:
        unpaid_note = f" Оплаты поставщику нет уже при {age} дн. в работе."

    # 3) Logistics signals — diagnose route/ETA even if supplier unpaid
    if has_logistics:
        if awb and any(
            k in awb_u for k in ("HOLD", "UNCONFIRMED", "PENDING", "TBA", "WAIT", "UNSOLVED")
        ):
            pay_bit = f"Оплата поставщику есть ({_pay_label()}). " if has_pay else "Оплаты поставщику нет. "
            return (
                f"{pay_bit}Логистика зависла: AWB «{awb}». Без обновлений {days_frozen} дн."
            )
        if eta is not None and eta < report_date:
            return (
                f"ETA {eta.strftime('%d.%m.%Y')} уже прошла, факт. поставки нет"
                f" ({_logistics_bits()}). Логистика просрочена / груз не доехал. "
                f"Пауза {days_frozen} дн.{unpaid_note}"
            )
        pay_bit = (
            f"Поставщику оплатили ({_pay_label()}), "
            if has_pay
            else "Оплаты поставщику в ТАЗ нет, "
        )
        return (
            f"{pay_bit}груз в логистике ({_logistics_bits()}), "
            f"но апдейтов нет {days_frozen} дн. — логистика без движения.{unpaid_note}"
        )

    # 4) Supplier known, but we are not paying them (and no logistics yet)
    if not has_pay:
        if age >= 14:
            return (
                f"Поставщик есть ({supplier[:40]}), оплаты поставщику нет, "
                f"хотя заказ в работе {age} дн. — пауза из‑за неоплаты поставщику."
            )
        if age >= 7:
            return (
                f"Поставщик ({supplier[:40]}) назначен, оплаты поставщику ещё нет "
                f"({age} дн. в работе) — ждём или держим оплату поставщику."
            )
        return (
            f"Поставщик ({supplier[:40]}) есть, оплаты поставщику пока нет "
            f"(ранняя стадия, {age} дн. в работе)."
        )

    # 5) Partial supplier payment left hanging
    if pay_bal is not None and pay_bal > 1:
        return (
            f"Поставщику платили ({_pay_label()}), но остаток поставщику "
            f"${pay_bal:,.0f} всё ещё висит — недоплата/хвост по поставщику, "
            f"без движения {days_frozen} дн."
        )

    # 6) Paid supplier, logistics never started
    return (
        f"Поставщику оплатили ({_pay_label()}), но логистика пустая "
        f"(нет даты начала движения, AWB и ETA) — пауза на отгрузке/логистике."
    )


def current_open_rows(taz: pd.DataFrame, client: str) -> pd.DataFrame:
    in_work = filter_in_work(taz, client)
    shipped = filter_shipped(taz, client)
    parts = [p for p in (in_work, shipped) if not p.empty]
    if not parts:
        return pd.DataFrame()
    out = pd.concat(parts, ignore_index=True)
    out = normalize_snapshot_columns(out)
    out = add_row_key(out)
    out["_status_norm"] = out[COL_STATUS].astype(str).str.strip()
    out["_client"] = client
    out["_sig"] = out.apply(movement_signature, axis=1)
    return out.drop_duplicates(subset="_key", keep="last")


def build_tables(
    taz: pd.DataFrame,
    history: list[tuple[date, str, pd.DataFrame]],
    report_date: date,
    stale_days: int = STALE_DAYS,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    awaiting_rows: list[dict] = []
    stale_rows: list[dict] = []
    cutoff = report_date - timedelta(days=stale_days)

    for client in CLIENTS:
        cur = current_open_rows(filter_client(taz, client), client)
        if cur.empty:
            continue
        for _, row in cur.iterrows():
            key = row["_key"]
            status = row["_status_norm"]
            if status == STATUS_CANCELLED or status.upper().endswith("CANCELLED"):
                continue

            if status == STATUS_FINISHED:
                since, source = earliest_same_status(
                    key, status, client, history, report_date
                )
                if since is None or since > cutoff:
                    continue
                awaiting_rows.append(
                    {
                        "Клиент": client,
                        "№ счета": row.get(COL_INVOICE),
                        "P/N": row.get(COL_PN),
                        "DESCRIPTION": row.get(COL_DESC),
                        "Status": status,
                        "QTY": row_qty(row),
                        "Сумма, USD": row_sale(row),
                        "Остаток к оплате, USD": parse_numeric(row.get(COL_BALANCE)),
                        "Факт. дата поставки": parse_date(row.get(COL_DELIVERY_ACTUAL)),
                        "FINISHED с": since,
                        "Дней в FINISHED": (report_date - since).days,
                        "Срез": source,
                        "Комментарий": row.get(COL_COMMENT, ""),
                        "_days": (report_date - since).days,
                    }
                )
                continue

            # «без движения» — deep signature
            since, source, _changed = earliest_same_movement(
                key, row["_sig"], client, history, report_date
            )
            if since is None or since > cutoff:
                continue
            days = (report_date - since).days
            stale_rows.append(
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
                    "Без движения с": since,
                    "Дней без движения": days,
                    "Поставщик": norm_text(row.get(COL_SUPPLIER))
                    or norm_text(row.get(COL_ROOT_SUPPLIER)),
                    "Вывод": diagnose_pause(row, report_date, days),
                    "_days": days,
                }
            )

    def _finalize(rows: list[dict], headers: list[str]) -> pd.DataFrame:
        if not rows:
            return pd.DataFrame(columns=headers)
        table = pd.DataFrame(rows)
        table = table.sort_values(
            ["_days", "Клиент", "Сумма, USD"], ascending=[False, True, False]
        )
        return table.drop(columns=["_days"]).reset_index(drop=True)

    return _finalize(awaiting_rows, HEADERS_AWAITING), _finalize(stale_rows, HEADERS_STALE)


def _write_sheet(
    ws,
    title: str,
    subtitle: str,
    note: str,
    headers: list[str],
    table: pd.DataFrame,
    days_col: str,
) -> None:
    ws.sheet_view.showGridLines = False
    last_col = len(headers)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    ws["A1"] = title
    ws["A1"].font = FONT_TITLE
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_col)
    ws["A2"] = subtitle
    ws["A2"].font = FONT_SUB
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=last_col)
    ws["A3"] = note
    ws["A3"].font = FONT_SUB

    header_row = 5
    for col_idx, name in enumerate(headers, start=1):
        cell = ws.cell(header_row, col_idx, name)
        cell.fill = FILL_HEADER
        cell.font = FONT_HEADER
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        cell.border = THIN

    money_cols = {"Сумма, USD", "Остаток к оплате, USD"}
    date_cols = {
        "Факт. дата поставки",
        "FINISHED с",
        "Дата взятия в работу",
        "Без движения с",
        "Дата начала движения",
        "Ожидаемая дата прихода",
    }

    for offset, rec in enumerate(table.to_dict(orient="records")):
        r_idx = header_row + 1 + offset
        days = int(rec.get(days_col) or 0)
        row_fill = FILL_STALE if days >= 28 else FILL_WARN
        for c_idx, name in enumerate(headers, start=1):
            val = rec.get(name)
            cell = ws.cell(r_idx, c_idx, val)
            cell.font = FONT_CELL
            cell.border = THIN
            cell.fill = row_fill
            if name in money_cols and isinstance(val, (int, float)):
                cell.number_format = "#,##0.00"
                cell.alignment = Alignment(horizontal="right")
            if name in date_cols and val:
                cell.number_format = "DD.MM.YYYY"
            if name == days_col:
                cell.alignment = Alignment(horizontal="right")
            if name == "Вывод":
                cell.alignment = Alignment(wrap_text=True, vertical="top")
        if "Вывод" in headers:
            ws.row_dimensions[r_idx].height = 48

    # widths
    default_w = {
        "Клиент": 12,
        "№ счета": 14,
        "P/N": 16,
        "DESCRIPTION": 26,
        "Status": 12,
        "QTY": 6,
        "Сумма, USD": 12,
        "Остаток к оплате, USD": 14,
        "Дата взятия в работу": 14,
        "Без движения с": 14,
        "Дней без движения": 12,
        "FINISHED с": 12,
        "Дней в FINISHED": 12,
        "Комментарий": 36,
        "Вывод": 62,
        "Срез": 32,
        "Поставщик": 20,
        "Факт. дата поставки": 14,
    }
    for idx, name in enumerate(headers, start=1):
        from openpyxl.utils import get_column_letter

        letter = get_column_letter(idx)
        ws.column_dimensions[letter].width = default_w.get(name, 14)

    end = header_row + max(len(table), 1)
    ws.auto_filter.ref = f"A{header_row}:{ws.cell(header_row, last_col).column_letter}{end}"
    ws.freeze_panes = f"A{header_row + 1}"


def write_excel(
    awaiting: pd.DataFrame,
    stale: pd.DataFrame,
    out_path: Path,
    report_date: date,
    stale_days: int,
) -> None:
    wb = Workbook()

    # Summary first
    sum_ws = wb.active
    sum_ws.title = "Сводка"
    sum_ws["A1"] = "FASTAIR — без движения / ожидают оплаты"
    sum_ws["A1"].font = FONT_TITLE
    sum_ws["A2"] = f"Срез ТАЗ {report_date.strftime('%d.%m.%Y')} · порог {stale_days} дн."
    sum_ws["A2"].font = FONT_SUB
    sum_ws.column_dimensions["A"].width = 28
    sum_ws.column_dimensions["B"].width = 14
    sum_ws.column_dimensions["C"].width = 18

    sum_ws["A4"] = "Вкладка"
    sum_ws["B4"] = "Позиций"
    sum_ws["C4"] = "Сумма, USD"
    for col in range(1, 4):
        sum_ws.cell(4, col).fill = FILL_HEADER
        sum_ws.cell(4, col).font = FONT_HEADER

    def _block(start_row: int, label: str, table: pd.DataFrame) -> int:
        sum_ws.cell(start_row, 1, label)
        if table.empty:
            sum_ws.cell(start_row, 2, 0)
            sum_ws.cell(start_row, 3, 0).number_format = "#,##0.00"
            return start_row + 1
        sum_ws.cell(start_row, 2, len(table))
        cell = sum_ws.cell(start_row, 3, float(table["Сумма, USD"].sum()))
        cell.number_format = "#,##0.00"
        r = start_row + 1
        for client, grp in table.groupby("Клиент"):
            sum_ws.cell(r, 1, f"  · {client}")
            sum_ws.cell(r, 2, len(grp))
            c = sum_ws.cell(r, 3, float(grp["Сумма, USD"].sum()))
            c.number_format = "#,##0.00"
            r += 1
        return r

    r = _block(5, "ожидают оплаты (FINISHED)", awaiting)
    r = _block(r + 1, "без движения", stale)

    sum_ws.cell(r + 1, 1, "Правила").font = Font(bold=True, color="022F40")
    sum_ws.cell(
        r + 2,
        1,
        "«ожидают оплаты»: только 4 FINISHED с остатком к оплате клиентом; "
        f"статус FINISHED уже был в срезе ≥ {stale_days} дн. назад.",
    )
    sum_ws.merge_cells(start_row=r + 2, start_column=1, end_row=r + 3, end_column=3)
    sum_ws.cell(r + 2, 1).alignment = Alignment(wrap_text=True, vertical="top")
    sum_ws.cell(
        r + 4,
        1,
        "«без движения»: все статусы кроме FINISHED и CANCELLED; "
        "не менялись Status, поставщик, оплата поставщику, BA–BD (движение/AWB/ETA/инвойс), "
        "факт. дата поставки, прогноз сроков, комментарий — ≥ "
        f"{stale_days} дн. Оплаты клиента игнорируются. "
        "Появление факт. даты поставки = отгрузка (= движение).",
    )
    sum_ws.merge_cells(start_row=r + 4, start_column=1, end_row=r + 6, end_column=3)
    sum_ws.cell(r + 4, 1).alignment = Alignment(wrap_text=True, vertical="top")
    sum_ws["A3"] = "розовый ≥28 дн. · жёлтый 14–27 дн."
    sum_ws["A3"].font = FONT_SUB

    await_ws = wb.create_sheet("ожидают оплаты")
    _write_sheet(
        await_ws,
        "Ожидают оплаты — FINISHED с остатком",
        f"ТАЗ {report_date.strftime('%d.%m.%Y')} · FINISHED без смены статуса ≥ {stale_days} дн.",
        f"Позиций: {len(awaiting)} · розовый ≥28 дн., жёлтый 14–27 дн.",
        HEADERS_AWAITING,
        awaiting,
        "Дней в FINISHED",
    )

    stale_ws = wb.create_sheet("без движения")
    _write_sheet(
        stale_ws,
        "Без движения — исполнение не двигается",
        f"ТАЗ {report_date.strftime('%d.%m.%Y')} · Status + поставщик/оплата пост./BA–BD/факт/ETA/коммент. "
        f"без изменений ≥ {stale_days} дн. (оплаты клиента не считаем)",
        f"Позиций: {len(stale)} · розовый ≥28 дн., жёлтый 14–27 дн.",
        HEADERS_STALE,
        stale,
        "Дней без движения",
    )

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


def generate_stale_report(
    taz_path: Path,
    output_dir: Path,
    report_date: date,
    stale_days: int = STALE_DAYS,
    stem: str | None = None,
) -> Path:
    """Build stale/awaiting workbook into output_dir; return path written."""
    taz = load_taz(taz_path, excluded_invoicers=set())  # include ФЭ for Belavia
    history = collect_history(output_dir)
    print(f"history snapshots: {len(history)}")
    for d, c, df in history:
        print(f"  {c} {d.isoformat()} n={len(df)}")

    awaiting, stale = build_tables(taz, history, report_date, stale_days=stale_days)
    out_stem = stem or f"stale_orders_{report_date.strftime('%d.%m.%Y')}"
    out = output_dir / f"{out_stem}.xlsx"
    write_excel(awaiting, stale, out, report_date, stale_days)

    print(f"awaiting payment: {len(awaiting)}")
    if not awaiting.empty:
        print(awaiting.groupby("Клиент").size().to_string())
    print(f"no movement: {len(stale)}")
    if not stale.empty:
        print(stale.groupby("Клиент").size().to_string())
        print(stale.groupby("Status").size().to_string())
        print("top days:", int(stale["Дней без движения"].max()))
    print(f"wrote {out}")
    return out


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--taz", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, default=Path("output"))
    parser.add_argument("--date", type=str, default=None, help="dd.mm.yyyy")
    parser.add_argument("--stale-days", type=int, default=STALE_DAYS)
    parser.add_argument("--stem", default=None)
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

    generate_stale_report(
        args.taz,
        args.output_dir,
        report_date,
        stale_days=args.stale_days,
        stem=args.stem,
    )


if __name__ == "__main__":
    main()
