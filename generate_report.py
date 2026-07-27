#!/usr/bin/env python3
"""Generate a client status Excel report from a TAZ export."""

from __future__ import annotations

import argparse
import re
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

# Internal TAZ column names
COL_INVOICE = "Номер счета"
COL_INVOICER = "Invoicer"
COL_STATUS = "Status"
COL_CUSTOMER = "Customer"
EXCLUDED_INVOICERS = {"ФЭ"}
COL_COMMENT = "Комментарии"
COL_PURCHASE = "Закупка, итого"
COL_SALE = "Продажная, итого"
COL_TRANSPORT = (
    "Стоимость доставки ПЛАН, за весь счет! Если в счете несколько строк, "
    'то "размазываем" равномерно планируюмую стоиомость транспорта на все позиции из счета.'
)
COL_FEE = "Transaction fee"
COL_PAY1 = "Оплачено клиентом, USD (cntr+shift+V)\nПервая группа платежей"
COL_PAY1_DATE = "Дата оплаты Клиентом\n\nПервая группа платежей"
COL_PAY2 = "Оплачено клиентом, USD (cntr+shift+V)\nЗавершающий платеж"
COL_PAY2_DATE = "Дата оплаты Клиентом\n\nЗавершающий платеж"
COL_BALANCE = "Остаток к оплате клиентом, USD"
COL_PN = "p/n"
COL_DESC = "DESCRIPTION"
COL_DELIVERY_ACTUAL = "ФАКТИЧЕСКАЯ ДАТА ПОСТАВКИ (СОГЛАСНО УСЛОВИЯМ ПОСТАВКИ)"
COL_ORDER_DATE = "ЗАКАЗ ВЗЯТ В РАБОТУ (ДАТА) ОТ КЛИЕНТА"

STATUSES_IN_WORK = {"1 NOT PAID", "2 PAID", "6 TROUBLE"}
STATUSES_SHIPPED = {"3 SHIPPED", "4 FINISHED"}

COLUMN_RENAME = {
    COL_INVOICE: "№ счета",
    COL_COMMENT: "Комментарий",
    "Sgmt": "Номер группы",
    "p/n": "P/N",
    "QTY IN PO": "QTY",
    "Units (UOM) for customer": "UOM",
    COL_ORDER_DATE: "ЗАКАЗ ВЗЯТ В РАБОТУ",
    "Дней на поставку (ЧИСЛО)": "Дней на поставку",
    COL_SALE: "Продажная, итого (без НДС)",
    COL_TRANSPORT: "Стоимость доставки ПЛАН, за весь счет! ",
    "Пошлина,сбор (ФАКТ) USD.": "Пошлина, сбор(ФАКТ) USD",
    "Дата оплаты Клиентом\n\nПервая группа платежей": "Дата оплаты\n\nПервая группа платежей",
    COL_PAY1: "Оплачено\n\nПервая группа платежей",
    "Дата оплаты Клиентом\n\nЗавершающий платеж": "Дата оплаты\n\nЗавершающий платеж",
    COL_PAY2: "Оплачено\n\nЗавершающий платеж",
    COL_BALANCE: "Остаток к оплате",
}

TOTAL_HEADERS = [
    "Закупка",
    "Транспорт",
    "FEE",
    "Продажа",
    "Маржа",
    "Оплачено клиентом",
    "К оплате клиентом",
]

# --- Styling ---
COLOR_NAVY = "1F3864"
COLOR_BLUE = "2F5597"
COLOR_LIGHT_BLUE = "D9E2F3"
COLOR_GREEN = "548235"
COLOR_LIGHT_GREEN = "E2EFDA"
COLOR_ORANGE = "C55A11"
COLOR_LIGHT_ORANGE = "FCE4D6"
COLOR_YELLOW = "FFF2CC"
COLOR_RED = "C00000"
COLOR_LIGHT_RED = "F8CBAD"
COLOR_WHITE = "FFFFFF"
COLOR_ZEBRA = "F7F9FC"
COLOR_SUBTOTAL = "FFF2CC"

FONT_TITLE = Font(name="Calibri", size=16, bold=True, color=COLOR_NAVY)
FONT_SUBTITLE = Font(name="Calibri", size=11, color="595959")
FONT_SECTION = Font(name="Calibri", size=12, bold=True, color=COLOR_WHITE)
FONT_HEADER = Font(name="Calibri", size=10, bold=True, color=COLOR_WHITE)
FONT_BODY = Font(name="Calibri", size=10)
FONT_BODY_BOLD = Font(name="Calibri", size=10, bold=True)
FONT_KPI = Font(name="Calibri", size=14, bold=True, color=COLOR_NAVY)

FILL_SECTION_WORK = PatternFill("solid", fgColor=COLOR_BLUE)
FILL_SECTION_SHIPPED = PatternFill("solid", fgColor=COLOR_GREEN)
FILL_HEADER = PatternFill("solid", fgColor=COLOR_NAVY)
FILL_SUBTOTAL = PatternFill("solid", fgColor=COLOR_SUBTOTAL)
FILL_TOTAL_ROW = PatternFill("solid", fgColor=COLOR_LIGHT_BLUE)
FILL_KPI = PatternFill("solid", fgColor=COLOR_LIGHT_BLUE)
FILL_ALERT = PatternFill("solid", fgColor=COLOR_LIGHT_ORANGE)

THIN = Side(style="thin", color="B4C6E7")
BORDER_THIN = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BORDER_BOTTOM = Border(bottom=Side(style="medium", color=COLOR_NAVY))

ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")

NUM_FMT = '#,##0.00'
DATE_FMT = "DD.MM.YYYY"

CURRENCY_COLUMNS = {
    "Закупка, ед.",
    "Закупка, итого",
    "Продажная, ед.",
    "Продажная, итого (без НДС)",
    "Стоимость доставки ПЛАН, за весь счет! ",
    "Стоимость доставки факт",
    "Transaction fee",
    "Пошлина, сбор(ФАКТ) USD",
    "СВХ стоимость (ФАКТ)",
    "Оплачено\n\nПервая группа платежей",
    "Оплачено\n\nЗавершающий платеж",
    "Остаток к оплате",
}

DATE_COLUMNS = {
    "ЗАКАЗ ВЗЯТ В РАБОТУ",
    "КРАЙНЯЯ ДАТА ПОСТАВКИ",
    "ФАКТИЧЕСКАЯ ДАТА ПОСТАВКИ (СОГЛАСНО УСЛОВИЯМ ПОСТАВКИ)",
    "Дата оплаты\n\nПервая группа платежей",
    "Дата оплаты\n\nЗавершающий платеж",
}

STATUS_FILLS = {
    "1 NOT PAID": PatternFill("solid", fgColor="FFF2CC"),
    "2 PAID": PatternFill("solid", fgColor="E2F0D9"),
    "3 SHIPPED": PatternFill("solid", fgColor="C6E0B4"),
    "4 FINISHED": PatternFill("solid", fgColor="BDD7EE"),
    "6 TROUBLE": PatternFill("solid", fgColor="F8CBAD"),
}

COLUMN_WIDTHS = {
    "№ счета": 14,
    "Invoicer": 8,
    "Менеджер продажи": 14,
    "Менеджер закупки": 14,
    "Status": 12,
    "Комментарий": 24,
    "Customer": 12,
    "Chnl": 8,
    "Номер группы": 10,
    "ТИП ВС (Продажи)": 12,
    "P/N": 16,
    "ALT P/N": 14,
    "DESCRIPTION": 28,
    "QTY": 8,
    "UOM": 6,
    "Category": 12,
    "ЗАКАЗ ВЗЯТ В РАБОТУ": 12,
    "Дней на поставку": 10,
    "Lead time": 10,
    "Destination": 12,
    "КРАЙНЯЯ ДАТА ПОСТАВКИ": 12,
    "ФАКТИЧЕСКАЯ ДАТА ПОСТАВКИ (СОГЛАСНО УСЛОВИЯМ ПОСТАВКИ)": 12,
    "Поставщик": 16,
    "Root supplier": 18,
    "PO #": 14,
    "Закупка, итого": 14,
    "Продажная, итого (без НДС)": 16,
    "Остаток к оплате": 14,
    "Тип оплаты": 18,
}

WEEKLY_HEADERS = ["№ счета", "P/N", "DESCRIPTION", "Status", "Дата", "Сумма, USD", "Примечание"]
WEEKLY_PAID_HEADERS = [
    "№ счета",
    "P/N",
    "DESCRIPTION",
    "Status",
    "Сумма, USD",
    "Оплата AS",
    "Дата AT",
    "Оплата AU",
    "Остаток AV",
    "Проверка",
]


@dataclass
class WeeklySection:
    title: str
    count: int
    total: float
    rows: list[dict[str, Any]]
    headers: list[str]


@dataclass
class WeeklySummary:
    week_start: date
    week_end: date
    new_orders: WeeklySection
    shipped_orders: WeeklySection
    paid_orders: WeeklySection


def parse_numeric(value: Any) -> float:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text or text.startswith("#"):
        return 0.0
    text = text.replace("\xa0", "").replace(" ", "")
    text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return 0.0


def parse_date(value: Any) -> date | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    if not text or text.upper() == "N/A":
        return None
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d.%m.%y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def load_taz(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=0)
    df.columns = [col.strip() if isinstance(col, str) else col for col in df.columns]
    invoice_col = get_invoice_col(df)
    if invoice_col != COL_INVOICE:
        df = df.rename(columns={invoice_col: COL_INVOICE})
    df = df[df[COL_STATUS].notna()]
    df = df[~((df[COL_STATUS] == "1 NOT PAID") & (df[COL_COMMENT] == "SAMPLE"))]
    if COL_INVOICER in df.columns:
        invoicer = df[COL_INVOICER].astype(str).str.strip()
        df = df[~invoicer.isin(EXCLUDED_INVOICERS)]
    return df


def get_invoice_col(df: pd.DataFrame) -> str:
    for col in (COL_INVOICE, "счет", "№ счета"):
        if col in df.columns:
            return col
    raise KeyError(f"Invoice column not found. Columns: {list(df.columns)[:10]}")


def normalize_invoice(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip()
    if text.endswith(".0") and text[:-2].isdigit():
        return text[:-2]
    return text


def add_row_key(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["_key"] = (
        out[COL_INVOICE].map(normalize_invoice)
        + "|"
        + out[COL_PN].astype(str).str.strip()
        + "|"
        + out[COL_DESC].astype(str).str.strip().str.slice(0, 60)
    )
    return out


def report_snapshot(df: pd.DataFrame) -> pd.DataFrame:
    parts = [filter_in_work(df), filter_shipped(df)]
    if not parts:
        return pd.DataFrame()
    snapshot = pd.concat(parts, ignore_index=True)
    snapshot = add_row_key(snapshot)
    return snapshot.drop_duplicates(subset="_key")


def report_snapshot_keys(df: pd.DataFrame) -> set[str]:
    snap = report_snapshot(df)
    if snap.empty:
        return set()
    return set(snap["_key"])


def date_in_range(value: Any, start: date, end: date) -> bool:
    parsed = parse_date(value)
    return parsed is not None and start <= parsed <= end


def row_sale(row: pd.Series) -> float:
    return parse_numeric(row.get(COL_SALE))


def row_total_paid(row: pd.Series) -> float:
    return parse_numeric(row.get(COL_PAY1)) + parse_numeric(row.get(COL_PAY2))


def verify_payment(row: pd.Series) -> tuple[bool, str]:
    balance = parse_numeric(row.get(COL_BALANCE))
    pay_as = parse_numeric(row.get(COL_PAY1))
    pay_au = parse_numeric(row.get(COL_PAY2))
    date_at = parse_date(row.get(COL_PAY2_DATE))

    if balance != 0:
        return False, "остаток AV ≠ 0"
    if pay_as <= 0 and pay_au <= 0:
        return False, "нет оплат AS/AU"
    if pay_au > 0 and date_at is None:
        return False, "есть AU, нет даты AT"
    return True, "OK"


def build_weekly_summary(
    current_df: pd.DataFrame,
    previous_df: pd.DataFrame,
    week_start: date,
    week_end: date,
) -> WeeklySummary:
    current = add_row_key(current_df)
    previous = add_row_key(previous_df)
    prev_snap_keys = report_snapshot_keys(previous_df)
    curr_snap_keys = report_snapshot_keys(current_df)
    prev_by_key = previous.set_index("_key", drop=False)

    new_rows: list[dict[str, Any]] = []
    for _, row in current.iterrows():
        order_date = parse_date(row.get(COL_ORDER_DATE))
        if order_date is None or not (week_start <= order_date <= week_end):
            continue
        if row[COL_STATUS] not in STATUSES_IN_WORK:
            continue
        new_rows.append(
            {
                "№ счета": row[COL_INVOICE],
                "P/N": row[COL_PN],
                "DESCRIPTION": row[COL_DESC],
                "Status": row[COL_STATUS],
                "Дата": order_date,
                "Сумма, USD": row_sale(row),
                "Примечание": row.get(COL_COMMENT, ""),
            }
        )

    shipped_rows: list[dict[str, Any]] = []
    for _, row in current.iterrows():
        if row[COL_STATUS] not in STATUSES_SHIPPED:
            continue
        delivery_date = parse_date(row.get(COL_DELIVERY_ACTUAL))
        delivery_in_week = delivery_date is not None and week_start <= delivery_date <= week_end

        prev_status = None
        if row["_key"] in prev_by_key.index:
            prev_row = prev_by_key.loc[row["_key"]]
            if isinstance(prev_row, pd.DataFrame):
                prev_row = prev_row.iloc[0]
            prev_status = prev_row[COL_STATUS]

        status_became_shipped = (
            prev_status is not None
            and prev_status not in STATUSES_SHIPPED
            and row[COL_STATUS] in STATUSES_SHIPPED
        )
        if not (status_became_shipped or delivery_in_week):
            continue
        note = "статус → отгружено" if status_became_shipped else "дата отгрузки в периоде"
        if status_became_shipped and delivery_in_week:
            note = "статус и дата отгрузки в периоде"
        shipped_rows.append(
            {
                "№ счета": row[COL_INVOICE],
                "P/N": row[COL_PN],
                "DESCRIPTION": row[COL_DESC],
                "Status": row[COL_STATUS],
                "Дата": delivery_date or parse_date(row.get(COL_ORDER_DATE)),
                "Сумма, USD": row_sale(row),
                "Примечание": note,
            }
        )

    paid_rows: list[dict[str, Any]] = []
    disappeared_keys = prev_snap_keys - curr_snap_keys
    current_by_key = current.set_index("_key", drop=False)
    curr_in_work_keys = set(add_row_key(filter_in_work(current_df))["_key"])

    for key in sorted(disappeared_keys):
        if key in curr_in_work_keys:
            continue
        if key in current_by_key.index:
            row = current_by_key.loc[key]
            if isinstance(row, pd.DataFrame):
                row = row.iloc[0]
        else:
            prev_match = previous[previous["_key"] == key]
            if prev_match.empty:
                continue
            row = prev_match.iloc[0]

        verified, note = verify_payment(row)
        if parse_numeric(row.get(COL_BALANCE)) != 0:
            continue
        paid_rows.append(
            {
                "№ счета": row[COL_INVOICE],
                "P/N": row[COL_PN],
                "DESCRIPTION": row[COL_DESC],
                "Status": row[COL_STATUS],
                "Сумма, USD": row_sale(row),
                "Оплата AS": parse_numeric(row.get(COL_PAY1)),
                "Дата AT": parse_date(row.get(COL_PAY2_DATE)),
                "Оплата AU": parse_numeric(row.get(COL_PAY2)),
                "Остаток AV": parse_numeric(row.get(COL_BALANCE)),
                "Проверка": note if verified else f"Проверить: {note}",
            }
        )

    return WeeklySummary(
        week_start=week_start,
        week_end=week_end,
        new_orders=WeeklySection(
            title="Новые заказы",
            count=len(new_rows),
            total=sum(r["Сумма, USD"] for r in new_rows),
            rows=new_rows,
            headers=WEEKLY_HEADERS,
        ),
        shipped_orders=WeeklySection(
            title="Отгруженные заказы",
            count=len(shipped_rows),
            total=sum(r["Сумма, USD"] for r in shipped_rows),
            rows=shipped_rows,
            headers=WEEKLY_HEADERS,
        ),
        paid_orders=WeeklySection(
            title="Оплаченные клиентом",
            count=len(paid_rows),
            total=sum(r["Сумма, USD"] for r in paid_rows),
            rows=paid_rows,
            headers=WEEKLY_PAID_HEADERS,
        ),
    )


def write_weekly_summary_sheet(ws, client: str, summary: WeeklySummary) -> None:
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 28
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 12
    ws.column_dimensions["F"].width = 14
    ws.column_dimensions["G"].width = 24
    for col in ("H", "I", "J"):
        ws.column_dimensions[col].width = 14

    ws.merge_cells("A1:J1")
    title = ws["A1"]
    title.value = f"Сводка за неделю — {client}"
    _style_cell(title, font=FONT_TITLE, alignment=ALIGN_LEFT)

    ws.merge_cells("A2:J2")
    subtitle = ws["A2"]
    subtitle.value = (
        f"Период: {summary.week_start.strftime('%d.%m.%Y')} — "
        f"{summary.week_end.strftime('%d.%m.%Y')}"
    )
    _style_cell(subtitle, font=FONT_SUBTITLE, alignment=ALIGN_LEFT)

    row = 4
    for section in (summary.new_orders, summary.shipped_orders, summary.paid_orders):
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=10)
        section_cell = ws.cell(row, 1)
        section_cell.value = section.title
        fill = FILL_SECTION_WORK if section.title == "Новые заказы" else (
            FILL_SECTION_SHIPPED if section.title == "Отгруженные заказы" else PatternFill("solid", fgColor=COLOR_LIGHT_ORANGE)
        )
        _style_cell(section_cell, font=FONT_SECTION, fill=fill, alignment=ALIGN_LEFT)
        row += 1

        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        totals = ws.cell(row, 1)
        totals.value = f"Количество: {section.count}   |   Сумма: {section.total:,.2f} USD"
        _style_cell(totals, font=FONT_BODY_BOLD, fill=FILL_KPI, alignment=ALIGN_LEFT, border=BORDER_THIN)
        row += 1

        headers = section.headers
        for col_idx, header in enumerate(headers, start=1):
            cell = ws.cell(row, col_idx)
            cell.value = header
            _style_cell(cell, font=FONT_HEADER, fill=FILL_HEADER, alignment=ALIGN_CENTER, border=BORDER_THIN)
        row += 1

        if not section.rows:
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=len(headers))
            empty = ws.cell(row, 1)
            empty.value = "Нет данных за период"
            _style_cell(empty, font=FONT_BODY, alignment=ALIGN_LEFT, border=BORDER_THIN)
            row += 2
            continue

        for item in section.rows:
            for col_idx, header in enumerate(headers, start=1):
                cell = ws.cell(row, col_idx)
                value = item.get(header)
                cell.value = value
                number_format = None
                if header in {"Сумма, USD", "Оплата AS", "Оплата AU", "Остаток AV"}:
                    number_format = NUM_FMT
                    align = ALIGN_RIGHT
                elif header == "Дата" or header == "Дата AT":
                    number_format = DATE_FMT if value else None
                    align = ALIGN_CENTER
                else:
                    align = ALIGN_LEFT
                fill = None
                if header == "Проверка" and str(value) != "OK":
                    fill = FILL_ALERT
                _style_cell(cell, font=FONT_BODY, fill=fill, alignment=align, border=BORDER_THIN, number_format=number_format)
            row += 1
        row += 2

    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=10)
    note = ws.cell(row, 1)
    note.value = (
        "Оплаченные позиции: исчезли из отчёта и проверены по колонкам "
        "AS (первая оплата), AT (дата завершающего платежа), AU (завершающий платёж), AV (остаток)."
    )
    _style_cell(note, font=FONT_SUBTITLE, alignment=ALIGN_LEFT)


def filter_client(df: pd.DataFrame, client: str) -> pd.DataFrame:
    return df[df[COL_CUSTOMER] == client].copy()


def filter_in_work(df: pd.DataFrame) -> pd.DataFrame:
    return df[df[COL_STATUS].isin(STATUSES_IN_WORK)].copy()


def filter_shipped(df: pd.DataFrame) -> pd.DataFrame:
    shipped = df[df[COL_STATUS].isin(STATUSES_SHIPPED)].copy()
    shipped["_balance"] = shipped[COL_BALANCE].map(parse_numeric)
    return shipped[shipped["_balance"] != 0].drop(columns="_balance")


def aggregate(df: pd.DataFrame) -> dict[str, float]:
    purchase = df[COL_PURCHASE].map(parse_numeric).sum()
    transport = df[COL_TRANSPORT].map(parse_numeric).sum()
    fee = df[COL_FEE].map(parse_numeric).sum()
    sale = df[COL_SALE].map(parse_numeric).sum()
    paid = df[COL_PAY1].map(parse_numeric).sum() + df[COL_PAY2].map(parse_numeric).sum()
    due = df[COL_BALANCE].map(parse_numeric).sum()
    margin = sale - purchase - transport - fee
    return {
        "Закупка": purchase,
        "Транспорт": transport,
        "FEE": fee,
        "Продажа": sale,
        "Маржа": margin,
        "Оплачено клиентом": paid,
        "К оплате клиентом": due,
    }


def shipped_balance_over_30_days(df: pd.DataFrame, report_date: date) -> float:
    total = 0.0
    for _, row in df.iterrows():
        delivery = parse_date(row.get(COL_DELIVERY_ACTUAL))
        if delivery is None:
            continue
        days = (report_date - delivery).days
        if days > 30:
            total += parse_numeric(row.get(COL_BALANCE))
    return total


def prepare_output_frame(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=list(COLUMN_RENAME.values()))
    out = df.rename(columns=COLUMN_RENAME)
    sort_by: list[str] = []
    if "ЗАКАЗ ВЗЯТ В РАБОТУ" in out.columns:
        out["_sort_order_date"] = pd.to_datetime(out["ЗАКАЗ ВЗЯТ В РАБОТУ"], errors="coerce")
        sort_by.append("_sort_order_date")
    if "№ счета" in out.columns:
        out["_sort_invoice"] = out["№ счета"].astype(str)
        sort_by.append("_sort_invoice")
    if sort_by:
        out = out.sort_values(sort_by, na_position="last")
        out = out.drop(columns=sort_by)
    return out.reset_index(drop=True)


def build_subtotal_row(totals: dict[str, float], columns: list[str]) -> dict[str, Any]:
    row: dict[str, Any] = {col: None for col in columns}
    row["Status"] = "ИТОГО"
    row["Комментарий"] = "Сводка по разделу"
    row["Chnl"] = ""
    row["Номер группы"] = ""
    if "UNIQUE\nUNIT\nCODE" in row:
        row["UNIQUE\nUNIT\nCODE"] = "-"
    value_map = {
        "Закупка, итого": totals["Закупка"],
        "Продажная, итого (без НДС)": totals["Продажа"],
        "Стоимость доставки ПЛАН, за весь счет! ": totals["Транспорт"],
        "Transaction fee": totals["FEE"],
        "Оплачено\n\nПервая группа платежей": totals["Оплачено клиентом"],
        "Остаток к оплате": totals["К оплате клиентом"],
    }
    for col, value in value_map.items():
        if col in row:
            row[col] = value
    return row


def _style_cell(cell, font=None, fill=None, alignment=None, border=None, number_format=None) -> None:
    if font is not None:
        cell.font = font
    if fill is not None:
        cell.fill = fill
    if alignment is not None:
        cell.alignment = alignment
    if border is not None:
        cell.border = border
    if number_format is not None:
        cell.number_format = number_format


def _apply_range_border(ws, min_row, max_row, min_col, max_col) -> None:
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            cell.border = BORDER_THIN


def _set_column_widths(ws, columns: list[str]) -> None:
    for idx, name in enumerate(columns, start=1):
        letter = get_column_letter(idx)
        ws.column_dimensions[letter].width = COLUMN_WIDTHS.get(name, 12)


def _format_detail_sheet(ws, columns: list[str], data_rows: int) -> None:
    last_col = len(columns)
    col_index = {name: idx for idx, name in enumerate(columns, start=1)}

    # Header row
    for col in range(1, last_col + 1):
        cell = ws.cell(1, col)
        _style_cell(cell, font=FONT_HEADER, fill=FILL_HEADER, alignment=ALIGN_CENTER, border=BORDER_THIN)

    # Subtotal row
    for col in range(1, last_col + 1):
        cell = ws.cell(2, col)
        _style_cell(cell, font=FONT_BODY_BOLD, fill=FILL_SUBTOTAL, alignment=ALIGN_LEFT, border=BORDER_THIN)
        header = columns[col - 1]
        if header in CURRENCY_COLUMNS and isinstance(cell.value, (int, float)):
            cell.number_format = NUM_FMT

    # Data rows
    for row_idx in range(3, data_rows + 2):
        zebra = PatternFill("solid", fgColor=COLOR_ZEBRA) if row_idx % 2 == 1 else None
        status = ws.cell(row_idx, col_index.get("Status", 0)).value if "Status" in col_index else None
        status_fill = STATUS_FILLS.get(str(status))

        for col in range(1, last_col + 1):
            cell = ws.cell(row_idx, col)
            header = columns[col - 1]
            fill = status_fill if header == "Status" and status_fill else zebra
            align = ALIGN_RIGHT if header in CURRENCY_COLUMNS or header == "QTY" else ALIGN_LEFT
            number_format = None
            if header in CURRENCY_COLUMNS:
                number_format = NUM_FMT
            elif header in DATE_COLUMNS and isinstance(cell.value, (datetime, date)):
                number_format = DATE_FMT
            _style_cell(
                cell,
                font=FONT_BODY,
                fill=fill,
                alignment=align,
                border=BORDER_THIN,
                number_format=number_format,
            )

    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A1:{get_column_letter(last_col)}{data_rows + 1}"
    ws.row_dimensions[1].height = 36
    ws.row_dimensions[2].height = 22
    _set_column_widths(ws, columns)


def write_detail_sheet(ws, df: pd.DataFrame, totals: dict[str, float]) -> None:
    columns = list(df.columns)
    subtotal = build_subtotal_row(totals, columns)
    ws.append(columns)
    ws.append([subtotal.get(col) for col in columns])
    for row in dataframe_to_rows(df, index=False, header=False):
        ws.append(row)
    _format_detail_sheet(ws, columns, len(df))


def write_total_sheet(
    ws,
    client: str,
    report_date: date,
    in_work_totals: dict[str, float],
    shipped_totals: dict[str, float],
    shipped_over_30: float,
    in_work_count: int,
    shipped_count: int,
) -> None:
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 22
    for col in range(2, 9):
        ws.column_dimensions[get_column_letter(col)].width = 16

    # Title block
    ws.merge_cells("A1:H1")
    title = ws["A1"]
    title.value = f"Отчёт по заказам — {client}"
    _style_cell(title, font=FONT_TITLE, alignment=Alignment(horizontal="left", vertical="center"))

    ws.merge_cells("A2:H2")
    subtitle = ws["A2"]
    subtitle.value = (
        f"Дата отчёта: {report_date.strftime('%d.%m.%Y')}   |   "
        f"В работе: {in_work_count} поз.   |   Отгружено: {shipped_count} поз."
    )
    _style_cell(subtitle, font=FONT_SUBTITLE, alignment=ALIGN_LEFT)

    # KPI cards
    kpi_row = 4
    kpis = [
        ("Продажа (в работе)", in_work_totals["Продажа"], COLOR_LIGHT_BLUE),
        ("К оплате (в работе)", in_work_totals["К оплате клиентом"], COLOR_LIGHT_ORANGE),
        ("Продажа (отгружено)", shipped_totals["Продажа"], COLOR_LIGHT_GREEN),
        ("К оплате (отгружено)", shipped_totals["К оплате клиентом"], COLOR_LIGHT_ORANGE),
    ]
    for i, (label, value, color) in enumerate(kpis):
        col = 1 + i * 2
        label_cell = ws.cell(kpi_row, col)
        value_cell = ws.cell(kpi_row + 1, col)
        ws.merge_cells(start_row=kpi_row, start_column=col, end_row=kpi_row, end_column=col + 1)
        ws.merge_cells(start_row=kpi_row + 1, start_column=col, end_row=kpi_row + 1, end_column=col + 1)
        label_cell.value = label
        value_cell.value = value
        fill = PatternFill("solid", fgColor=color)
        _style_cell(label_cell, font=FONT_BODY_BOLD, fill=fill, alignment=ALIGN_CENTER, border=BORDER_THIN)
        _style_cell(value_cell, font=FONT_KPI, fill=fill, alignment=ALIGN_CENTER, border=BORDER_THIN, number_format=NUM_FMT)

    def write_block(start_row: int, title: str, totals: dict[str, float], section_fill: PatternFill) -> int:
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=8)
        section_cell = ws.cell(start_row, 1)
        section_cell.value = title
        _style_cell(section_cell, font=FONT_SECTION, fill=section_fill, alignment=ALIGN_LEFT)

        header_row = start_row + 1
        ws.cell(header_row, 1).value = "Раздел"
        for idx, header in enumerate(TOTAL_HEADERS, start=2):
            cell = ws.cell(header_row, idx)
            cell.value = header
            _style_cell(cell, font=FONT_HEADER, fill=FILL_HEADER, alignment=ALIGN_CENTER, border=BORDER_THIN)
        _style_cell(ws.cell(header_row, 1), font=FONT_HEADER, fill=FILL_HEADER, alignment=ALIGN_CENTER, border=BORDER_THIN)

        data_row = header_row + 1
        ws.cell(data_row, 1).value = title
        for idx, header in enumerate(TOTAL_HEADERS, start=2):
            cell = ws.cell(data_row, idx)
            cell.value = totals[header]
            fill = FILL_ALERT if header == "К оплате клиентом" else None
            _style_cell(
                cell,
                font=FONT_BODY_BOLD,
                fill=fill,
                alignment=ALIGN_RIGHT,
                border=BORDER_THIN,
                number_format=NUM_FMT,
            )
        _style_cell(ws.cell(data_row, 1), font=FONT_BODY_BOLD, alignment=ALIGN_LEFT, border=BORDER_THIN)
        _apply_range_border(ws, header_row, data_row, 1, 8)
        return data_row + 2

    row = write_block(8, "В работе", in_work_totals, FILL_SECTION_WORK)
    row = write_block(row, "Отгружено", shipped_totals, FILL_SECTION_SHIPPED)

    note_row = row
    ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=7)
    note = ws.cell(note_row, 1)
    note.value = "Из них отгружено более 30 дней назад (остаток к оплате)"
    _style_cell(note, font=FONT_BODY_BOLD, fill=FILL_ALERT, alignment=ALIGN_RIGHT, border=BORDER_THIN)

    alert_cell = ws.cell(note_row, 8)
    alert_cell.value = shipped_over_30
    _style_cell(alert_cell, font=FONT_KPI, fill=FILL_ALERT, alignment=ALIGN_RIGHT, border=BORDER_THIN, number_format=NUM_FMT)


def load_report_snapshot_as_previous(path: Path) -> pd.DataFrame:
    frames: list[pd.DataFrame] = []
    for sheet in ("В работе", "Отгружено"):
        df = pd.read_excel(path, sheet_name=sheet)
        status_col = "Status" if "Status" in df.columns else COL_STATUS
        df = df[df[status_col] != "ИТОГО"]
        rename_map: dict[str, str] = {status_col: COL_STATUS}
        for src, dst in (
            ("№ счета", COL_INVOICE),
            ("счет", COL_INVOICE),
            ("P/N", COL_PN),
            ("Комментарий", COL_COMMENT),
            ("Продажная, итого (без НДС)", COL_SALE),
            ("Остаток к оплате", COL_BALANCE),
            ("ЗАКАЗ ВЗЯТ В РАБОТУ", COL_ORDER_DATE),
            ("ФАКТИЧЕСКАЯ ДАТА ПОСТАВКИ (СОГЛАСНО УСЛОВИЯМ ПОСТАВКИ)", COL_DELIVERY_ACTUAL),
        ):
            if src in df.columns:
                rename_map[src] = dst
        frames.append(df.rename(columns=rename_map))
    if not frames:
        return pd.DataFrame()
    merged = pd.concat(frames, ignore_index=True)
    if COL_INVOICE not in merged.columns:
        raise KeyError(f"Previous report {path} does not contain invoice column")
    return merged


def generate_report(
    input_path: Path,
    output_path: Path,
    client: str = "Utair",
    report_date: date | None = None,
    previous_input: Path | None = None,
    previous_report: Path | None = None,
    previous_date: date | None = None,
    week_days: int = 7,
) -> Path:
    report_date = report_date or date.today()
    df = load_taz(input_path)
    client_df = filter_client(df, client)

    in_work_raw = filter_in_work(client_df)
    shipped_raw = filter_shipped(client_df)
    in_work_df = prepare_output_frame(in_work_raw)
    shipped_df = prepare_output_frame(shipped_raw)

    in_work_totals = aggregate(in_work_raw)
    shipped_totals = aggregate(shipped_raw)
    shipped_over_30 = shipped_balance_over_30_days(shipped_raw, report_date)

    wb = Workbook()
    total_ws = wb.active
    total_ws.title = "Total"
    total_ws.sheet_properties.tabColor = COLOR_NAVY
    write_total_sheet(
        total_ws,
        client,
        report_date,
        in_work_totals,
        shipped_totals,
        shipped_over_30,
        len(in_work_df),
        len(shipped_df),
    )

    if previous_input or previous_report:
        if previous_input:
            previous_df = filter_client(load_taz(previous_input), client)
        else:
            previous_df = filter_client(load_report_snapshot_as_previous(previous_report), client)
        week_end = report_date
        week_start = previous_date or (report_date - timedelta(days=week_days))
        summary = build_weekly_summary(client_df, previous_df, week_start, week_end)
        weekly_ws = wb.create_sheet("Сводка за неделю")
        weekly_ws.sheet_properties.tabColor = "7030A0"
        write_weekly_summary_sheet(weekly_ws, client, summary)

    shipped_ws = wb.create_sheet("Отгружено")
    shipped_ws.sheet_properties.tabColor = COLOR_GREEN
    write_detail_sheet(shipped_ws, shipped_df, shipped_totals)

    in_work_ws = wb.create_sheet("В работе")
    in_work_ws.sheet_properties.tabColor = COLOR_ORANGE
    write_detail_sheet(in_work_ws, in_work_df, in_work_totals)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path


def parse_report_date(value: str) -> date:
    for fmt in ("%d.%m.%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            continue
    raise argparse.ArgumentTypeError(f"Invalid date: {value}. Use DD.MM.YYYY or YYYY-MM-DD.")


def default_output_name(client: str, report_date: date) -> str:
    safe_client = re.sub(r'[\\/:*?"<>|]', "-", client.strip())
    return f"{safe_client} статус {report_date.strftime('%d.%m.%Y')}.xlsx"


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate client status Excel report from TAZ export.")
    parser.add_argument("--input", "-i", type=Path, required=True, help="Path to TAZ .xlsx file")
    parser.add_argument("--output", "-o", type=Path, help="Output .xlsx path")
    parser.add_argument("--client", "-c", default="Utair", help="Customer name (default: Utair)")
    parser.add_argument(
        "--date",
        "-d",
        type=parse_report_date,
        default=date.today(),
        help="Report date for filename and >30 days calc (default: today)",
    )
    parser.add_argument(
        "--previous-input",
        type=Path,
        help="Previous TAZ export for weekly summary comparison",
    )
    parser.add_argument(
        "--previous-report",
        type=Path,
        help="Previous generated report .xlsx (fallback if previous TAZ is unavailable)",
    )
    parser.add_argument(
        "--previous-date",
        type=parse_report_date,
        help="Start date of weekly summary period (default: report date minus 7 days)",
    )
    parser.add_argument(
        "--week-days",
        type=int,
        default=7,
        help="Length of summary period in days if --previous-date is not set (default: 7)",
    )
    args = parser.parse_args()

    output = args.output or Path("reports") / default_output_name(args.client, args.date)
    result = generate_report(
        args.input,
        output,
        client=args.client,
        report_date=args.date,
        previous_input=args.previous_input,
        previous_report=args.previous_report,
        previous_date=args.previous_date,
        week_days=args.week_days,
    )
    print(f"Report saved: {result}")
    print(f"Client: {args.client}")
    print(f"В работе rows: {len(filter_in_work(filter_client(load_taz(args.input), args.client)))}")
    print(f"Отгружено rows: {len(filter_shipped(filter_client(load_taz(args.input), args.client)))}")
    if args.previous_input or args.previous_report:
        print("Weekly summary sheet: included")


if __name__ == "__main__":
    main()
