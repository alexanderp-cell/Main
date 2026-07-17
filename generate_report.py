#!/usr/bin/env python3
"""Generate a client status Excel report from a TAZ export."""

from __future__ import annotations

import argparse
import re
from datetime import date, datetime
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
COL_PAY2 = "Оплачено клиентом, USD (cntr+shift+V)\nЗавершающий платеж"
COL_BALANCE = "Остаток к оплате клиентом, USD"
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
    df = df[df[COL_STATUS].notna()]
    df = df[~((df[COL_STATUS] == "1 NOT PAID") & (df[COL_COMMENT] == "SAMPLE"))]
    if COL_INVOICER in df.columns:
        invoicer = df[COL_INVOICER].astype(str).str.strip()
        df = df[~invoicer.isin(EXCLUDED_INVOICERS)]
    return df


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
    sort_cols = []
    if "ЗАКАЗ ВЗЯТ В РАБОТУ" in out.columns:
        sort_cols.append("ЗАКАЗ ВЗЯТ В РАБОТУ")
    if "№ счета" in out.columns:
        sort_cols.append("№ счета")
    if sort_cols:
        out = out.sort_values(sort_cols, na_position="last")
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


def generate_report(
    input_path: Path,
    output_path: Path,
    client: str = "Utair",
    report_date: date | None = None,
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
    args = parser.parse_args()

    output = args.output or Path("reports") / default_output_name(args.client, args.date)
    result = generate_report(args.input, output, client=args.client, report_date=args.date)
    print(f"Report saved: {result}")
    print(f"Client: {args.client}")
    print(f"В работе rows: {len(filter_in_work(filter_client(load_taz(args.input), args.client)))}")
    print(f"Отгружено rows: {len(filter_shipped(filter_client(load_taz(args.input), args.client)))}")


if __name__ == "__main__":
    main()
