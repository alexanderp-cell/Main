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
from openpyxl.utils.dataframe import dataframe_to_rows

# Internal TAZ column names
COL_INVOICE = "Номер счета"
COL_STATUS = "Status"
COL_CUSTOMER = "Customer"
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
    row["Status"] = "1 NOT PAID"
    row["Комментарий"] = "SAMPLE"
    row["Chnl"] = "TBA"
    row["Номер группы"] = "TBA"
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


def write_sheet(ws, df: pd.DataFrame, totals: dict[str, float]) -> None:
    columns = list(df.columns)
    subtotal = build_subtotal_row(totals, columns)
    ws.append(columns)
    ws.append([subtotal.get(col) for col in columns])
    for row in dataframe_to_rows(df, index=False, header=False):
        ws.append(row)


def write_total_sheet(
    ws,
    in_work_totals: dict[str, float],
    shipped_totals: dict[str, float],
    shipped_over_30: float,
) -> None:
    ws.append([None, *TOTAL_HEADERS])
    ws.append(["В работе", *[in_work_totals[h] for h in TOTAL_HEADERS]])
    ws.append([None, "Итого"])
    ws.append([None, *[in_work_totals[h] for h in TOTAL_HEADERS]])
    ws.append([])
    ws.append([])
    ws.append([])
    ws.append([None, *TOTAL_HEADERS])
    ws.append(["Отгружено", *[shipped_totals[h] for h in TOTAL_HEADERS]])
    ws.append([None, None, None, None, None, None, None, "Из них отгружено более 30 дней назад"])
    ws.append([None, None, None, None, None, None, None, shipped_over_30])


def generate_report(
    input_path: Path,
    output_path: Path,
    client: str = "Utair",
    report_date: date | None = None,
) -> Path:
    report_date = report_date or date.today()
    df = load_taz(input_path)
    client_df = filter_client(df, client)

    in_work_df = prepare_output_frame(filter_in_work(client_df))
    shipped_df = prepare_output_frame(filter_shipped(client_df))

    in_work_totals = aggregate(filter_in_work(client_df))
    shipped_totals = aggregate(filter_shipped(client_df))
    shipped_over_30 = shipped_balance_over_30_days(filter_shipped(client_df), report_date)

    wb = Workbook()
    total_ws = wb.active
    total_ws.title = "Total"
    write_total_sheet(total_ws, in_work_totals, shipped_totals, shipped_over_30)

    shipped_ws = wb.create_sheet("Отгружено")
    write_sheet(shipped_ws, shipped_df, shipped_totals)

    in_work_ws = wb.create_sheet("В работе")
    write_sheet(in_work_ws, in_work_df, in_work_totals)

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
