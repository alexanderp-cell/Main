#!/usr/bin/env python3
"""Generate monthly HTML reports for realized (FINISHED + fully paid) TAZ orders."""

from __future__ import annotations

import argparse
import calendar
import hashlib
import json
import re
import urllib.request
import xml.etree.ElementTree as ET
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path
from statistics import mean
from typing import Any

import pandas as pd

# --- TAZ columns ---
COL_INVOICE = "Номер счета"
COL_STATUS = "Status"
COL_CUSTOMER = "Customer"
COL_PN = "p/n"
COL_DESC = "DESCRIPTION"
COL_SUPPLIER = "Поставщик"
COL_PAY_TYPE = "Тип оплаты"
COL_ORDER_DATE = "ЗАКАЗ ВЗЯТ В РАБОТУ (ДАТА) ОТ КЛИЕНТА"
COL_DELIVERY = "ФАКТИЧЕСКАЯ ДАТА ПОСТАВКИ (СОГЛАСНО УСЛОВИЯМ ПОСТАВКИ)"
COL_PURCHASE = "Закупка, итого"
COL_SALE = "Продажная, итого"
COL_TRANSPORT_PLAN = (
    "Стоимость доставки ПЛАН, за весь счет! Если в счете несколько строк, "
    'то "размазываем" равномерно планируюмую стоиомость транспорта на все позиции из счета.'
)
COL_TRANSPORT_FACT = "Стоимость доставки факт"
COL_FEE = "Transaction fee"
COL_DUTY = "Пошлина,сбор (ФАКТ) USD."
COL_SVH = "СВХ стоимость (ФАКТ)"
COL_BALANCE = "Остаток к оплате клиентом, USD"
COL_PAY1_DATE = "Дата оплаты Клиентом\n\nПервая группа платежей"
COL_PAY1 = "Оплачено клиентом, USD (cntr+shift+V)\nПервая группа платежей"
COL_PAY2_DATE = "Дата оплаты Клиентом\n\nЗавершающий платеж"
COL_PAY2 = "Оплачено клиентом, USD (cntr+shift+V)\nЗавершающий платеж"
COL_SUPPLIER_PAY_DATE = "Дата оплаты поставщику"
COL_SUPPLIER_PAY = "Сумма оплаты поставщику"

EXCLUDED_STATUSES = {"5 CANCELLED", "7 REFUND", "8 WARRANTY", "CANCELLED", "REFUND", "WARRANTY"}
COMPLEX_SETTLEMENT_SUPPLIERS = ("LUFTHANSA", "BLUE OCEAN")
CLIENT_PAY_NORM_DAYS = 14

CBR_CACHE_PATH = Path("/tmp/cbr_usd_rub_cache.json")


def parse_numeric(value: Any) -> float:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return 0.0
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    text = str(value).strip().replace("\xa0", "").replace(" ", "").replace(",", ".")
    if not text or text.startswith("#"):
        return 0.0
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
    if not text or text.lower() in {"nan", "none", "nat"}:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(text[:19], fmt).date()
        except ValueError:
            continue
    try:
        ts = pd.to_datetime(value, dayfirst=True, errors="coerce")
        if pd.isna(ts):
            return None
        return ts.date()
    except Exception:
        return None


def load_taz(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=0)
    df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]
    # transport plan column may have trailing space variants
    rename = {}
    for c in df.columns:
        if isinstance(c, str) and c.startswith("Стоимость доставки ПЛАН"):
            rename[c] = COL_TRANSPORT_PLAN
        if isinstance(c, str) and c.startswith("Пошлина,сбор"):
            rename[c] = COL_DUTY
    if rename:
        df = df.rename(columns=rename)
    return df


def is_complex_supplier(name: Any) -> bool:
    s = str(name or "").upper()
    return any(x in s for x in COMPLEX_SETTLEMENT_SUPPLIERS)


def normalize_pay_type(value: Any) -> str:
    s = str(value or "").strip().lower()
    if "взаимозачет" in s:
        return "netting"
    if "частичн" in s:
        return "partial"
    if "предоплат" in s:
        return "prepay"
    if "постоплат" in s:
        return "postpay"
    return "unknown"


def completion_payment_date(pay_type: str, pay1: date | None, pay2: date | None) -> date | None:
    """Date when the order is considered fully paid by the client."""
    if pay_type == "prepay":
        return pay1 or pay2
    if pay_type == "partial":
        return pay2 or pay1
    if pay_type == "postpay":
        return pay2 or pay1
    # unknown / netting: prefer final payment, else first
    return pay2 or pay1


def client_paid_usd(pay1_amt: float, pay2_amt: float) -> float:
    return pay1_amt + pay2_amt


@dataclass
class RealizedRow:
    invoice: str
    customer: str
    pn: str
    description: str
    supplier: str
    pay_type_raw: str
    pay_type: str
    order_date: date | None
    delivery_date: date | None
    pay1_date: date | None
    pay2_date: date | None
    supplier_pay_date: date | None
    paid_complete_date: date | None
    report_event_date: date | None
    sale: float
    purchase: float
    transport_plan: float
    transport_fact: float
    fee: float
    duty: float
    svh: float
    client_paid: float
    supplier_paid: float
    margin_plan: float
    margin_fact: float
    margin_plan_pct: float | None
    margin_fact_pct: float | None
    stage1_days: int | None = None
    stage2_days: int | None = None
    total_lead_days: int | None = None
    client_pay_days: int | None = None
    fx_rub: float | None = None
    fx_rate_client: float | None = None
    fx_rate_supplier: float | None = None
    flags: list[str] = field(default_factory=list)
    missing_delivery: bool = False


class CbrUsdRub:
    """CBR USD/RUB daily rates with local cache. Rate = RUB per 1 USD."""

    def __init__(self, cache_path: Path = CBR_CACHE_PATH) -> None:
        self.cache_path = cache_path
        self.cache: dict[str, float] = {}
        if cache_path.exists():
            try:
                self.cache = json.loads(cache_path.read_text(encoding="utf-8"))
            except Exception:
                self.cache = {}

    def save(self) -> None:
        self.cache_path.write_text(json.dumps(self.cache, ensure_ascii=False, indent=0), encoding="utf-8")

    def rate_on(self, d: date) -> float | None:
        # Walk back up to 10 days for weekends/holidays
        for i in range(0, 11):
            day = d - timedelta(days=i)
            key = day.isoformat()
            if key in self.cache:
                return self.cache[key]
            rate = self._fetch(day)
            if rate is not None:
                self.cache[key] = rate
                # also store requested date mapping to found rate for speed
                if i > 0:
                    self.cache[d.isoformat()] = rate
                return rate
        return None

    def _fetch(self, d: date) -> float | None:
        url = f"https://www.cbr.ru/scripts/XML_daily.asp?date_req={d.strftime('%d/%m/%Y')}"
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=30) as resp:
                raw = resp.read()
            root = ET.fromstring(raw)
            for valute in root.findall("Valute"):
                if (valute.findtext("CharCode") or "") == "USD":
                    nominal = float((valute.findtext("Nominal") or "1").replace(",", "."))
                    value = float((valute.findtext("Value") or "0").replace(",", "."))
                    return value / nominal
        except Exception:
            return None
        return None


def build_realized_rows(df: pd.DataFrame, cbr: CbrUsdRub) -> tuple[list[RealizedRow], list[RealizedRow]]:
    rows: list[RealizedRow] = []
    missing_delivery_rows: list[RealizedRow] = []

    for _, raw in df.iterrows():
        status = str(raw.get(COL_STATUS) or "").strip()
        if status in EXCLUDED_STATUSES:
            continue
        if status != "4 FINISHED":
            continue

        balance = parse_numeric(raw.get(COL_BALANCE))
        if abs(balance) > 0.009:
            continue

        sale = parse_numeric(raw.get(COL_SALE))
        purchase = parse_numeric(raw.get(COL_PURCHASE))
        transport_plan = parse_numeric(raw.get(COL_TRANSPORT_PLAN))
        transport_fact = parse_numeric(raw.get(COL_TRANSPORT_FACT))
        fee = parse_numeric(raw.get(COL_FEE))
        duty = parse_numeric(raw.get(COL_DUTY))
        svh = parse_numeric(raw.get(COL_SVH))

        margin_plan = sale - purchase - transport_plan - fee
        margin_fact = sale - purchase - transport_fact - fee - duty - svh
        margin_plan_pct = (margin_plan / sale * 100.0) if sale else None
        margin_fact_pct = (margin_fact / sale * 100.0) if sale else None

        pay_type_raw = str(raw.get(COL_PAY_TYPE) or "").strip()
        pay_type = normalize_pay_type(pay_type_raw)
        pay1_date = parse_date(raw.get(COL_PAY1_DATE))
        pay2_date = parse_date(raw.get(COL_PAY2_DATE))
        pay1_amt = parse_numeric(raw.get(COL_PAY1))
        pay2_amt = parse_numeric(raw.get(COL_PAY2))
        paid_complete = completion_payment_date(pay_type, pay1_date, pay2_date)
        delivery = parse_date(raw.get(COL_DELIVERY))
        order_date = parse_date(raw.get(COL_ORDER_DATE))
        supplier_pay_date = parse_date(raw.get(COL_SUPPLIER_PAY_DATE))
        supplier_paid = parse_numeric(raw.get(COL_SUPPLIER_PAY))
        client_paid = client_paid_usd(pay1_amt, pay2_amt)
        if client_paid <= 0 and sale > 0:
            client_paid = sale  # fallback for display/FX base if amounts empty but balance 0

        supplier = str(raw.get(COL_SUPPLIER) or "").strip()
        flags: list[str] = []

        row = RealizedRow(
            invoice=str(raw.get(COL_INVOICE) or "").strip(),
            customer=str(raw.get(COL_CUSTOMER) or "").strip() or "(без клиента)",
            pn=str(raw.get(COL_PN) or "").strip(),
            description=str(raw.get(COL_DESC) or "").strip(),
            supplier=supplier,
            pay_type_raw=pay_type_raw or "не указан",
            pay_type=pay_type,
            order_date=order_date,
            delivery_date=delivery,
            pay1_date=pay1_date,
            pay2_date=pay2_date,
            supplier_pay_date=supplier_pay_date,
            paid_complete_date=paid_complete,
            report_event_date=None,
            sale=sale,
            purchase=purchase,
            transport_plan=transport_plan,
            transport_fact=transport_fact,
            fee=fee,
            duty=duty,
            svh=svh,
            client_paid=client_paid,
            supplier_paid=supplier_paid,
            margin_plan=margin_plan,
            margin_fact=margin_fact,
            margin_plan_pct=margin_plan_pct,
            margin_fact_pct=margin_fact_pct,
            flags=flags,
        )

        if delivery is None:
            row.missing_delivery = True
            row.flags.append("нет факт-даты отгрузки")
            missing_delivery_rows.append(row)
            continue

        if paid_complete is None:
            row.flags.append("нет даты полной оплаты клиента")
            missing_delivery_rows.append(row)
            continue

        row.report_event_date = max(delivery, paid_complete)

        # Lead times
        complex_sup = is_complex_supplier(supplier)
        if order_date and delivery:
            row.total_lead_days = (delivery - order_date).days

        if complex_sup:
            row.flags.append("сложный поставщик (Lufthansa/Blue Ocean) — этапы 1/2 не делим")
            if order_date and delivery:
                row.stage1_days = None
                row.stage2_days = None
        else:
            if order_date and supplier_pay_date:
                row.stage1_days = (supplier_pay_date - order_date).days
            if supplier_pay_date and delivery:
                row.stage2_days = (delivery - supplier_pay_date).days

            if row.stage1_days is not None and row.stage1_days < 0:
                row.flags.append("аномалия этапа 1: оплата поставщику раньше взятия в работу")
            if row.stage2_days is not None and row.stage2_days < 0:
                row.flags.append("аномалия этапа 2: похоже на постоплату поставщику (отгрузка раньше оплаты)")
            if row.stage1_days is not None and row.stage1_days > 365:
                row.flags.append("аномалия этапа 1: >365 дней")
            if row.stage2_days is not None and row.stage2_days > 365:
                row.flags.append("аномалия этапа 2: >365 дней")
            if supplier_pay_date is None:
                row.flags.append("нет даты оплаты поставщику — этапы 1/2 н/д")

        # Client payment speed after delivery (final payment)
        final_client_pay = pay2_date or pay1_date
        if delivery and final_client_pay and final_client_pay >= delivery:
            row.client_pay_days = (final_client_pay - delivery).days
        elif delivery and final_client_pay and final_client_pay < delivery:
            # prepaid before delivery — not measured as "after shipment delay"
            row.client_pay_days = None
            if pay_type in {"postpay", "partial"}:
                row.flags.append("итоговый платёж клиента раньше отгрузки")

        # FX for classic postpay: client pays after supplier pay
        if (
            paid_complete
            and supplier_pay_date
            and paid_complete > supplier_pay_date
            and client_paid > 0
        ):
            k_client = cbr.rate_on(paid_complete)
            k_supplier = cbr.rate_on(supplier_pay_date)
            row.fx_rate_client = k_client
            row.fx_rate_supplier = k_supplier
            if k_client is not None and k_supplier is not None:
                row.fx_rub = client_paid * (k_client - k_supplier)
                # absurd flags
                if abs(k_client - k_supplier) > 25:
                    row.flags.append("курсы: большая разница (>25 RUB/USD)")
                if row.fx_rub is not None and abs(row.fx_rub) > abs(client_paid) * max(k_client, k_supplier) * 0.35:
                    row.flags.append("курсовая разница выглядит аномально большой")
                if complex_sup and row.fx_rub is not None and abs(row.fx_rub) > 500_000:
                    row.flags.append("сложный поставщик: крупная курсовая сумма — проверьте")

        rows.append(row)

    return rows, missing_delivery_rows


def filter_month(rows: list[RealizedRow], year: int, month: int) -> list[RealizedRow]:
    return [
        r
        for r in rows
        if r.report_event_date is not None
        and r.report_event_date.year == year
        and r.report_event_date.month == month
    ]


def avg(vals: list[float | int | None]) -> float | None:
    clean = [float(v) for v in vals if v is not None]
    return mean(clean) if clean else None


def fmt_money(v: float | None, digits: int = 0) -> str:
    if v is None:
        return "—"
    return f"{v:,.{digits}f}".replace(",", " ")


def fmt_pct(v: float | None) -> str:
    if v is None:
        return "—"
    return f"{v:.1f}%"


def fmt_days(v: float | int | None) -> str:
    if v is None:
        return "—"
    if isinstance(v, float):
        return f"{v:.1f}"
    return str(v)


def client_pay_comment(avg_days: float | None, n: int) -> str:
    if avg_days is None or n == 0:
        return "недостаточно данных по оплате после отгрузки"
    if avg_days <= CLIENT_PAY_NORM_DAYS:
        return f"ок: в среднем {avg_days:.1f} дн. после отгрузки (норма ≤{CLIENT_PAY_NORM_DAYS})"
    if avg_days <= 30:
        return f"задержки: в среднем {avg_days:.1f} дн. после отгрузки (норма ≤{CLIENT_PAY_NORM_DAYS})"
    return f"часто задерживает: в среднем {avg_days:.1f} дн. после отгрузки (норма ≤{CLIENT_PAY_NORM_DAYS})"


def render_html(
    year: int,
    month: int,
    rows: list[RealizedRow],
    missing_delivery: list[RealizedRow],
) -> str:
    month_name = [
        "",
        "январь",
        "февраль",
        "март",
        "апрель",
        "май",
        "июнь",
        "июль",
        "август",
        "сентябрь",
        "октябрь",
        "ноябрь",
        "декабрь",
    ][month]
    last_day = calendar.monthrange(year, month)[1]
    period = f"01.{month:02d}.{year} — {last_day:02d}.{month:02d}.{year}"

    # Attention: missing delivery date, paid in this month (or unpaid-date unknown)
    attention = []
    for r in missing_delivery:
        if "нет факт-даты отгрузки" not in r.flags and "нет даты полной оплаты клиента" not in " ".join(r.flags):
            continue
        # show in month if payment completion falls in month, else if no payment date show all for visibility in first months only via payment
        if r.paid_complete_date and r.paid_complete_date.year == year and r.paid_complete_date.month == month:
            attention.append(r)
        elif r.delivery_date is None and r.paid_complete_date is None:
            # no dates to bucket — skip noisy global dump
            pass
        elif r.delivery_date is None and r.pay2_date and r.pay2_date.year == year and r.pay2_date.month == month:
            attention.append(r)
        elif r.delivery_date is None and r.pay1_date and r.pay1_date.year == year and r.pay1_date.month == month:
            attention.append(r)

    total_sale = sum(r.sale for r in rows)
    total_plan = sum(r.margin_plan for r in rows)
    total_fact = sum(r.margin_fact for r in rows)
    total_fx = sum(r.fx_rub for r in rows if r.fx_rub is not None)
    fx_n = sum(1 for r in rows if r.fx_rub is not None)
    avg_s1 = avg([r.stage1_days for r in rows])
    avg_s2 = avg([r.stage2_days for r in rows])
    avg_lead = avg([r.total_lead_days for r in rows])
    avg_client_pay = avg([r.client_pay_days for r in rows])
    flagged = [r for r in rows if r.flags]

    by_client: dict[str, list[RealizedRow]] = defaultdict(list)
    by_supplier: dict[str, list[RealizedRow]] = defaultdict(list)
    for r in rows:
        by_client[r.customer].append(r)
        supplier_name = r.supplier.strip() if r.supplier and r.supplier.strip() else "(без поставщика)"
        by_supplier[supplier_name].append(r)
    clients_sorted = sorted(by_client.keys(), key=lambda c: (-sum(x.sale for x in by_client[c]), c))
    suppliers_sorted = sorted(by_supplier.keys(), key=lambda s: (-sum(x.sale for x in by_supplier[s]), s))

    def group_block(
        title: str,
        items: list[RealizedRow],
        *,
        mode: str,
    ) -> str:
        c_sale = sum(r.sale for r in items)
        c_plan = sum(r.margin_plan for r in items)
        c_fact = sum(r.margin_fact for r in items)
        c_fx = sum(r.fx_rub for r in items if r.fx_rub is not None)
        c_s1 = avg([r.stage1_days for r in items])
        c_s2 = avg([r.stage2_days for r in items])
        c_pay = avg([r.client_pay_days for r in items])
        c_pay_n = sum(1 for r in items if r.client_pay_days is not None)
        comment = client_pay_comment(c_pay, c_pay_n)
        plan_pct = (c_plan / c_sale * 100) if c_sale else None
        fact_pct = (c_fact / c_sale * 100) if c_sale else None
        delta = c_fact - c_plan

        rows_html = []
        for r in sorted(items, key=lambda x: (x.report_event_date or date.min, x.invoice, x.pn)):
            flag_html = ""
            if r.flags:
                flag_html = "<div class='flags'>" + "<br>".join(f"⚠ {f}" for f in r.flags) + "</div>"
            third_col = r.supplier if mode == "client" else r.customer
            rows_html.append(
                f"""
<tr class="{'flagged' if r.flags else ''}">
  <td>{r.invoice}</td>
  <td>{r.pn}<div class="muted">{r.description[:60]}</div></td>
  <td>{third_col}</td>
  <td>{r.pay_type_raw}</td>
  <td class="num">{fmt_money(r.sale, 0)}</td>
  <td class="num">{fmt_money(r.margin_plan, 0)}<div class="muted">{fmt_pct(r.margin_plan_pct)}</div></td>
  <td class="num">{fmt_money(r.margin_fact, 0)}<div class="muted">{fmt_pct(r.margin_fact_pct)}</div></td>
  <td class="num">{fmt_days(r.stage1_days)}</td>
  <td class="num">{fmt_days(r.stage2_days)}</td>
  <td class="num">{fmt_days(r.client_pay_days)}</td>
  <td class="num">{fmt_money(r.fx_rub, 0)}</td>
  <td>{flag_html or "—"}</td>
</tr>"""
            )

        third_header = "Поставщик" if mode == "client" else "Клиент"
        pay_block = ""
        if mode == "client":
            pay_block = f"""
      <div><div class="label">Оплата после отгрузки</div><div class="value">{fmt_days(c_pay)} дн.</div><div class="muted">{comment}</div></div>"""

        return f"""
<details class="group">
  <summary>
    <span class="group-name">{title}</span>
    <span class="pill">{len(items)} поз.</span>
    <span class="pill">продажа {fmt_money(c_sale, 0)} USD</span>
    <span class="pill">маржа факт {fmt_money(c_fact, 0)} USD ({fmt_pct(fact_pct)})</span>
    <span class="pill">Δ маржи {fmt_money(delta, 0)}</span>
  </summary>
  <div class="group-body">
    <div class="kpi-grid small">
      <div><div class="label">Маржа план</div><div class="value">{fmt_money(c_plan, 0)} USD ({fmt_pct(plan_pct)})</div></div>
      <div><div class="label">Маржа факт</div><div class="value">{fmt_money(c_fact, 0)} USD ({fmt_pct(fact_pct)})</div></div>
      <div><div class="label">Этап 1 / этап 2, дн.</div><div class="value">{fmt_days(c_s1)} / {fmt_days(c_s2)}</div></div>
      {pay_block}
      <div><div class="label">Курсовая разница</div><div class="value">{fmt_money(c_fx, 0)} RUB</div></div>
    </div>
    <table>
      <thead>
        <tr>
          <th>Счёт</th><th>P/N</th><th>{third_header}</th><th>Тип оплаты</th>
          <th>Продажа USD</th><th>Маржа план</th><th>Маржа факт</th>
          <th>Этап 1</th><th>Этап 2</th><th>Оплата клиента</th><th>Курс. разница RUB</th><th>Заметки</th>
        </tr>
      </thead>
      <tbody>
        {''.join(rows_html)}
      </tbody>
    </table>
  </div>
</details>
"""

    attention_html = ""
    if attention:
        att_rows = []
        for r in attention[:200]:
            att_rows.append(
                f"<tr><td>{r.customer}</td><td>{r.invoice}</td><td>{r.pn}</td>"
                f"<td>{r.paid_complete_date.strftime('%d.%m.%Y') if r.paid_complete_date else '—'}</td>"
                f"<td>{', '.join(r.flags) or 'нет факт-даты отгрузки'}</td></tr>"
            )
        attention_html = f"""
<section class="attention">
  <h2>Требуют внимания ({len(attention)})</h2>
  <p>FINISHED + остаток 0, но не хватает факт-даты отгрузки и/или даты полной оплаты — дозаполните в ТАЗ.</p>
  <table>
    <thead><tr><th>Клиент</th><th>Счёт</th><th>P/N</th><th>Оплата</th><th>Комментарий</th></tr></thead>
    <tbody>{''.join(att_rows)}</tbody>
  </table>
</section>
"""

    flags_summary = ""
    if flagged:
        flags_summary = (
            f"<p class='muted'>Строк с пометками/аномалиями: <b>{len(flagged)}</b> "
            f"(см. детали в списках клиентов и поставщиков).</p>"
        )

    client_blocks = "\n".join(group_block(c, by_client[c], mode="client") for c in clients_sorted)
    supplier_blocks = "\n".join(group_block(s, by_supplier[s], mode="supplier") for s in suppliers_sorted)

    return f"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8"/>
<title>Реализованные заказы — {month_name} {year}</title>
<style>
:root {{
  --bg:#e7ecef;
  --card:#ffffff;
  --ink:#202020;
  --muted:#4c4c4c;
  --navy:#022f40;
  --cyan:#d5fbff;
  --orange:#fe621d;
  --line:#c4c4c4;
  --warn:#fe621d;
}}
* {{ box-sizing:border-box; }}
body {{
  margin:0;
  font-family: Arial, Calibri, Helvetica, sans-serif;
  color:var(--ink);
  background: linear-gradient(180deg, #022f40 0 160px, var(--bg) 160px);
  line-height:1.45;
}}
.wrap {{ max-width:1200px; margin:0 auto; padding:28px 20px 64px; }}
.brand {{
  display:flex; align-items:center; justify-content:space-between; gap:12px;
  color:#ffffff; margin-bottom:18px;
}}
.brand-mark {{
  font-size:13px; letter-spacing:.12em; text-transform:uppercase;
  background:var(--cyan); color:var(--navy); padding:6px 10px; font-weight:700;
}}
h1 {{ font-size:30px; margin:0 0 6px; color:#ffffff; font-weight:700; }}
.sub {{ color:#d5fbff; margin-bottom:18px; font-size:14px; }}
h2 {{ font-size:20px; margin:28px 0 12px; color:var(--navy); }}
.card {{
  background:var(--card); border:1px solid var(--line); border-radius:8px;
  padding:18px 18px 8px;
}}
.kpi-grid {{
  display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:12px; margin:12px 0 18px;
}}
.kpi-grid.small {{ grid-template-columns:repeat(3,minmax(0,1fr)); }}
.kpi, .kpi-grid > div {{
  background:#f7fbfc; border:1px solid #d7e2e6; border-radius:8px; padding:12px 14px;
}}
.label {{ font-size:11px; text-transform:uppercase; letter-spacing:.04em; color:#757575; }}
.value {{ font-size:20px; font-weight:700; margin-top:4px; font-variant-numeric:tabular-nums; color:var(--navy); }}
.muted {{ color:#757575; font-size:12px; margin-top:4px; }}
.tabs {{ display:flex; gap:8px; margin:8px 0 14px; }}
.tabs a {{
  text-decoration:none; color:var(--navy); background:#ffffff; border:1px solid var(--line);
  padding:8px 12px; border-radius:6px; font-size:13px; font-weight:700;
}}
.tabs a:hover {{ background:var(--cyan); }}
details.group {{
  background:var(--card); border:1px solid var(--line); border-radius:8px;
  margin:10px 0; overflow:hidden;
}}
details.group summary {{
  cursor:pointer; list-style:none; padding:14px 16px; display:flex; flex-wrap:wrap; gap:8px; align-items:center;
  background:linear-gradient(90deg, #022f40 0%, #022f40 12px, #ffffff 12px);
}}
details.group summary::-webkit-details-marker {{ display:none; }}
.group-name {{ font-weight:700; font-size:16px; margin-right:auto; color:var(--navy); padding-left:8px; }}
.pill {{
  font-size:12px; background:var(--cyan); border:1px solid #b6e8ef; color:var(--navy);
  border-radius:4px; padding:4px 10px; font-weight:700;
}}
.group-body {{ padding:0 14px 16px; border-top:1px solid var(--line); }}
table {{ width:100%; border-collapse:collapse; font-size:13px; }}
th, td {{ border-bottom:1px solid #e8e8e8; padding:8px 6px; vertical-align:top; text-align:left; }}
th {{ font-size:11px; text-transform:uppercase; letter-spacing:.03em; color:#4c4c4c; background:#f3f7f8; }}
td.num {{ text-align:right; font-variant-numeric:tabular-nums; white-space:nowrap; }}
tr.flagged {{ background:#fff4ec; }}
.flags {{ color:var(--orange); font-size:12px; }}
.attention {{
  margin-top:28px; padding:16px; border-radius:8px; border:1px solid #fe621d; background:#fff7f2;
}}
.footer {{ margin-top:28px; color:#4c4c4c; font-size:12px; }}
@media (max-width:900px) {{
  .kpi-grid, .kpi-grid.small {{ grid-template-columns:repeat(2,minmax(0,1fr)); }}
}}
</style>
</head>
<body>
<div class="wrap">
  <div class="brand"><div class="brand-mark">FASTAIR</div><div style="font-size:13px;opacity:.9">realized orders report</div></div>
  <h1>Реализованные заказы — {month_name} {year}</h1>
  <div class="sub">Период {period} · статус FINISHED · остаток клиента = 0 · месяц = max(факт отгрузки, полная оплата)</div>

  <section class="card">
    <h2 style="margin-top:0">Сводка за месяц</h2>
    <div class="kpi-grid">
      <div class="kpi"><div class="label">Позиций</div><div class="value">{len(rows)}</div></div>
      <div class="kpi"><div class="label">Клиентов</div><div class="value">{len(by_client)}</div></div>
      <div class="kpi"><div class="label">Поставщиков</div><div class="value">{len(by_supplier)}</div></div>
      <div class="kpi"><div class="label">Продажи</div><div class="value">{fmt_money(total_sale, 0)}<div class="muted">USD</div></div></div>
      <div class="kpi"><div class="label">Курсовая разница</div><div class="value">{fmt_money(total_fx, 0)}<div class="muted">RUB · {fx_n} сделок</div></div></div>
      <div class="kpi"><div class="label">Маржа план</div><div class="value">{fmt_money(total_plan, 0)}<div class="muted">{fmt_pct((total_plan/total_sale*100) if total_sale else None)} от продаж</div></div></div>
      <div class="kpi"><div class="label">Маржа факт</div><div class="value">{fmt_money(total_fact, 0)}<div class="muted">{fmt_pct((total_fact/total_sale*100) if total_sale else None)} от продаж</div></div></div>
      <div class="kpi"><div class="label">Δ маржи (факт − план)</div><div class="value">{fmt_money(total_fact - total_plan, 0)}<div class="muted">USD</div></div></div>
      <div class="kpi"><div class="label">Сроки, дн.</div><div class="value">{fmt_days(avg_lead)}<div class="muted">этап1 {fmt_days(avg_s1)} · этап2 {fmt_days(avg_s2)} · оплата после отгрузки {fmt_days(avg_client_pay)}</div></div></div>
    </div>
    {flags_summary}
  </section>

  <div class="tabs">
    <a href="#clients">По клиентам</a>
    <a href="#suppliers">По поставщикам</a>
  </div>

  <h2 id="clients">Клиенты</h2>
  {client_blocks or "<p>Нет реализованных позиций за месяц.</p>"}

  <h2 id="suppliers">Поставщики</h2>
  {supplier_blocks or "<p>Нет реализованных позиций за месяц.</p>"}

  {attention_html}

  <div class="footer">
    FASTAIR · источник ТАЗ · курсы USD ЦБ РФ (XML_daily) ·
    этап1: взятие в работу → оплата поставщику; этап2: оплата поставщику → факт отгрузки
    (кроме Lufthansa / Blue Ocean). Курсовая разница только если оплата клиента позже оплаты поставщику:
    Q(оплата клиента USD) × (курс_клиента − курс_поставщика).
  </div>
</div>
</body>
</html>
"""


def generate_month_report(
    all_rows: list[RealizedRow],
    missing: list[RealizedRow],
    year: int,
    month: int,
    output_path: Path,
) -> Path:
    month_rows = filter_month(all_rows, year, month)
    html = render_html(year, month, month_rows, missing)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(html, encoding="utf-8")
    return output_path


def main() -> None:
    parser = argparse.ArgumentParser(description="Monthly realized orders HTML reports from TAZ")
    parser.add_argument("--input", "-i", type=Path, required=True)
    parser.add_argument("--output-dir", "-o", type=Path, default=Path("output/realized"))
    parser.add_argument("--year", type=int, default=2026)
    parser.add_argument("--months", type=str, default="1-6", help="e.g. 1-6 or 1,2,3")
    args = parser.parse_args()

    if "-" in args.months:
        a, b = args.months.split("-", 1)
        months = list(range(int(a), int(b) + 1))
    else:
        months = [int(x) for x in args.months.split(",") if x.strip()]

    print("Loading TAZ...", args.input)
    df = load_taz(args.input)
    cbr = CbrUsdRub()
    print("Building realized rows + fetching CBR rates as needed...")
    rows, missing = build_realized_rows(df, cbr)
    cbr.save()
    print(f"Realized rows with dates: {len(rows)}")
    print(f"Attention (missing delivery / incomplete): {len(missing)}")

    for m in months:
        out = args.output_dir / f"realized_orders_{args.year}-{m:02d}.html"
        generate_month_report(rows, missing, args.year, m, out)
        month_n = len(filter_month(rows, args.year, m))
        print(f"Saved {out} ({month_n} rows)")


if __name__ == "__main__":
    main()
