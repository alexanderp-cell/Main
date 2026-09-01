#!/usr/bin/env python3
"""Generate Utair TUZ performance HTML report (group sheets only)."""

from __future__ import annotations

import html
import json
import re
import statistics
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

import openpyxl

TUZ_PATH = Path("/tmp/utair_analysis/TUZ_ТУЗ полный файл 01.09.2026.xlsx")
TAZ_PATH = Path("/tmp/utair_analysis/TAZ_ТАЗ полный файл 28.08.2026.xlsx")
OUTPUT_PATH = Path("/workspace/UTAIR_TUZ_performance_report.html")

EARLY_STATUSES = {
    "0. Начальный этап",
    "1. Проценка у поставщиков",
    "8. Backup",
    "11. Внесено в ТАЗ",
}

GROUP_SHEETS = [
    "Группа A",
    "Группа B",
    "Группа C",
    "Группа 2 (old)",
    "Группа 3 (old)",
    "Группа 5 (old)",
]
OLD_SHEETS = {"Группа 2 (old)", "Группа 3 (old)", "Группа 5 (old)"}

PRICE_BUCKETS = [
    ("0–10 тыс. $", 0, 10_000),
    ("11–25 тыс. $", 10_001, 25_000),
    ("26–50 тыс. $", 25_001, 50_000),
    ("51–75 тыс. $", 50_001, 75_000),
    ("76–100 тыс. $", 75_001, 100_000),
    ("101–150 тыс. $", 100_001, 150_000),
    ("151–200 тыс. $", 150_001, 200_000),
    ("200+ тыс. $", 200_001, 9_999_999),
]

NOT_FOUND_STATUSES = {
    "1. Не нашли",
    "2. Не будет проквотировано",
    "1. Вне компетенции",
    "1. Пропуск",
}


def is_utair(customer) -> bool:
    if not customer:
        return False
    text = str(customer).lower()
    return "utair" in text or text == "utg" or "компонентс (utg)" in text


def parse_num(value) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if text.startswith("="):
        return None
    text = text.upper()
    if text in {"", "TBA", "N/A", "NA", "#VALUE!", "-", "БЕЗ ЦЕНЫ"}:
        return None
    match = re.search(r"[-+]?\d[\d\s.,]*", text)
    if not match:
        return None
    token = match.group().replace(" ", "").replace(",", "")
    try:
        return float(token)
    except ValueError:
        return None


def to_dt(value) -> datetime | None:
    return value if isinstance(value, datetime) else None


def hours_between(start: datetime | None, end: datetime | None) -> float | None:
    if not start or not end:
        return None
    delta = (end - start).total_seconds() / 3600
    if delta < 0:
        return None
    return delta


def fmt_hours(value: float | None) -> str:
    if value is None:
        return "—"
    if value < 24:
        return f"{value:.0f} ч"
    return f"{value / 24:.1f} д"


def fmt_money(value: float | None) -> str:
    if value is None:
        return "—"
    return f"${value:,.0f}"


def fmt_pct(value: float | None) -> str:
    if value is None:
        return "—"
    return f"{value:.0f}%"


def bucket_for_sale(value: float | None) -> str | None:
    if value is None:
        return None
    for label, low, high in PRICE_BUCKETS:
        if low <= value <= high:
            return label
    return None


def sheet_columns(sheet: str) -> dict[str, int]:
    if sheet in OLD_SHEETS:
        return {
            "status": 1,
            "request_dt": 2,
            "request_no": 3,
            "client": 5,
            "pn": 7,
            "alt_pn": 8,
            "description": 9,
            "qty": 10,
            "quote_dt": 12,
            "root_price": 14,
            "supplier_price": 16,
            "transit": 18,
            "cond": 19,
            "offered": 25,
            "sent_dt": 26,
            "accepted_dt": 27,
            "invoice": 28,
        }
    return {
        "status": 1,
        "request_dt": 2,
        "request_no": 3,
        "client": 5,
        "pn": 7,
        "alt_pn": 8,
        "description": 9,
        "qty": 10,
        "quote_dt": 15,
        "root_price": 17,
        "supplier_price": 19,
        "transit": 21,
        "cond": 22,
        "offered": 28,
        "sent_dt": 29,
        "accepted_dt": 30,
        "invoice": 31,
    }


@dataclass
class OfferRow:
    sheet: str
    status: str | None
    request_dt: datetime | None
    quote_dt: datetime | None
    sent_dt: datetime | None
    accepted_dt: datetime | None
    root_price: float | None
    supplier_price: float | None
    transit: float | None
    offered: float | None
    cond: str | None
    invoice: str | None
    pn_bold: bool = False
    alt_bold: bool = False


@dataclass
class RequestLine:
    sheet: str
    request_no: str
    pn: str
    alt_pn: str | None
    description: str | None
    qty: float | None
    offers: list[OfferRow] = field(default_factory=list)

    @property
    def key(self) -> tuple[str, str, str]:
        return (self.sheet, self.request_no, self.pn)

    def primary_status(self) -> str:
        priority = [
            "7. Клиент согласовал",
            "5. Есть интерес",
            "4. Квотация направлена клиенту",
            "3. Цена на деталь получена",
            "1. Проценка у поставщиков",
            "6. Клиент отказал",
            "8. Backup",
        ]
        statuses = {str(o.status) for o in self.offers if o.status}
        for label in priority:
            if label in statuses:
                return label
        return next(iter(statuses), "—")

    def market_rows(self) -> list[OfferRow]:
        """Supplier quotes on meaningful rows (exclude Backup / not-found statuses)."""
        rows = []
        for offer in self.offers:
            if offer.status in NOT_FOUND_STATUSES or offer.status == "8. Backup":
                continue
            if offer.supplier_price is not None:
                rows.append(offer)
        return rows

    def has_supplier_price(self) -> bool:
        return len(self.market_rows()) > 0

    def supplier_without_offer(self) -> bool:
        """Supplier quote exists, status past procurement, but no client Offered × Qty."""
        if not self.market_rows() or self.sale_value() is not None:
            return False
        return self.primary_status() not in EARLY_STATUSES

    def has_root_price(self) -> bool:
        return any(o.root_price is not None for o in self.offers)

    def is_not_found(self) -> bool:
        if self.found_on_market():
            return False
        statuses = {o.status for o in self.offers if o.status}
        return bool(statuses & NOT_FOUND_STATUSES) or self.primary_status() in NOT_FOUND_STATUSES

    def is_in_procurement(self) -> bool:
        return self.primary_status() == "1. Проценка у поставщиков" and not self.found_on_market()

    def selected_offer(self) -> OfferRow | None:
        sent = [o for o in self.offers if o.sent_dt and o.offered is not None]
        if sent:
            return sorted(sent, key=lambda o: o.sent_dt)[0]
        priced = [o for o in self.offers if o.offered is not None and o.status not in {"8. Backup"}]
        if priced:
            return sorted(priced, key=lambda o: o.offered)[0]
        priced = [o for o in self.offers if o.offered is not None]
        return priced[0] if priced else None

    def sale_value(self) -> float | None:
        offer = self.selected_offer()
        if not offer or offer.offered is None or not self.qty:
            return None
        return offer.offered * self.qty

    def markup_pct(self) -> float | None:
        offer = self.selected_offer()
        if not offer or offer.offered is None:
            return None
        cost = (offer.supplier_price or 0) + (offer.transit or 0)
        if cost <= 0:
            cost = offer.root_price
        if not cost or cost <= 0:
            return None
        return (offer.offered - cost) / cost * 100

    def won(self) -> bool:
        if any(o.status == "7. Клиент согласовал" for o in self.offers):
            return True
        return any(o.invoice for o in self.offers)

    def found_on_market(self) -> bool:
        """Found = есть Supplier Price на строке, которая не Backup и не «не нашли»."""
        return len(self.market_rows()) > 0

    def order_value(self) -> float | None:
        if not self.won():
            return None
        sale = self.sale_value()
        if sale is not None:
            return sale
        qty = self.qty or 1
        for offer in self.offers:
            if offer.status == "7. Клиент согласовал" and offer.offered is not None:
                return offer.offered * qty
        for offer in self.offers:
            if offer.offered is not None:
                return offer.offered * qty
        for offer in self.market_rows():
            return offer.supplier_price * qty
        return None

    def timing(self) -> dict[str, float | None]:
        request_dt = next((o.request_dt for o in self.offers if o.request_dt), None)
        quote_dts = [o.quote_dt for o in self.offers if o.quote_dt]
        sent_dts = [o.sent_dt for o in self.offers if o.sent_dt]
        quote_dt = min(quote_dts) if quote_dts else None
        sent_dt = min(sent_dts) if sent_dts else None
        return {
            "procurement": hours_between(request_dt, quote_dt),
            "sales": hours_between(quote_dt, sent_dt),
            "total": hours_between(request_dt, sent_dt),
        }


def load_request_lines(path: Path) -> list[RequestLine]:
    wb_vals = openpyxl.load_workbook(path, read_only=False, data_only=True)
    wb_fmt = openpyxl.load_workbook(path, read_only=False, data_only=False)
    grouped: dict[tuple[str, str, str], RequestLine] = {}

    for sheet in GROUP_SHEETS:
        if sheet not in wb_vals.sheetnames:
            continue
        ws_vals = wb_vals[sheet]
        ws_fmt = wb_fmt[sheet]
        cols = sheet_columns(sheet)

        for row_vals, row_fmt in zip(ws_vals.iter_rows(min_row=2), ws_fmt.iter_rows(min_row=2)):
            values = [cell.value for cell in row_vals]
            if not values:
                continue

            def val(key: str):
                index = cols[key] - 1
                return values[index] if index < len(values) else None

            client = val("client")
            if not is_utair(client):
                continue

            request_no = val("request_no")
            pn = val("pn")
            if request_no in (None, "") or pn in (None, ""):
                continue

            request_no_s = str(request_no).strip()
            pn_s = str(pn).strip()
            key = (sheet, request_no_s, pn_s)
            if key not in grouped:
                grouped[key] = RequestLine(
                    sheet=sheet,
                    request_no=request_no_s,
                    pn=pn_s,
                    alt_pn=str(val("alt_pn")).strip() if val("alt_pn") not in (None, "") else None,
                    description=str(val("description")).strip() if val("description") else None,
                    qty=parse_num(val("qty")),
                    offers=[],
                )

            pn_cell = row_fmt[cols["pn"] - 1]
            alt_cell = row_fmt[cols["alt_pn"] - 1]
            invoice_raw = val("invoice")
            invoice = None
            if invoice_raw not in (None, ""):
                invoice_text = str(invoice_raw).strip()
                if len(invoice_text) <= 20 and re.match(r"^[\d-]+$", invoice_text.replace(".0", "")):
                    invoice = invoice_text.replace(".0", "")

            grouped[key].offers.append(
                OfferRow(
                    sheet=sheet,
                    status=str(val("status")).strip() if val("status") else None,
                    request_dt=to_dt(val("request_dt")),
                    quote_dt=to_dt(val("quote_dt")),
                    sent_dt=to_dt(val("sent_dt")),
                    accepted_dt=to_dt(val("accepted_dt")),
                    root_price=parse_num(val("root_price")),
                    supplier_price=parse_num(val("supplier_price")),
                    transit=parse_num(val("transit")),
                    offered=parse_num(val("offered")),
                    cond=str(val("cond")).strip() if val("cond") else None,
                    invoice=invoice,
                    pn_bold=bool(pn_cell.font and pn_cell.font.bold),
                    alt_bold=bool(alt_cell.font and alt_cell.font.bold),
                )
            )

    wb_vals.close()
    wb_fmt.close()
    return list(grouped.values())


def load_taz_orders_2026(path: Path) -> dict[str, Any]:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["ORDERS"] if "ORDERS" in wb.sheetnames else wb[wb.sheetnames[0]]
    total = 0.0
    lines = 0
    invoices: set[str] = set()

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row:
            continue
        customer = row[6]
        if not is_utair(customer):
            continue
        work_dt = row[16]
        if not isinstance(work_dt, datetime) or work_dt.year != 2026:
            continue
        qty = parse_num(row[13]) or 1
        sale_ea = parse_num(row[32])
        sale_total = parse_num(row[33])
        if sale_total is not None:
            total += sale_total
        elif sale_ea is not None:
            total += sale_ea * qty
        lines += 1
        if row[0] not in (None, ""):
            invoices.add(str(row[0]).strip())

    wb.close()
    return {
        "total": total,
        "lines": lines,
        "invoices": len(invoices),
    }


def median(values: list[float]) -> float | None:
    return statistics.median(values) if values else None


def aggregate(lines: list[RequestLine]) -> dict[str, Any]:
    by_bucket: dict[str, list[RequestLine]] = {label: [] for label, _, _ in PRICE_BUCKETS}
    unpriced: list[RequestLine] = []

    for line in lines:
        sale = line.sale_value()
        bucket = bucket_for_sale(sale)
        if bucket:
            by_bucket[bucket].append(line)
        else:
            unpriced.append(line)

    def summarize(group: list[RequestLine]) -> dict[str, Any]:
        proc = [t["procurement"] for line in group for t in [line.timing()] if t["procurement"] is not None]
        sales = [t["sales"] for line in group for t in [line.timing()] if t["sales"] is not None]
        total = [t["total"] for line in group for t in [line.timing()] if t["total"] is not None]
        found = sum(1 for line in group if line.found_on_market())
        not_found = sum(1 for line in group if line.is_not_found())
        in_procurement = sum(1 for line in group if line.is_in_procurement())
        supplier_no_offer = sum(1 for line in group if line.supplier_without_offer())
        won = sum(1 for line in group if line.won())
        tuz_orders_total = sum(line.order_value() or 0 for line in group if line.won())
        sent = sum(1 for line in group if any(o.sent_dt for o in line.offers))
        pending_proc = in_procurement
        offer_counts = [len(line.market_rows()) for line in group if line.market_rows()]
        markups = [line.markup_pct() for line in group if line.markup_pct() is not None]

        return {
            "count": len(group),
            "found": found,
            "found_pct": (found / len(group) * 100) if group else 0,
            "not_found": not_found,
            "in_procurement": in_procurement,
            "supplier_no_offer": supplier_no_offer,
            "won": won,
            "won_pct": (won / len(group) * 100) if group else 0,
            "tuz_orders_total": tuz_orders_total,
            "sent": sent,
            "pending_proc": pending_proc,
            "avg_offers": statistics.mean(offer_counts) if offer_counts else 0,
            "median_markup": median(markups),
            "median_proc": median(proc),
            "median_sales": median(sales),
            "median_total": median(total),
            "lines": group,
        }

    taz_orders = load_taz_orders_2026(TAZ_PATH) if TAZ_PATH.exists() else None

    overall = summarize(lines)
    if taz_orders:
        overall["orders_total"] = taz_orders["total"]
        overall["taz_lines_2026"] = taz_orders["lines"]
        overall["taz_invoices_2026"] = taz_orders["invoices"]
    else:
        overall["orders_total"] = overall["tuz_orders_total"]
        overall["taz_lines_2026"] = None
        overall["taz_invoices_2026"] = None
    buckets = {label: summarize(by_bucket[label]) for label, _, _ in PRICE_BUCKETS}
    buckets["Без продажной оценки"] = summarize(unpriced)

    return {
        "generated_at": datetime.now().strftime("%d.%m.%Y %H:%M"),
        "source": TUZ_PATH.name,
        "taz_source": TAZ_PATH.name if TAZ_PATH.exists() else None,
        "overall": overall,
        "buckets": buckets,
    }


def render_table_rows(lines: list[RequestLine]) -> str:
    rows = sorted(lines, key=lambda line: line.sale_value() or 0, reverse=True)
    html_rows = []
    for line in rows[:200]:
        offer = line.selected_offer()
        best_supplier = min((o.supplier_price for o in line.market_rows()), default=None)
        display_offer = offer.offered if offer else None
        timing = line.timing()
        alt_flag = ""
        if any(o.alt_bold or o.pn_bold for o in line.offers):
            alt_flag = " <span class='tag alt'>alt P/N</span>"
        html_rows.append(
            "<tr>"
            f"<td>{html.escape(line.request_no)}</td>"
            f"<td><b>{html.escape(line.pn)}</b>{alt_flag}"
            f"{('<div class=\"muted\">ALT: ' + html.escape(line.alt_pn) + '</div>') if line.alt_pn else ''}</td>"
            f"<td>{html.escape(line.description or '—')}</td>"
            f"<td>{line.qty or '—'}</td>"
            f"<td>{html.escape(line.primary_status())}</td>"
            f"<td>{len(line.market_rows())}</td>"
            f"<td>{fmt_money(display_offer)}"
            f"{('<div class=\"muted\">Supp: ' + fmt_money(best_supplier) + '</div>') if (display_offer is None and best_supplier is not None) else ''}</td>"
            f"<td>{fmt_money(line.sale_value())}</td>"
            f"<td>{fmt_pct(line.markup_pct())}</td>"
            f"<td>{fmt_hours(timing['procurement'])}</td>"
            f"<td>{fmt_hours(timing['sales'])}</td>"
            f"<td>{fmt_hours(timing['total'])}</td>"
            f"<td>{'✓' if line.won() else '—'}</td>"
            "</tr>"
        )
    if len(rows) > 200:
        html_rows.append(
            f"<tr><td colspan='13' class='muted'>… ещё {len(rows)-200} строк</td></tr>"
        )
    return "\n".join(html_rows)


def render_html(data: dict[str, Any]) -> str:
    overall = data["overall"]
    bucket_order = [label for label, _, _ in PRICE_BUCKETS] + ["Без продажной оценки"]

    bucket_sections = []
    for label in bucket_order:
        bucket = data["buckets"][label]
        if bucket["count"] == 0:
            continue
        extra = ""
        if label == "Без продажной оценки" and bucket.get("supplier_no_offer"):
            extra = f" · supplier price без Offered: {bucket['supplier_no_offer']}"
        bucket_sections.append(
            f"""
            <details class="bucket" {'open' if label == bucket_order[0] else ''}>
              <summary>
                <span class="bucket-title">{html.escape(label)}</span>
                <span class="bucket-meta">{bucket['count']} запросов · найдено {bucket['found_pct']:.0f}% · не найдено {bucket['not_found']} · заказов {bucket['won']} · медиана B→AC {fmt_hours(bucket['median_total'])}{extra}</span>
              </summary>
              <div class="bucket-body">
                <div class="mini-grid">
                  <div class="mini"><div class="k">Найдено (Supplier Price)</div><div class="v">{bucket['found']} <span class="sub">({bucket['found_pct']:.0f}%)</span></div></div>
                  <div class="mini warn"><div class="k">Не найдено</div><div class="v">{bucket['not_found']}</div></div>
                  <div class="mini warn"><div class="k">В проценке</div><div class="v">{bucket['in_procurement']}</div></div>
                  <div class="mini"><div class="k">Отправлено клиенту</div><div class="v">{bucket['sent']}</div></div>
                  <div class="mini"><div class="k">Заказ получен</div><div class="v">{bucket['won']} <span class="sub">({bucket['won_pct']:.0f}%)</span></div></div>
                  <div class="mini accent"><div class="k">Сумма заказов (ТУЗ)</div><div class="v">{fmt_money(bucket['tuz_orders_total'])}</div></div>
                  <div class="mini"><div class="k">Ср. офферов / запрос</div><div class="v">{bucket['avg_offers']:.1f}</div></div>
                  <div class="mini"><div class="k">B→O / O→AC / B→AC</div><div class="v">{fmt_hours(bucket['median_proc'])} / {fmt_hours(bucket['median_sales'])} / {fmt_hours(bucket['median_total'])}</div></div>
                </div>
                <div class="table-wrap">
                  <table>
                    <thead>
                      <tr>
                        <th>Request №</th><th>P/N</th><th>Description</th><th>Qty</th><th>Статус</th>
                        <th>Офферов</th><th>Offered/ea</th><th>Sale $</th><th>Наценка</th>
                        <th>B→O</th><th>O→AC</th><th>B→AC</th><th>Заказ</th>
                      </tr>
                    </thead>
                    <tbody>
                      {render_table_rows(bucket['lines'])}
                    </tbody>
                  </table>
                </div>
              </div>
            </details>
            """
        )

    return f"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>UTair — оценка ТУЗ</title>
<link rel="preconnect" href="https://fonts.googleapis.com"/>
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin/>
<link href="https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&family=Fraunces:opsz,wght@9..144,600;9..144,700&display=swap" rel="stylesheet"/>
<style>
:root {{
  --ink:#1a2332; --ink-soft:#3d4a5c; --muted:#6b7a8f; --line:rgba(26,35,50,.10);
  --paper:#f3f6f8; --panel:rgba(255,255,255,.84); --teal:#0b6e6b; --teal-deep:#084e4c;
  --amber:#c9782a; --danger:#b42318; --ok:#0f7a4c; --shadow:0 16px 40px rgba(26,35,50,.08); --radius:18px;
}}
* {{ box-sizing:border-box; }}
body {{
  margin:0; color:var(--ink); font-family:"Manrope",system-ui,sans-serif;
  background:radial-gradient(1000px 500px at 0% 0%, rgba(11,110,107,.12), transparent 55%), var(--paper);
}}
.wrap {{ max-width:1180px; margin:0 auto; padding:36px 18px 72px; }}
h1 {{ font-family:"Fraunces",Georgia,serif; font-size:clamp(2rem,4vw,3rem); color:var(--teal-deep); margin:0 0 8px; }}
.lead {{ color:var(--ink-soft); max-width:46rem; line-height:1.55; margin:0 0 18px; }}
.meta {{ color:var(--muted); font-size:.88rem; margin-bottom:24px; }}
.grid {{ display:grid; grid-template-columns:repeat(4,1fr); gap:12px; margin:18px 0 28px; }}
@media(max-width:980px){{ .grid {{ grid-template-columns:repeat(2,1fr); }} }}
@media(max-width:560px){{ .grid {{ grid-template-columns:1fr; }} }}
.kpi {{ background:var(--panel); border:1px solid var(--line); border-radius:16px; padding:16px; box-shadow:var(--shadow); }}
.kpi .k {{ font-size:.72rem; text-transform:uppercase; letter-spacing:.05em; color:var(--muted); font-weight:700; }}
.kpi .v {{ font-family:"Fraunces",Georgia,serif; font-size:1.55rem; margin-top:4px; color:var(--teal-deep); }}
.kpi .h {{ font-size:.84rem; color:var(--ink-soft); margin-top:4px; }}
.kpi.accent {{ background:linear-gradient(135deg,#d7efed,#fff); }}
.kpi.warn {{ background:linear-gradient(135deg,#f8ead8,#fff); }}
.panel {{ background:var(--panel); border:1px solid var(--line); border-radius:var(--radius); padding:18px; box-shadow:var(--shadow); margin:18px 0; }}
.panel h2 {{ font-family:"Fraunces",Georgia,serif; margin:0 0 12px; font-size:1.35rem; }}
.note {{ font-size:.9rem; color:var(--ink-soft); border-left:3px solid var(--teal); padding:10px 12px; background:rgba(255,255,255,.65); }}
details.bucket {{ background:var(--panel); border:1px solid var(--line); border-radius:16px; box-shadow:var(--shadow); margin:14px 0; overflow:hidden; }}
details.bucket > summary {{
  list-style:none; cursor:pointer; padding:16px 18px; display:flex; flex-wrap:wrap; gap:8px 16px;
  align-items:center; justify-content:space-between; background:linear-gradient(180deg,#fff,rgba(255,255,255,.7));
}}
details.bucket > summary::-webkit-details-marker {{ display:none; }}
details.bucket[open] > summary {{ border-bottom:1px solid var(--line); }}
.bucket-title {{ font-family:"Fraunces",Georgia,serif; font-size:1.15rem; font-weight:700; color:var(--teal-deep); }}
.bucket-meta {{ color:var(--muted); font-size:.86rem; }}
.bucket-body {{ padding:14px 16px 18px; }}
.mini-grid {{ display:grid; grid-template-columns:repeat(4,1fr); gap:10px; margin-bottom:14px; }}
@media(max-width:860px){{ .mini-grid {{ grid-template-columns:repeat(2,1fr); }} }}
.mini {{ background:#fff; border:1px solid var(--line); border-radius:12px; padding:10px 12px; }}
.mini.warn {{ background:#fff7ed; }}
.mini .k {{ font-size:.68rem; text-transform:uppercase; letter-spacing:.04em; color:var(--muted); font-weight:700; }}
.mini .v {{ font-size:1rem; font-weight:700; margin-top:3px; color:var(--ink); }}
.mini .sub {{ font-size:.82rem; color:var(--muted); font-weight:600; }}
.table-wrap {{ overflow:auto; border:1px solid var(--line); border-radius:12px; }}
table {{ width:100%; border-collapse:collapse; font-size:.82rem; background:#fff; }}
th,td {{ padding:8px 7px; border-bottom:1px solid var(--line); text-align:left; vertical-align:top; }}
th {{ font-size:.68rem; text-transform:uppercase; letter-spacing:.04em; color:var(--muted); background:#f8fafb; position:sticky; top:0; }}
tr:hover td {{ background:#fafcfd; }}
.muted {{ color:var(--muted); font-size:.82rem; }}
.tag {{ display:inline-block; font-size:.68rem; font-weight:700; padding:2px 6px; border-radius:999px; }}
.tag.alt {{ background:#fde68a; color:#7c5a00; }}
</style>
</head>
<body>
<div class="wrap">
  <h1>UTair — оценка работы по ТУЗ</h1>
  <p class="lead">Анализ проработки запросов Ютэйр по групповым листам ТУЗ: скорость закупок и продаж, качество проценки, наценка и конверсия в заказы. Категории — по продажной стоимости (Offered × Qty).</p>
  <div class="meta">Источник: {html.escape(data['source'])} · сформировано {html.escape(data['generated_at'])} · Questions v2 не включён</div>

  <div class="grid">
    <div class="kpi accent"><div class="k">Запросов (строк RFQ)</div><div class="v">{overall['count']}</div><div class="h">Request № + P/N</div></div>
    <div class="kpi"><div class="k">Найдено на рынке</div><div class="v">{overall['found']}</div><div class="h">{overall['found_pct']:.0f}% — только Supplier Price</div></div>
    <div class="kpi warn"><div class="k">Не найдено / в проценке</div><div class="v">{overall['not_found']} / {overall['in_procurement']}</div><div class="h">без supplier quote / ещё ищем</div></div>
    <div class="kpi"><div class="k">Отправлено клиенту</div><div class="v">{overall['sent']}</div><div class="h">есть дата в AC</div></div>
    <div class="kpi accent"><div class="k">Заказов получено</div><div class="v">{overall['won']}</div><div class="h">{overall['won_pct']:.1f}% конверсия</div></div>
    <div class="kpi accent"><div class="k">Сумма заказов 2026</div><div class="v">{fmt_money(overall['orders_total'])}</div><div class="h">ТАЗ ORDERS · {overall.get('taz_invoices_2026') or '—'} счетов · {overall.get('taz_lines_2026') or '—'} строк</div></div>
    <div class="kpi"><div class="k">ТУЗ выигранные (Offered×Qty)</div><div class="v">{fmt_money(overall['tuz_orders_total'])}</div><div class="h">{overall['won']} RFQ со статусом «согласовал» / invoice</div></div>
    <div class="kpi"><div class="k">Медиана B→O</div><div class="v">{fmt_hours(overall['median_proc'])}</div><div class="h">закупки: запрос → цена</div></div>
    <div class="kpi"><div class="k">Медиана O→AC / B→AC</div><div class="v">{fmt_hours(overall['median_sales'])} / {fmt_hours(overall['median_total'])}</div><div class="h">продажи / полный цикл</div></div>
  </div>

  <div class="panel">
    <h2>Как читать отчёт</h2>
    <div class="note">
      <b>Найдено на рынке</b> — есть <b>Supplier Price per unit</b> на строке, которая не Backup и не «не нашли». Root Price <u>не считается</u> рыночным оффером.<br/>
      <b>Без продажной оценки</b> — нет Offered × Qty; из них <b>{overall.get('supplier_no_offer', 0)}</b> — статус ≥ «Цена получена», есть Supplier Price, но Offered ещё пустой (без строк «Проценка» / Backup).<br/>
      <b>B→O</b> — от внесения запроса (B) до получения цены с рынка (O). · <b>O→AC</b> — от цены до отправки оффера (AC).<br/>
      <b>Сумма заказов 2026</b> — продажная итого (col AH) по листу ORDERS в ТАЗ, Ютэйр, дата взятия в работу 2026. ТУЗ-оценка — Offered × Qty по выигранным RFQ.<br/>
      <b>alt P/N</b> — P/N выделен жирным в ТУЗ (часто предложен альтернативный номер).
    </div>
  </div>

  <div class="panel">
    <h2>Категории по продажной стоимости</h2>
    {''.join(bucket_sections)}
  </div>
</div>
</body>
</html>"""


def main():
    lines = load_request_lines(TUZ_PATH)
    data = aggregate(lines)
    OUTPUT_PATH.write_text(render_html(data), encoding="utf-8")
    print(f"Generated {OUTPUT_PATH}")
    print(json.dumps({k: data['overall'][k] for k in ['count','found','won','pending_proc']}, ensure_ascii=False))


if __name__ == "__main__":
    main()
