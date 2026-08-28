#!/usr/bin/env python3
"""Weekly HTML report: TAZ snapshot diff — TROUBLE, cancellations, warranty."""

from __future__ import annotations

import argparse
import html
import zipfile
from collections import defaultdict
from dataclasses import dataclass, field, replace
from datetime import date, datetime, timedelta
from pathlib import Path
from statistics import mean
from typing import Any

import pandas as pd

# --- TAZ columns (shared with other generators) ---
COL_INVOICE = "Номер счета"
COL_STATUS = "Status"
COL_CUSTOMER = "Customer"
COL_PN = "p/n"
COL_DESC = "DESCRIPTION"
COL_SUPPLIER = "Поставщик"
COL_ROOT_SUPPLIER = "Root supplier"
COL_CATEGORY = "Category"
COL_QTY = "QTY IN PO"
COL_SALE = "Продажная, итого"
COL_PURCHASE = "Закупка, итого"
COL_TRANSPORT_PLAN = (
    "Стоимость доставки ПЛАН, за весь счет! Если в счете несколько строк, "
    'то "размазываем" равномерно планируюмую стоиомость транспорта на все позиции из счета.'
)
COL_FEE = "Transaction fee"
COL_DEADLINE = "КРАЙНЯЯ ДАТА ПОСТАВКИ"
COL_DAYS_TO_DELIVER = "Дней на поставку (ЧИСЛО)"
COL_ORDER_DATE = "ЗАКАЗ ВЗЯТ В РАБОТУ (ДАТА) ОТ КЛИЕНТА"
COL_DELIVERY = "ФАКТИЧЕСКАЯ ДАТА ПОСТАВКИ (СОГЛАСНО УСЛОВИЯМ ПОСТАВКИ)"
COL_COMMENT = "Комментарии"

STATUS_TROUBLE = "6 TROUBLE"
STATUS_CANCEL = "5 CANCELLED"
STATUS_REFUND = "7 REFUND"
STATUS_WARRANTY = "8 WARRANTY"
STATUS_FINISHED = "4 FINISHED"
STATUS_SHIPPED = "3 SHIPPED"

# Statuses that can “become cancelled” during the week (user rule)
CANCEL_SOURCE = {"1 NOT PAID", "2 PAID", STATUS_TROUBLE, "0 NOT PAID"}
TERMINAL_BAD = {STATUS_CANCEL, STATUS_REFUND}
RESOLVED_FROM_TROUBLE = {"1 NOT PAID", "2 PAID", STATUS_SHIPPED, STATUS_FINISHED}


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


def parse_report_date(text: str) -> date:
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(text.strip(), fmt).date()
        except ValueError:
            continue
    raise ValueError(f"Cannot parse date: {text!r}")


def normalize_invoice(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip()
    if text.endswith(".0") and text[:-2].isdigit():
        return text[:-2]
    return text


def load_taz(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=0)
    df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]
    rename: dict[str, str] = {}
    for c in df.columns:
        if isinstance(c, str) and c.startswith("Стоимость доставки ПЛАН"):
            rename[c] = COL_TRANSPORT_PLAN
    if rename:
        df = df.rename(columns=rename)
    return df


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


COMMERCIAL_EPS = 1.0  # USD — ниже считаем «без изменений»


def row_margin_plan(row: pd.Series) -> float:
    sale = parse_numeric(row.get(COL_SALE))
    purchase = parse_numeric(row.get(COL_PURCHASE))
    transport = parse_numeric(row.get(COL_TRANSPORT_PLAN))
    fee = parse_numeric(row.get(COL_FEE))
    return sale - purchase - transport - fee


def fmt_money(v: float | None, digits: int = 0) -> str:
    if v is None:
        return "—"
    return f"{v:,.{digits}f}".replace(",", " ")


def fmt_pct(v: float | None) -> str:
    if v is None:
        return "—"
    return f"{v:.1f}%"


def fmt_days_range(min_d: int | None, max_d: int | None) -> str:
    if min_d is None and max_d is None:
        return "—"
    if min_d is not None and max_d is not None:
        if min_d == max_d:
            return f"{min_d} дн."
        return f"{min_d}–{max_d} дн."
    if min_d is not None:
        return f"≥{min_d} дн."
    return f"≤{max_d} дн."


def esc(text: Any) -> str:
    return html.escape(str(text or ""), quote=True)


def index_rows(df: pd.DataFrame) -> dict[str, pd.Series]:
    keyed = add_row_key(df).drop_duplicates(subset="_key", keep="first")
    keyed = keyed.set_index("_key", drop=False)
    return {str(k): keyed.loc[k] for k in keyed.index}


@dataclass
class StatusChange:
    key: str
    invoice: str
    customer: str
    pn: str
    description: str
    category: str
    supplier: str
    root_supplier: str
    prev_status: str
    curr_status: str
    change_kind: str
    sale_prev: float
    sale_curr: float
    margin_prev: float
    margin_curr: float
    qty_prev: float
    qty_curr: float
    deadline_prev: date | None
    deadline_curr: date | None
    days_to_deliver_prev: float | None
    days_to_deliver_curr: float | None
    trouble_min_days: int | None = None
    trouble_max_days: int | None = None
    notes: list[str] = field(default_factory=list)

    @property
    def margin_delta(self) -> float:
        return self.margin_curr - self.margin_prev

    @property
    def sale_delta(self) -> float:
        return self.sale_curr - self.sale_prev

    @property
    def deadline_delta_days(self) -> int | None:
        if self.deadline_prev and self.deadline_curr:
            return (self.deadline_curr - self.deadline_prev).days
        return None


def build_notes(category: str, qty_prev: float, qty_curr: float) -> list[str]:
    notes: list[str] = []
    cat = (category or "").upper()
    if qty_prev and qty_curr and abs(qty_prev - qty_curr) > 0.009:
        notes.append(f"кол-во {qty_prev:g} → {qty_curr:g}")
        if "EXP" in cat:
            notes.append("EXP: вероятно правка количества / ед. изм.")
        elif "ROT" in cat:
            notes.append("ROTABLE: проверьте замену юнита")
    return notes


def build_commercial_notes(prev: pd.Series, curr: pd.Series) -> list[str]:
    """Детализация изменений по коммерции между двумя снимками."""
    notes: list[str] = []

    def money_delta(label: str, prev_v: float, curr_v: float, up: str, down: str) -> str | None:
        d = curr_v - prev_v
        if abs(d) < COMMERCIAL_EPS:
            return None
        return f"{label} {up if d > 0 else down} на {fmt_money(abs(d), 0)} USD"

    sale_p = parse_numeric(prev.get(COL_SALE))
    sale_c = parse_numeric(curr.get(COL_SALE))
    purch_p = parse_numeric(prev.get(COL_PURCHASE))
    purch_c = parse_numeric(curr.get(COL_PURCHASE))
    trans_p = parse_numeric(prev.get(COL_TRANSPORT_PLAN))
    trans_c = parse_numeric(curr.get(COL_TRANSPORT_PLAN))
    fee_p = parse_numeric(prev.get(COL_FEE))
    fee_c = parse_numeric(curr.get(COL_FEE))

    for part in (
        money_delta("закупка", purch_p, purch_c, "увеличилась", "уменьшилась"),
        money_delta("транспорт", trans_p, trans_c, "увеличился", "уменьшился"),
        money_delta("продажа", sale_p, sale_c, "увеличилась", "уменьшилась"),
        money_delta("fee", fee_p, fee_c, "увеличился", "уменьшился"),
    ):
        if part:
            notes.append(part)

    sup_p = str(prev.get(COL_SUPPLIER) or "").strip()
    sup_c = str(curr.get(COL_SUPPLIER) or "").strip()
    if sup_p.lower() != sup_c.lower() and (sup_p or sup_c):
        notes.append(f"поставщик: {sup_p or '—'} → {sup_c or '—'}")

    root_p = str(prev.get(COL_ROOT_SUPPLIER) or "").strip()
    root_c = str(curr.get(COL_ROOT_SUPPLIER) or "").strip()
    if root_p.lower() != root_c.lower() and (root_p or root_c):
        notes.append(f"root supplier: {root_p or '—'} → {root_c or '—'}")

    return notes


def fmt_margin_delta_cell(e: StatusChange) -> tuple[str, str]:
    """Δ маржа план = (продажа − закупка − транспорт − fee)_curr − _prev."""
    d = e.margin_delta
    has_commercial = any(
        kw in " ".join(e.notes)
        for kw in ("закупка", "транспорт", "продажа", "fee", "поставщик", "root supplier")
    )
    if abs(d) < COMMERCIAL_EPS and not has_commercial:
        return "—", ""
    css = "pos" if d > COMMERCIAL_EPS else ("neg" if d < -COMMERCIAL_EPS else "")
    return fmt_money(d, 0), css


def compare_snapshots(
    prev_rows: dict[str, pd.Series],
    curr_rows: dict[str, pd.Series],
    period_days: int,
) -> tuple[list[StatusChange], list[StatusChange], list[StatusChange], list[StatusChange]]:
    """Return trouble_events, cancellations, refunds, warranty_transitions."""
    trouble: list[StatusChange] = []
    cancellations: list[StatusChange] = []
    refunds: list[StatusChange] = []
    warranty: list[StatusChange] = []

    common = set(prev_rows) & set(curr_rows)
    for key in common:
        prev = prev_rows[key]
        curr = curr_rows[key]
        ps = str(prev.get(COL_STATUS) or "").strip()
        cs = str(curr.get(COL_STATUS) or "").strip()
        if ps == cs:
            continue

        sale_prev = parse_numeric(prev.get(COL_SALE))
        sale_curr = parse_numeric(curr.get(COL_SALE))
        qty_prev = parse_numeric(prev.get(COL_QTY))
        qty_curr = parse_numeric(curr.get(COL_QTY))
        category = str(prev.get(COL_CATEGORY) or curr.get(COL_CATEGORY) or "").strip()
        qty_notes = build_notes(category, qty_prev, qty_curr)
        commercial_notes = build_commercial_notes(prev, curr)

        base = StatusChange(
            key=key,
            invoice=normalize_invoice(prev.get(COL_INVOICE)),
            customer=str(prev.get(COL_CUSTOMER) or curr.get(COL_CUSTOMER) or "").strip(),
            pn=str(prev.get(COL_PN) or "").strip(),
            description=str(prev.get(COL_DESC) or "").strip(),
            category=category,
            supplier=str(prev.get(COL_SUPPLIER) or curr.get(COL_SUPPLIER) or "").strip(),
            root_supplier=str(
                prev.get(COL_ROOT_SUPPLIER) or curr.get(COL_ROOT_SUPPLIER) or ""
            ).strip(),
            prev_status=ps,
            curr_status=cs,
            change_kind="",
            sale_prev=sale_prev,
            sale_curr=sale_curr,
            margin_prev=row_margin_plan(prev),
            margin_curr=row_margin_plan(curr),
            qty_prev=qty_prev,
            qty_curr=qty_curr,
            deadline_prev=parse_date(prev.get(COL_DEADLINE)),
            deadline_curr=parse_date(curr.get(COL_DEADLINE)),
            days_to_deliver_prev=parse_numeric(prev.get(COL_DAYS_TO_DELIVER)) or None,
            days_to_deliver_curr=parse_numeric(curr.get(COL_DAYS_TO_DELIVER)) or None,
            notes=qty_notes,
        )

        # --- TROUBLE lifecycle ---
        if ps == STATUS_TROUBLE and cs == STATUS_TROUBLE:
            trouble.append(
                replace(
                    base,
                    change_kind="ongoing",
                    trouble_min_days=period_days,
                    trouble_max_days=None,
                    notes=[*qty_notes, *commercial_notes, f"в TROUBLE минимум {period_days} дн. (оба снимка)"],
                )
            )
        elif ps != STATUS_TROUBLE and cs == STATUS_TROUBLE:
            trouble.append(
                replace(
                    base,
                    change_kind="entered",
                    trouble_min_days=0,
                    trouble_max_days=period_days,
                    notes=[*qty_notes, *commercial_notes, "новый TROUBLE за период"],
                )
            )
        elif ps == STATUS_TROUBLE and cs in RESOLVED_FROM_TROUBLE:
            ev_notes = [*qty_notes, *commercial_notes]
            dd = base.deadline_delta_days
            if dd is not None and dd != 0:
                ev_notes.append(f"срок поставки сдвинут на {dd:+d} дн.")
            dtd_prev = base.days_to_deliver_prev
            dtd_curr = base.days_to_deliver_curr
            if dtd_prev is not None and dtd_curr is not None and abs(dtd_prev - dtd_curr) > 0.1:
                ev_notes.append(f"«дней на поставку» {dtd_prev:g} → {dtd_curr:g}")
            if not commercial_notes:
                ev_notes.append("коммерция без изменений")
            trouble.append(
                replace(
                    base,
                    change_kind="resolved",
                    trouble_min_days=0,
                    trouble_max_days=period_days,
                    notes=ev_notes,
                )
            )
        elif ps == STATUS_TROUBLE and cs in TERMINAL_BAD:
            trouble.append(
                replace(
                    base,
                    change_kind="cancelled_from_trouble",
                    trouble_min_days=0,
                    trouble_max_days=period_days,
                    notes=[*qty_notes, *commercial_notes, "TROUBLE → отмена/возврат"],
                )
            )
        elif ps == STATUS_TROUBLE and cs == STATUS_WARRANTY:
            trouble.append(
                replace(
                    base,
                    change_kind="warranty_from_trouble",
                    trouble_min_days=0,
                    trouble_max_days=period_days,
                    notes=[*qty_notes, *commercial_notes, "TROUBLE → гарантия"],
                )
            )

        # --- Cancellations / refunds (indirect date) ---
        if ps in CANCEL_SOURCE and cs == STATUS_CANCEL:
            cancellations.append(
                replace(base, change_kind="cancelled", notes=[*qty_notes, *commercial_notes])
            )
        elif ps in CANCEL_SOURCE and cs == STATUS_REFUND:
            refunds.append(
                replace(base, change_kind="refunded", notes=[*qty_notes, *commercial_notes])
            )

        # --- Warranty (active pipeline → warranty, same rule as cancellations) ---
        if ps in CANCEL_SOURCE and cs == STATUS_WARRANTY:
            warranty.append(
                replace(base, change_kind="warranty", notes=[*qty_notes, *commercial_notes])
            )

    return trouble, cancellations, refunds, warranty


def avg(vals: list[float | int | None]) -> float | None:
    clean = [float(v) for v in vals if v is not None]
    return mean(clean) if clean else None


def trouble_kpis(events: list[StatusChange], period_days: int) -> dict[str, Any]:
    entered = [e for e in events if e.change_kind == "entered"]
    ongoing = [e for e in events if e.change_kind == "ongoing"]
    resolved = [e for e in events if e.change_kind == "resolved"]
    cancelled = [e for e in events if e.change_kind == "cancelled_from_trouble"]
    warranty = [e for e in events if e.change_kind == "warranty_from_trouble"]

    at_start = len(ongoing) + len(resolved) + len(cancelled) + len(warranty)
    closed = len(resolved) + len(cancelled) + len(warranty)
    resolve_rate = (len(resolved) / at_start * 100) if at_start else None
    cancel_rate = (len(cancelled) / at_start * 100) if at_start else None

    ongoing_days = [e.trouble_min_days for e in ongoing if e.trouble_min_days is not None]
    resolved_margin = [e.margin_delta for e in resolved]
    positive = sum(1 for d in resolved_margin if d > 1)
    negative = sum(1 for d in resolved_margin if d < -1)

    return {
        "entered": len(entered),
        "ongoing": len(ongoing),
        "resolved": len(resolved),
        "cancelled": len(cancelled),
        "warranty": len(warranty),
        "at_start": at_start,
        "resolve_rate": resolve_rate,
        "cancel_rate": cancel_rate,
        "avg_ongoing_days": avg(ongoing_days),
        "avg_margin_delta": avg(resolved_margin),
        "resolved_positive": positive,
        "resolved_negative": negative,
        "period_days": period_days,
    }


FASTAIR_CSS = """
:root {
  --bg:#e7ecef; --card:#ffffff; --ink:#202020; --muted:#4c4c4c;
  --navy:#022f40; --cyan:#d5fbff; --orange:#fe621d; --line:#c4c4c4;
}
* { box-sizing:border-box; }
body {
  margin:0; font-family: Arial, Calibri, Helvetica, sans-serif;
  color:var(--ink); background: linear-gradient(180deg, #022f40 0 160px, var(--bg) 160px);
  line-height:1.45;
}
.wrap { max-width:1200px; margin:0 auto; padding:28px 20px 64px; }
.brand { display:flex; align-items:center; justify-content:space-between; gap:12px; color:#fff; margin-bottom:18px; }
.brand-mark { font-size:13px; letter-spacing:.12em; text-transform:uppercase; background:var(--cyan); color:var(--navy); padding:6px 10px; font-weight:700; }
h1 { font-size:30px; margin:0 0 6px; color:#fff; font-weight:700; }
.sub { color:#d5fbff; margin-bottom:18px; font-size:14px; }
h2 { font-size:20px; margin:28px 0 12px; color:var(--navy); }
h3 { font-size:16px; margin:18px 0 10px; color:var(--navy); }
.card { background:var(--card); border:1px solid var(--line); border-radius:8px; padding:18px; }
.kpi-grid { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:12px; margin:12px 0 18px; }
.kpi-grid > div { background:#f7fbfc; border:1px solid #d7e2e6; border-radius:8px; padding:12px 14px; }
.label { font-size:11px; text-transform:uppercase; letter-spacing:.04em; color:#757575; }
.value { font-size:20px; font-weight:700; margin-top:4px; font-variant-numeric:tabular-nums; color:var(--navy); }
.muted { color:#757575; font-size:12px; margin-top:4px; }
.trouble-block { border:2px solid var(--orange); border-radius:10px; padding:16px; background:#fff9f5; margin:20px 0; }
.trouble-block h2 { margin-top:0; color:#b3470f; }
details.group { background:var(--card); border:1px solid var(--line); border-radius:8px; margin:10px 0; overflow:hidden; }
details.group summary { cursor:pointer; list-style:none; padding:14px 16px; display:flex; flex-wrap:wrap; gap:8px; align-items:center; background:linear-gradient(90deg,#022f40 0%,#022f40 12px,#fff 12px); }
details.group summary::-webkit-details-marker { display:none; }
.group-name { font-weight:700; font-size:15px; margin-right:auto; color:var(--navy); padding-left:8px; }
.pill { font-size:12px; background:var(--cyan); border:1px solid #b6e8ef; color:var(--navy); border-radius:4px; padding:4px 10px; font-weight:700; }
.pill.warn { background:#ffe8d9; border-color:#ffc4a3; color:#b3470f; }
.pill.ok { background:#e6f7ea; border-color:#b8e6c1; color:#1f6b32; }
.group-body { padding:0 14px 16px; border-top:1px solid var(--line); }
details.client { margin:8px 0; border-color:#d7e2e6; }
details.client summary { padding:10px 14px; background:linear-gradient(90deg,#022f40 0%,#022f40 8px,#f7fbfc 8px); }
details.client .group-name { font-size:14px; font-weight:600; }
details.client .group-body { padding:0 8px 12px; }
table { width:100%; border-collapse:collapse; font-size:13px; }
th, td { border-bottom:1px solid #e8e8e8; padding:8px 6px; vertical-align:top; text-align:left; }
th { font-size:11px; text-transform:uppercase; letter-spacing:.03em; color:#4c4c4c; background:#f3f7f8; }
td.num { text-align:right; font-variant-numeric:tabular-nums; white-space:nowrap; }
.pos { color:#1f6b32; } .neg { color:#c0392b; }
.footer { margin-top:28px; color:#4c4c4c; font-size:12px; }
@media (max-width:900px) { .kpi-grid { grid-template-columns:repeat(2,minmax(0,1fr)); } }
"""


def group_by_customer(items: list[StatusChange]) -> list[tuple[str, list[StatusChange]]]:
    by_client: dict[str, list[StatusChange]] = defaultdict(list)
    for e in items:
        name = e.customer.strip() if e.customer and e.customer.strip() else "(без клиента)"
        by_client[name].append(e)
    return sorted(by_client.items(), key=lambda x: (-sum(i.sale_curr for i in x[1]), x[0]))


def render_trouble_table(items: list[StatusChange]) -> str:
    if not items:
        return "<p class='muted'>Нет изменений за период.</p>"
    rows = []
    for e in sorted(items, key=lambda x: (x.invoice, x.pn)):
        margin_txt, margin_cls = fmt_margin_delta_cell(e)
        dd = e.deadline_delta_days
        dd_txt = f"{dd:+d}" if dd is not None else "—"
        root = esc(e.root_supplier) if e.root_supplier else "—"
        rows.append(
            f"""<tr>
  <td>{esc(e.invoice)}</td>
  <td>{esc(e.pn)}<div class='muted'>{esc(e.description[:50])}</div></td>
  <td>{esc(e.category)}</td>
  <td>{esc(e.supplier) or "—"}</td>
  <td>{root}</td>
  <td class='num'>{fmt_days_range(e.trouble_min_days, e.trouble_max_days)}</td>
  <td class='num'>{fmt_money(e.sale_prev, 0)} → {fmt_money(e.sale_curr, 0)}</td>
  <td class='num {margin_cls}'>{margin_txt}</td>
  <td class='num'>{dd_txt}</td>
  <td>{esc('; '.join(e.notes) or '—')}</td>
</tr>"""
        )
    return f"""<table>
  <thead><tr>
    <th>Счёт</th><th>P/N</th><th>Cat.</th><th>Поставщик</th><th>Root supplier</th>
    <th>В TROUBLE</th><th>Продажа USD</th><th>Δ маржа</th><th>Δ срок</th><th>Комментарий</th>
  </tr></thead>
  <tbody>{''.join(rows)}</tbody>
</table>"""


def render_cancel_table(items: list[StatusChange]) -> str:
    if not items:
        return "<p class='muted'>Нет изменений за период.</p>"
    rows = []
    for e in sorted(items, key=lambda x: (x.invoice, x.pn)):
        margin_txt, margin_cls = fmt_margin_delta_cell(e)
        kind = "отмена" if e.curr_status == STATUS_CANCEL else "возврат"
        if e.change_kind == "cancelled_from_trouble":
            kind = "из TROUBLE → " + kind
        rows.append(
            f"""<tr>
  <td>{esc(e.invoice)}</td>
  <td>{esc(e.pn)}<div class='muted'>{esc(e.description[:50])}</div></td>
  <td>{esc(e.category)}</td>
  <td>{esc(kind)}</td>
  <td class='num'>{fmt_money(e.sale_prev, 0)}</td>
  <td class='num {margin_cls}'>{margin_txt}</td>
  <td>{esc('; '.join(e.notes) or '—')}</td>
</tr>"""
        )
    return f"""<table>
  <thead><tr>
    <th>Счёт</th><th>P/N</th><th>Cat.</th><th>Тип</th>
    <th>Продажа USD</th><th>Δ маржа</th><th>Комментарий</th>
  </tr></thead>
  <tbody>{''.join(rows)}</tbody>
</table>"""


def render_client_groups(
    items: list[StatusChange],
    table_fn,
) -> str:
    if not items:
        return "<p class='muted'>Нет изменений за период.</p>"
    blocks = []
    for client, client_items in group_by_customer(items):
        sale = sum(e.sale_curr for e in client_items)
        margin_d = sum(e.margin_delta for e in client_items)
        blocks.append(
            f"""
<details class="group client">
  <summary>
    <span class="group-name">{esc(client)}</span>
    <span class="pill">{len(client_items)} поз.</span>
    <span class="pill">{fmt_money(sale, 0)} USD</span>
    <span class="pill">Δ {fmt_money(margin_d, 0)}</span>
  </summary>
  <div class="group-body">{table_fn(client_items)}</div>
</details>"""
        )
    return "".join(blocks)


def render_section(
    title: str,
    items: list[StatusChange],
    table_fn,
    pill_class: str = "",
) -> str:
    if not items:
        return f"""
<details class="group">
  <summary>
    <span class="group-name">{esc(title)}</span>
    <span class="pill">0 поз.</span>
  </summary>
  <div class="group-body"><p class='muted'>Нет изменений за период.</p></div>
</details>"""
    sale = sum(e.sale_curr for e in items)
    margin_d = sum(e.margin_delta for e in items)
    clients = len(group_by_customer(items))
    pill = f"pill {pill_class}".strip()
    inner = render_client_groups(items, table_fn)
    return f"""
<details class="group">
  <summary>
    <span class="group-name">{esc(title)}</span>
    <span class="{pill}">{len(items)} поз.</span>
    <span class="pill">{clients} кли.</span>
    <span class="pill">продажа {fmt_money(sale, 0)} USD</span>
    <span class="pill">Δ маржа {fmt_money(margin_d, 0)}</span>
  </summary>
  <div class="group-body">{inner}</div>
</details>"""


def merge_cancel_refund(
    cancellations: list[StatusChange],
    refunds: list[StatusChange],
    trouble: list[StatusChange],
) -> list[StatusChange]:
    """Combine cancel/refund events, dedupe by row key."""
    by_key: dict[str, StatusChange] = {}
    for items in (cancellations, refunds):
        for e in items:
            by_key[e.key] = e
    for e in trouble:
        if e.change_kind == "cancelled_from_trouble":
            by_key[e.key] = e
    return list(by_key.values())


def render_html_report(
    prev_date: date,
    curr_date: date,
    trouble: list[StatusChange],
    cancellations: list[StatusChange],
    refunds: list[StatusChange],
    warranty: list[StatusChange],
) -> str:
    period_days = max(1, (curr_date - prev_date).days)
    kpis = trouble_kpis(trouble, period_days)
    period = f"{prev_date.strftime('%d.%m.%Y')} → {curr_date.strftime('%d.%m.%Y')} ({period_days} дн.)"

    new_trouble = [e for e in trouble if e.change_kind == "entered"]
    resolved_trouble = [e for e in trouble if e.change_kind == "resolved"]
    cancel_refund = merge_cancel_refund(cancellations, refunds, trouble)

    sections = "\n".join([
        render_section("Новые TROUBLE", new_trouble, render_trouble_table, "warn"),
        render_section("Решённые TROUBLE", resolved_trouble, render_trouble_table, "ok"),
        render_section("Отмены и возвраты", cancel_refund, render_cancel_table, "warn"),
    ])

    return f"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8"/>
<title>TROUBLE / CANCEL — {curr_date.strftime('%d.%m.%Y')}</title>
<style>{FASTAIR_CSS}</style>
</head>
<body>
<div class="wrap">
  <div class="brand"><div class="brand-mark">FASTAIR</div><div style="font-size:13px;opacity:.9">TROUBLE / CANCEL</div></div>
  <h1>TROUBLE / CANCEL</h1>
  <div class="sub">Период {period} · сравнение двух выгрузок ТАЗ</div>

  <section class="card">
    <h2 style="margin-top:0">Сводка</h2>
    <div class="kpi-grid">
      <div><div class="label">Новые TROUBLE</div><div class="value">{len(new_trouble)}</div></div>
      <div><div class="label">Решённые TROUBLE</div><div class="value">{len(resolved_trouble)}<div class="muted">{fmt_pct(kpis['resolve_rate'])} от стартовых</div></div></div>
      <div><div class="label">Отмены + возвраты</div><div class="value">{len(cancel_refund)}</div></div>
      <div><div class="label">Всё ещё в TROUBLE</div><div class="value">{kpis['ongoing']}<div class="muted">без смены статуса за период</div></div></div>
      <div><div class="label">Δ маржа (решённые)</div><div class="value">{fmt_money(kpis['avg_margin_delta'], 0)}<div class="muted">+{kpis['resolved_positive']} / −{kpis['resolved_negative']}</div></div></div>
      <div><div class="label">Гарантии (новые)</div><div class="value">{len(warranty)}</div></div>
    </div>
    <p class="muted">TROUBLE: EXP — кол-во/ед. изм.; ROTABLE — замена юнита. Отмена/возврат: было NOT PAID / PAID / TROUBLE → CANCELLED / REFUND.</p>
  </section>

  {sections}

  <div class="footer">
    FASTAIR · источник ТАЗ · {esc(prev_date.isoformat())} → {esc(curr_date.isoformat())} ·
    ключ строки: счёт + P/N + описание
  </div>
</div>
</body>
</html>
"""


def main() -> None:
    parser = argparse.ArgumentParser(description="HTML report: TAZ diff (TROUBLE, cancel, warranty)")
    parser.add_argument("--previous", "-p", type=Path, required=True, help="Earlier TAZ snapshot")
    parser.add_argument("--current", "-c", type=Path, required=True, help="Later TAZ snapshot")
    parser.add_argument("--previous-date", type=str, help="Date of previous snapshot DD.MM.YYYY")
    parser.add_argument("--current-date", type=str, help="Date of current snapshot DD.MM.YYYY")
    parser.add_argument("--output", "-o", type=Path, default=Path("output/changes/trouble_cancel.html"))
    args = parser.parse_args()

    prev_date = parse_report_date(args.previous_date) if args.previous_date else None
    curr_date = parse_report_date(args.current_date) if args.current_date else None
    if prev_date is None:
        prev_date = date.fromtimestamp(args.previous.stat().st_mtime)
    if curr_date is None:
        curr_date = date.fromtimestamp(args.current.stat().st_mtime)

    print(f"Loading previous TAZ: {args.previous} ({prev_date})")
    print(f"Loading current TAZ:  {args.current} ({curr_date})")
    prev_rows = index_rows(load_taz(args.previous))
    curr_rows = index_rows(load_taz(args.current))
    period_days = max(1, (curr_date - prev_date).days)

    trouble, cancellations, refunds, warranty = compare_snapshots(prev_rows, curr_rows, period_days)
    html_out = render_html_report(prev_date, curr_date, trouble, cancellations, refunds, warranty)

    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(html_out, encoding="utf-8")
    zip_path = args.output.with_suffix(".zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.write(args.output, args.output.name)
    print(f"Saved {args.output}")
    print(f"Saved {zip_path}  ← скачайте ZIP и откройте HTML из распакованной папки")
    print(
        f"TROUBLE events: {len(trouble)} | cancellations: {len(cancellations)} | "
        f"refunds: {len(refunds)} | warranty: {len(warranty)}"
    )


if __name__ == "__main__":
    main()
