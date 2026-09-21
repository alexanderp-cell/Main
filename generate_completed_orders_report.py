#!/usr/bin/env python3
"""FASTAIR — Выполненные заказы: сроки поставки + оплата JT/KT + транспорт.

Периоды (по столбцу W — факт. дата поставки):
  1) прошедшая неделя (7 дней, заканчивая датой среза ТАЗ)
  2) 01.06 — 18.09.2026
  3) весь 2026

В каждом периоде:
  1) срок поставок по клиентам и по поставщикам (STK, FINISHED, W−Q);
     для IBERIA и JET TECHNIC дополнительно срок оплаты (AW−Q, STK, без постоплаты)
  2) стоимость транспорта план vs факт
"""

from __future__ import annotations

import argparse
import math
import re
import zipfile
from dataclasses import dataclass, field
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

COL_STATUS = "Status"
COL_CUSTOMER = "Customer"
COL_CATEGORY = "Category"
COL_LEAD_TIME = "Lead time"
COL_ORDER_DATE = "ЗАКАЗ ВЗЯТ В РАБОТУ (ДАТА) ОТ КЛИЕНТА"
COL_DELIVERY = "ФАКТИЧЕСКАЯ ДАТА ПОСТАВКИ (СОГЛАСНО УСЛОВИЯМ ПОСТАВКИ)"
COL_SUPPLIER = "Поставщик"
COL_PAY_DATE = "Дата оплаты поставщику"
COL_MOVEMENT = "Дата начала движения"
COL_TRANSPORT_PLAN = (
    "Стоимость доставки ПЛАН, за весь счет! Если в счете несколько строк, "
    "то \"размазываем\" равномерно планируюмую стоиомость транспорта на все позиции из счета."
)
COL_TRANSPORT_FACT = "Стоимость доставки факт"

CAT_ROTABLE = "ROTABLE"
CAT_EXPENDABLE = "EXPENDABLE"

PAY_SUPPLIERS = {"IBERIA", "JET TECHNIC"}


def parse_numeric(value) -> float:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return 0.0
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if isinstance(value, float) and math.isnan(value):
            return 0.0
        return float(value)
    text = str(value).strip().replace("\xa0", "").replace(" ", "").replace(",", ".")
    if not text or text.startswith("#"):
        return 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0


def parse_date(value) -> date | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    ts = pd.to_datetime(value, errors="coerce", dayfirst=True)
    if pd.isna(ts):
        return None
    return ts.date()


def html_escape(text: str) -> str:
    return (
        str(text)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def fmt_days(v: float | None) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "—"
    return f"{float(v):.1f}"


def fmt_money(v: float) -> str:
    return f"{v:,.0f}".replace(",", " ")


def fmt_pct(v: float | None) -> str:
    if v is None:
        return "—"
    return f"{v:.1f}%"


def load_taz(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=0)
    df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]
    for c in list(df.columns):
        if isinstance(c, str) and c.startswith("Дата начала движения"):
            if c != COL_MOVEMENT:
                df = df.rename(columns={c: COL_MOVEMENT})
            break
    # transport plan column may differ slightly — match by prefix
    if COL_TRANSPORT_PLAN not in df.columns:
        for c in df.columns:
            if isinstance(c, str) and c.startswith("Стоимость доставки ПЛАН"):
                df = df.rename(columns={c: COL_TRANSPORT_PLAN})
                break
    return df


def prepare(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["_status"] = out[COL_STATUS].astype(str).str.strip().str.upper()
    out["_finished"] = out["_status"].str.contains("FINISHED", na=False)
    out["_lead"] = out[COL_LEAD_TIME].astype(str).str.strip().str.upper()
    out["_stk"] = out["_lead"].eq("STK")
    out["_client"] = out[COL_CUSTOMER].fillna("").map(
        lambda v: "" if (isinstance(v, float) and pd.isna(v)) else str(v).strip()
    )
    out.loc[out["_client"].str.lower().isin({"", "nan", "none", "<na>"}), "_client"] = ""
    out["_supplier"] = out[COL_SUPPLIER].fillna("").map(
        lambda v: "" if (isinstance(v, float) and pd.isna(v)) else str(v).strip()
    )
    out.loc[
        out["_supplier"].eq("") | out["_supplier"].str.lower().isin({"nan", "none", "<na>"}),
        "_supplier",
    ] = "— без поставщика —"
    out["_cat"] = out[COL_CATEGORY].astype(str).str.strip().str.upper()
    out["_q"] = out[COL_ORDER_DATE].map(parse_date)
    out["_w"] = out[COL_DELIVERY].map(parse_date)
    out["_aw"] = out[COL_PAY_DATE].map(parse_date)
    if COL_MOVEMENT in out.columns:
        out["_ba"] = out[COL_MOVEMENT].map(parse_date)
    else:
        out["_ba"] = None
    out["_days"] = [
        (w - q).days if w and q else None for w, q in zip(out["_w"], out["_q"])
    ]
    out["_pay_days"] = [
        (aw - q).days if aw and q else None for aw, q in zip(out["_aw"], out["_q"])
    ]
    plan_col = COL_TRANSPORT_PLAN if COL_TRANSPORT_PLAN in out.columns else None
    fact_col = COL_TRANSPORT_FACT if COL_TRANSPORT_FACT in out.columns else None
    out["_plan"] = out[plan_col].map(parse_numeric) if plan_col else 0.0
    out["_fact"] = out[fact_col].map(parse_numeric) if fact_col else 0.0
    return out


def in_period(d: date | None, start: date, end: date) -> bool:
    return d is not None and start <= d <= end


def lead_mask(df: pd.DataFrame, start: date, end: date) -> pd.Series:
    return (
        df["_finished"]
        & df["_stk"]
        & df["_w"].map(lambda d: in_period(d, start, end))
        & df["_q"].notna()
        & df["_days"].notna()
        & (df["_days"] >= 0)
        & df["_cat"].isin({CAT_ROTABLE, CAT_EXPENDABLE})
    )


def transport_mask(df: pd.DataFrame, start: date, end: date) -> pd.Series:
    return df["_finished"] & df["_w"].map(lambda d: in_period(d, start, end))


def pay_usable_mask(rows: pd.DataFrame) -> pd.Series:
    """STK payment speed: AW−Q, exclude postpay and unknown/positive LT."""
    ba = rows["_ba"]
    aw = rows["_aw"]
    postpay = ba.notna() & (aw.isna() | (ba < aw))
    return (
        aw.notna()
        & rows["_q"].notna()
        & rows["_pay_days"].notna()
        & (rows["_pay_days"] >= 0)
        & ~postpay
        & rows["_stk"]
    )


@dataclass
class SideStats:
    avg: float | None = None
    n: int = 0


@dataclass
class EntityLead:
    name: str
    rotable: SideStats = field(default_factory=SideStats)
    expendable: SideStats = field(default_factory=SideStats)
    pay_avg: float | None = None
    pay_median: float | None = None
    pay_n: int = 0
    show_pay: bool = False
    # why pay_n < delivery n (IBERIA / JET TECHNIC)
    pay_excl_unpaid: int = 0
    pay_excl_negative: int = 0
    pay_excl_postpay: int = 0

    @property
    def delivery_n(self) -> int:
        return self.rotable.n + self.expendable.n


def pay_exclusion_breakdown(part: pd.DataFrame) -> tuple[pd.DataFrame, int, int, int]:
    """Return usable rows + mutually exclusive exclusion counts vs delivery set."""
    no_aw = part["_aw"].isna()
    neg = part["_aw"].notna() & part["_pay_days"].notna() & (part["_pay_days"] < 0)
    ba = part["_ba"]
    aw = part["_aw"]
    post = part["_aw"].notna() & ~neg & ba.notna() & (ba < aw)
    usable = part.loc[pay_usable_mask(part)]
    return usable, int(no_aw.sum()), int(neg.sum()), int(post.sum())


def side_stats(series: pd.Series) -> SideStats:
    if series.empty:
        return SideStats(None, 0)
    return SideStats(float(series.mean()), int(len(series)))


def build_entities(
    rows: pd.DataFrame, key: str, *, with_pay: bool
) -> list[EntityLead]:
    names = sorted(rows[key].unique(), key=lambda s: str(s).casefold())
    out: list[EntityLead] = []
    for name in names:
        if not name or str(name).lower() in {"nan", "none"}:
            continue
        part = rows[rows[key] == name]
        ent = EntityLead(
            name=str(name),
            rotable=side_stats(part.loc[part["_cat"] == CAT_ROTABLE, "_days"].astype(float)),
            expendable=side_stats(
                part.loc[part["_cat"] == CAT_EXPENDABLE, "_days"].astype(float)
            ),
        )
        if with_pay and str(name).upper() in PAY_SUPPLIERS:
            ent.show_pay = True
            usable, unpaid, negative, postpay = pay_exclusion_breakdown(part)
            ent.pay_excl_unpaid = unpaid
            ent.pay_excl_negative = negative
            ent.pay_excl_postpay = postpay
            if not usable.empty:
                days = usable["_pay_days"].astype(float)
                ent.pay_avg = float(days.mean())
                ent.pay_median = float(days.median())
                ent.pay_n = int(len(days))
        out.append(ent)
    # sort by total n desc, then name
    out.sort(
        key=lambda e: (
            -(e.rotable.n + e.expendable.n),
            e.name.casefold(),
        )
    )
    return out


@dataclass
class TransportRow:
    name: str
    plan: float
    fact: float
    n: int

    @property
    def delta(self) -> float:
        return self.fact - self.plan


@dataclass
class PeriodBlock:
    title: str
    start: date
    end: date
    lead_avg: float | None
    lead_median: float | None
    lead_n: int
    clients: list[EntityLead]
    suppliers: list[EntityLead]
    transport_total_plan: float
    transport_total_fact: float
    transport_n: int
    transport_by_client: list[TransportRow]


def build_transport(rows: pd.DataFrame) -> tuple[float, float, int, list[TransportRow]]:
    total_plan = float(rows["_plan"].sum())
    total_fact = float(rows["_fact"].sum())
    total_n = int(len(rows))
    grouped = (
        rows.groupby("_client", dropna=False)
        .agg(plan=("_plan", "sum"), fact=("_fact", "sum"), n=("_plan", "size"))
        .reset_index()
    )
    items: list[TransportRow] = []
    for _, rec in grouped.iterrows():
        name = str(rec["_client"]).strip() or "— без клиента —"
        if name.lower() in {"nan", "none", "<na>"}:
            name = "— без клиента —"
        plan = float(rec["plan"])
        fact = float(rec["fact"])
        if plan == 0 and fact == 0:
            continue
        items.append(TransportRow(name=name, plan=plan, fact=fact, n=int(rec["n"])))
    items.sort(key=lambda r: (-abs(r.delta), -r.fact, r.name.casefold()))
    return total_plan, total_fact, total_n, items


def build_period(df: pd.DataFrame, title: str, start: date, end: date) -> PeriodBlock:
    lead_rows = df.loc[lead_mask(df, start, end)].copy()
    lead_rows = lead_rows[lead_rows["_client"].ne("")]
    clients = build_entities(lead_rows, "_client", with_pay=False)
    suppliers = build_entities(lead_rows, "_supplier", with_pay=True)

    tr_rows = df.loc[transport_mask(df, start, end)].copy()
    plan, fact, n, by_client = build_transport(tr_rows)

    days = lead_rows["_days"].astype(float) if not lead_rows.empty else pd.Series(dtype=float)
    return PeriodBlock(
        title=title,
        start=start,
        end=end,
        lead_avg=float(days.mean()) if len(days) else None,
        lead_median=float(days.median()) if len(days) else None,
        lead_n=int(len(days)),
        clients=clients,
        suppliers=suppliers,
        transport_total_plan=plan,
        transport_total_fact=fact,
        transport_n=n,
        transport_by_client=by_client,
    )


def avg_cell(avg: float | None, n: int) -> str:
    if n == 0 or avg is None:
        return '<td class="num empty" data-value="">—</td>'
    shown = f"{float(avg):.1f}"
    return f'<td class="num" data-value="{shown}">{shown}</td>'


def n_cell(n: int) -> str:
    if not n:
        return '<td class="num empty" data-value="">—</td>'
    return f'<td class="num" data-value="{n}">{n}</td>'


def render_lead_table(
    entities: list[EntityLead],
    *,
    name_header: str,
    table_id: str,
    with_pay: bool,
    overall_avg: float | None,
    overall_n: int,
) -> str:
    rows = []
    pay_notes = []
    for e in entities:
        cls = ""
        up = e.name.upper()
        if up == "JET TECHNIC":
            cls = "channel-jt"
        elif up == "IBERIA":
            cls = "channel-kt"
        pay_cells = ""
        if with_pay:
            if e.show_pay:
                pay_cells = (
                    f"{avg_cell(e.pay_avg, e.pay_n)}"
                    f"{avg_cell(e.pay_median, e.pay_n)}"
                    f"{n_cell(e.pay_n)}"
                )
                bits = []
                if e.pay_excl_unpaid:
                    bits.append(f"без даты оплаты AW: {e.pay_excl_unpaid}")
                if e.pay_excl_postpay:
                    bits.append(f"постоплата (BA&lt;AW): {e.pay_excl_postpay}")
                if e.pay_excl_negative:
                    bits.append(f"оплата раньше Q: {e.pay_excl_negative}")
                excl = "; ".join(bits) if bits else "исключений нет"
                pay_notes.append(
                    f"<li><strong>{html_escape(e.name)}</strong>: "
                    f"поставка n={e.delivery_n} (ротабл {e.rotable.n} + расходка {e.expendable.n}), "
                    f"в среднее оплаты n={e.pay_n}. "
                    f"Не входят: {excl}.</li>"
                )
            else:
                pay_cells = (
                    '<td class="num empty" data-value="">—</td>'
                    '<td class="num empty" data-value="">—</td>'
                    '<td class="num empty" data-value="">—</td>'
                )
        rows.append(
            f'<tr class="{cls}">'
            f'<td data-value="{html_escape(e.name.casefold())}">{html_escape(e.name)}</td>'
            f"{avg_cell(e.rotable.avg, e.rotable.n)}{n_cell(e.rotable.n)}"
            f"{avg_cell(e.expendable.avg, e.expendable.n)}{n_cell(e.expendable.n)}"
            f"{pay_cells}</tr>"
        )

    pay_headers = ""
    if with_pay:
        pay_headers = (
            '<th data-type="num">Оплата ср., дн. <span class="arrow">↕</span></th>'
            '<th data-type="num">Оплата мед. <span class="arrow">↕</span></th>'
            '<th data-type="num">Оплата n <span class="arrow">↕</span></th>'
        )

    pay_note_html = ""
    if pay_notes:
        pay_note_html = (
            '<div class="note-box">'
            "<strong>Почему «Оплата n» ≠ ротабл n + расходка n</strong>"
            "<p>Ротабл/расходка — все поставки STK в периоде. "
            "Оплата n — только позиции, по которым считаем AW−Q "
            "(есть дата оплаты, нет постоплаты BA&lt;AW, оплата не раньше Q).</p>"
            f"<ul>{''.join(pay_notes)}</ul></div>"
        )

    return f"""
<div class="table-scroll">
<table class="sortable" id="{html_escape(table_id)}">
  <thead>
    <tr>
      <th data-type="str">{html_escape(name_header)} <span class="arrow">↕</span></th>
      <th data-type="num">Ротабл, дн. <span class="arrow">↕</span></th>
      <th data-type="num">Ротабл n <span class="arrow">↕</span></th>
      <th data-type="num">Расходка, дн. <span class="arrow">↕</span></th>
      <th data-type="num">Расходка n <span class="arrow">↕</span></th>
      {pay_headers}
    </tr>
  </thead>
  <tbody>
    {''.join(rows) if rows else '<tr><td colspan="8" class="empty">Нет позиций</td></tr>'}
  </tbody>
</table>
</div>
{pay_note_html}
<p class="hint">
  Срок поставки = W − Q (дн.), статус FINISHED, столбец S (Lead time) = <strong>только STK</strong>
  (числовые lead time / «5 days» и т.п. не входят — как в отчёте по срокам поставки),
  категории ROTABLE / EXPENDABLE.
  Всего позиций в выборке: {overall_n}, средний срок {fmt_days(overall_avg)} дн.
</p>
"""


def render_transport(period: PeriodBlock, table_id: str) -> str:
    delta = period.transport_total_fact - period.transport_total_plan
    pct = (
        (delta / period.transport_total_plan * 100.0)
        if period.transport_total_plan
        else None
    )
    rows = []
    for r in period.transport_by_client:
        d = r.delta
        rows.append(
            f"<tr>"
            f"<td>{html_escape(r.name)}</td>"
            f'<td class="num" data-value="{r.plan}">{fmt_money(r.plan)}</td>'
            f'<td class="num" data-value="{r.fact}">{fmt_money(r.fact)}</td>'
            f'<td class="num" data-value="{d}">{fmt_money(d)}</td>'
            f'<td class="num" data-value="{r.n}">{r.n}</td>'
            f"</tr>"
        )
    return f"""
<div class="kpis three">
  <div class="highlight"><div class="label">План</div><div class="value">{fmt_money(period.transport_total_plan)}</div><div class="muted">USD · сумма по строкам</div></div>
  <div><div class="label">Факт</div><div class="value">{fmt_money(period.transport_total_fact)}</div><div class="muted">{period.transport_n} поз. FINISHED</div></div>
  <div><div class="label">Факт − план</div><div class="value">{fmt_money(delta)}</div><div class="muted">{fmt_pct(pct)} к плану</div></div>
</div>
<div class="table-scroll">
<table class="sortable" id="{html_escape(table_id)}">
  <thead>
    <tr>
      <th data-type="str">Клиент <span class="arrow">↕</span></th>
      <th data-type="num">План, USD <span class="arrow">↕</span></th>
      <th data-type="num">Факт, USD <span class="arrow">↕</span></th>
      <th data-type="num">Δ факт−план <span class="arrow">↕</span></th>
      <th data-type="num">n <span class="arrow">↕</span></th>
    </tr>
  </thead>
  <tbody>
    {''.join(rows) if rows else '<tr><td colspan="5" class="empty">Нет данных по транспорту</td></tr>'}
  </tbody>
</table>
</div>
<p class="hint">
  Период по столбцу W. Статус FINISHED (без фильтра STK).
  План уже «размазан» по строкам счёта в ТАЗ — суммируем строки.
</p>
"""


def render_period(period: PeriodBlock, idx: int, *, opened: bool) -> str:
    open_attr = " open" if opened else ""
    range_s = f"{period.start.strftime('%d.%m.%Y')} — {period.end.strftime('%d.%m.%Y')}"
    return f"""
<details class="period"{open_attr}>
  <summary>
    <span class="period-title">{html_escape(period.title)}</span>
    <span class="period-meta">{range_s} · поставка STK {period.lead_n} поз. · ср. {fmt_days(period.lead_avg)} дн.</span>
  </summary>
  <div class="period-body">
    <details class="block" open>
      <summary>1. Срок поставок</summary>
      <div class="block-body">
        <div class="kpis">
          <div class="highlight"><div class="label">Средний срок</div><div class="value">{fmt_days(period.lead_avg)} дн.</div><div class="muted">медиана {fmt_days(period.lead_median)}</div></div>
          <div><div class="label">Позиций</div><div class="value">{period.lead_n}</div><div class="muted">FINISHED · STK · по W</div></div>
          <div><div class="label">Клиентов</div><div class="value">{len(period.clients)}</div><div class="muted">в таблице</div></div>
          <div><div class="label">Поставщиков</div><div class="value">{len(period.suppliers)}</div><div class="muted">в таблице</div></div>
        </div>

        <details class="subblock" open>
          <summary>По клиентам</summary>
          <div class="block-body">
            {render_lead_table(
                period.clients,
                name_header="Клиент",
                table_id=f"clients-{idx}",
                with_pay=False,
                overall_avg=period.lead_avg,
                overall_n=period.lead_n,
            )}
          </div>
        </details>

        <details class="subblock">
          <summary>По поставщикам <span class="tag">IBERIA / JET TECHNIC — ещё срок оплаты</span></summary>
          <div class="block-body">
            {render_lead_table(
                period.suppliers,
                name_header="Поставщик",
                table_id=f"suppliers-{idx}",
                with_pay=True,
                overall_avg=period.lead_avg,
                overall_n=period.lead_n,
            )}
          </div>
        </details>
      </div>
    </details>

    <details class="block">
      <summary>2. Стоимость транспорта: план vs факт</summary>
      <div class="block-body">
        {render_transport(period, f"transport-{idx}")}
      </div>
    </details>
  </div>
</details>
"""


def render_html(periods: list[PeriodBlock], source_name: str) -> str:
    sections = "\n".join(
        render_period(p, i, opened=(i == 0)) for i, p in enumerate(periods)
    )
    return f"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>FASTAIR — Выполненные заказы</title>
<style>
:root {{
  --bg:#e7ecef; --card:#fff; --ink:#202020; --muted:#5a5a5a;
  --navy:#022f40; --cyan:#d5fbff; --line:#c4c4c4; --zebra:#f5fafb;
  --jt:#022f40; --kt:#c45c26;
}}
* {{ box-sizing:border-box; }}
body {{
  margin:0; font-family: Arial, Calibri, Helvetica, sans-serif;
  color:var(--ink); background: linear-gradient(180deg, #022f40 0 140px, var(--bg) 140px);
  line-height:1.45;
}}
.wrap {{ max-width:1100px; margin:0 auto; padding:28px 20px 64px; }}
.brand {{
  background:var(--cyan); color:var(--navy); display:inline-block;
  padding:6px 10px; font-weight:700; margin-bottom:10px;
  font-size:13px; letter-spacing:.14em; text-transform:uppercase;
}}
h1 {{ color:#fff; font-size:28px; margin:0 0 6px; font-weight:700; }}
.sub {{ color:#d5fbff; margin:0 0 18px; font-size:14px; }}
details.period {{
  background:var(--card); border:1px solid var(--line); border-radius:8px;
  margin-bottom:14px; overflow:hidden;
}}
details.period > summary {{
  cursor:pointer; list-style:none; padding:14px 16px;
  background:#f3f8fa; border-bottom:1px solid transparent;
  display:flex; flex-wrap:wrap; gap:8px 16px; align-items:baseline;
}}
details.period[open] > summary {{ border-bottom-color:var(--line); }}
details.period > summary::-webkit-details-marker {{ display:none; }}
.period-title {{ font-size:17px; font-weight:700; color:var(--navy); }}
.period-meta {{ font-size:13px; color:var(--muted); }}
.period-body {{ padding:12px 16px 16px; }}
details.block, details.subblock {{
  border:1px solid #d7e2e6; border-radius:8px; margin:10px 0; background:#fff;
}}
details.block > summary, details.subblock > summary {{
  cursor:pointer; padding:10px 12px; font-weight:700; color:var(--navy);
  list-style:none; background:#f7fbfc;
}}
details.block > summary::-webkit-details-marker,
details.subblock > summary::-webkit-details-marker {{ display:none; }}
details.block > summary::before,
details.subblock > summary::before {{
  content:"▸ "; color:#7a8a90;
}}
details.block[open] > summary::before,
details.subblock[open] > summary::before {{ content:"▾ "; }}
.block-body {{ padding:12px; }}
.tag {{
  display:inline-block; margin-left:8px; font-size:11px; font-weight:600;
  color:#8a3b12; background:#fff4ec; border:1px solid #f0c7a8;
  padding:2px 6px; border-radius:4px;
}}
.kpis {{ display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:10px; margin-bottom:14px; }}
.kpis.three {{ grid-template-columns:repeat(3,minmax(0,1fr)); }}
.kpis > div {{ background:#f7fbfc; border:1px solid #d7e2e6; border-radius:8px; padding:12px 14px; }}
.kpis > div.highlight {{ background:#e8f6f8; border-color:#9ed7e0; }}
.label {{ font-size:11px; text-transform:uppercase; letter-spacing:.04em; color:#757575; }}
.value {{ font-size:22px; font-weight:700; margin-top:4px; color:var(--navy); font-variant-numeric:tabular-nums; }}
.muted {{ color:var(--muted); font-size:12px; margin-top:4px; }}
.table-scroll {{ overflow-x:auto; }}
table {{ width:100%; border-collapse:collapse; font-size:14px; }}
th, td {{ border-bottom:1px solid #e8e8e8; padding:9px 8px; text-align:left; vertical-align:middle; }}
th {{
  font-size:12px; text-transform:uppercase; letter-spacing:.03em; color:#fff;
  background:var(--navy); cursor:pointer; user-select:none; white-space:nowrap;
}}
th:hover {{ background:#03425a; }}
th .arrow {{ opacity:.45; margin-left:6px; font-size:11px; }}
th.sorted .arrow {{ opacity:1; }}
td.num {{ text-align:right; font-variant-numeric:tabular-nums; white-space:nowrap; }}
td.empty, .empty {{ color:#9a9a9a; }}
tbody tr:nth-child(even) {{ background:var(--zebra); }}
tbody tr:hover {{ background:#eef7f9; }}
tr.channel-jt td {{ background:#e8f2f5 !important; font-weight:700; }}
tr.channel-kt td {{ background:#f8ebe3 !important; font-weight:700; }}
.hint {{ margin-top:12px; color:var(--muted); font-size:12px; }}
.rules {{
  background:#e8f6f8; border:1px solid #9ed7e0; border-radius:8px;
  padding:12px 14px; margin-bottom:16px; font-size:13px; color:var(--navy);
}}
.note-box {{
  background:#fff8f2; border:1px solid #f0c7a8; border-radius:8px;
  padding:12px 14px; margin-top:12px; font-size:13px; color:#5a3a22;
}}
.note-box ul {{ margin:8px 0 0; padding-left:18px; }}
.note-box li {{ margin:4px 0; }}
@media (max-width:800px) {{
  .kpis, .kpis.three {{ grid-template-columns:repeat(2,minmax(0,1fr)); }}
}}
</style>
</head>
<body>
<div class="wrap">
  <div class="brand">FASTAIR</div>
  <h1>Выполненные заказы</h1>
  <div class="sub">Сроки поставки · оплата IBERIA / JET TECHNIC · транспорт план/факт · источник {html_escape(source_name)}</div>
  <div class="rules">
    <strong>Правила срока поставки</strong> (как в отчёте по поставкам):
    статус <strong>FINISHED</strong> · столбец S (Lead time) = <strong>только STK</strong>
    (заказы с числовым lead time не входят) · период по столбцу <strong>W</strong>
    · дни = W − Q · категории ROTABLE / EXPENDABLE.
  </div>
  {sections}
  <p class="hint">
    Попадание в период — по столбцу W (факт. дата поставки).
    Срок поставки: FINISHED + Lead time = STK (прочие lead time исключены), дни = W − Q.
    Срок оплаты (IBERIA / JET TECHNIC): подмножество тех же поставок — AW − Q без постоплаты и без оплаты раньше Q.
  </p>
</div>
<script>
(function () {{
  function cellValue(td, type) {{
    const raw = td.getAttribute('data-value');
    if (raw === null || raw === '') {{
      const t = (td.textContent || '').trim();
      if (type === 'num') {{
        const n = Number(t.replace(/\\s/g,'').replace('%','').replace('—',''));
        return Number.isFinite(n) ? n : Number.POSITIVE_INFINITY;
      }}
      return t.toLowerCase();
    }}
    return type === 'num' ? Number(raw) : String(raw);
  }}
  document.querySelectorAll('table.sortable').forEach(table => {{
    const tbody = table.tBodies[0];
    if (!tbody) return;
    const headers = [...table.tHead.rows[0].cells];
    let sortCol = -1;
    let asc = true;
    headers.forEach((th, idx) => {{
      th.dataset.col = String(idx);
      th.addEventListener('click', () => {{
        const col = Number(th.dataset.col);
        const type = th.dataset.type || 'str';
        if (sortCol === col) asc = !asc;
        else {{ sortCol = col; asc = type !== 'num'; }}
        headers.forEach((h, i) => {{
          h.classList.toggle('sorted', i === col);
          const arrow = h.querySelector('.arrow');
          if (arrow) arrow.textContent = i === col ? (asc ? '↑' : '↓') : '↕';
        }});
        const rows = [...tbody.rows];
        rows.sort((a, b) => {{
          const av = cellValue(a.cells[col], type);
          const bv = cellValue(b.cells[col], type);
          let cmp = type === 'num' ? (av - bv) : String(av).localeCompare(String(bv), 'ru');
          if (type === 'num') {{
            const aEmpty = !isFinite(av);
            const bEmpty = !isFinite(bv);
            if (aEmpty !== bEmpty) return aEmpty ? 1 : -1;
          }}
          return asc ? cmp : -cmp;
        }});
        rows.forEach(r => tbody.appendChild(r));
      }});
    }});
  }});
}})();
</script>
</body>
</html>
"""


def infer_taz_end(path: Path) -> date:
    """Prefer date embedded in filename like «ТАЗ 18.09.2026.xlsx»."""
    m = re.search(r"(\d{2})\.(\d{2})\.(\d{4})", path.name)
    if m:
        d, mo, y = map(int, m.groups())
        try:
            return date(y, mo, d)
        except ValueError:
            pass
    return date(2026, 9, 18)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--taz",
        type=Path,
        default=Path("/tmp/taz_history/ТАЗ 18.09.2026.xlsx"),
    )
    parser.add_argument("--out-dir", type=Path, default=Path("output"))
    parser.add_argument("--stem", default="completed_orders")
    args = parser.parse_args()

    end = infer_taz_end(args.taz)
    week_start = end - timedelta(days=6)

    df = prepare(load_taz(args.taz))
    periods = [
        build_period(df, "Прошедшая неделя", week_start, end),
        build_period(df, "01.06 — 18.09.2026", date(2026, 6, 1), date(2026, 9, 18)),
        build_period(df, "Весь 2026", date(2026, 1, 1), date(2026, 12, 31)),
    ]

    for p in periods:
        print(
            f"{p.title}: lead_n={p.lead_n} avg={p.lead_avg} "
            f"clients={len(p.clients)} suppliers={len(p.suppliers)} "
            f"transport plan={p.transport_total_plan:.0f} fact={p.transport_total_fact:.0f}"
        )

    html = render_html(periods, args.taz.name)
    args.out_dir.mkdir(parents=True, exist_ok=True)
    out_html = args.out_dir / f"{args.stem}.html"
    out_html.write_text(html, encoding="utf-8")
    with zipfile.ZipFile(args.out_dir / f"{args.stem}.zip", "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(f"{args.stem}.html", html)
    print(f"wrote {out_html}")


if __name__ == "__main__":
    main()
