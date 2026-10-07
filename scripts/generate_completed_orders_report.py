#!/usr/bin/env python3
"""FASTAIR — Выполненные заказы (пересобранная шапка KPI).

Периоды (по столбцу W — факт. дата поставки), дата среза = дата из имени ТАЗ:
  1) прошедшая неделя (7 дней, включая дату среза)
  2) прошедший полный месяц (например при срезе 02.10 → сентябрь)
  3) прошедший полный квартал (например при срезе 02.10 → Q3)
  4) прошедший год = текущий календарный год с 01.01 по дату среза

Попадание в отчёт: заполненная факт. дата поставки W; период по W.
Исключаем статусы CANCELLED / REFUND / WARRANTY / SCRAPPED / LOST и т.п.
Фильтр FINISHED не используем (это оформление отгрузочных документов, не факт выполнения).

Шапка периода:
  · выполнено заказов (уник. номер счёта)
  · выручка = Σ «Продажная, итого»
  · маржа = выручка − закупка − транспорт − transaction fee
  · сроки поставок факт: среднее W−Q, только STK;
    для IBERIA + JET TECHNIC дополнительно ср. срок оплаты AW−Q (STK, без постоплаты)
  · транспорт: план / факт / дельта

Под шапкой — выпадающие списки по клиентам и по поставщикам
(внутри — позиции с P/N + Description).
"""

from __future__ import annotations

import argparse
import calendar
import math
import re
from dataclasses import dataclass, field
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

COL_INVOICE = "Номер счета"
COL_STATUS = "Status"
COL_CUSTOMER = "Customer"
COL_CATEGORY = "Category"
COL_LEAD_TIME = "Lead time"
COL_ORDER_DATE = "ЗАКАЗ ВЗЯТ В РАБОТУ (ДАТА) ОТ КЛИЕНТА"
COL_DELIVERY = "ФАКТИЧЕСКАЯ ДАТА ПОСТАВКИ (СОГЛАСНО УСЛОВИЯМ ПОСТАВКИ)"
COL_SUPPLIER = "Поставщик"
COL_PN = "p/n"
COL_DESC = "DESCRIPTION"
COL_PAY_DATE = "Дата оплаты поставщику"
COL_MOVEMENT = "Дата начала движения"
COL_SALE = "Продажная, итого"
COL_BUY = "Закупка, итого"
COL_FEE = "Transaction fee"
COL_TRANSPORT_PLAN = (
    "Стоимость доставки ПЛАН, за весь счет! Если в счете несколько строк, "
    "то \"размазываем\" равномерно планируюмую стоиомость транспорта на все позиции из счета."
)
COL_TRANSPORT_FACT = "Стоимость доставки факт"
PAY_SUPPLIERS = {"IBERIA", "JET TECHNIC"}
# Не считаем выполнением: отмены, возвраты, гарантия, списание, потеря.
EXCLUDED_STATUS_TOKENS = (
    "CANCEL",
    "REFUND",
    "WARRANTY",
    "SCRAP",
    "LOST",
)

CAT_ROTABLE = "ROTABLE"
CAT_EXPENDABLE = "EXPENDABLE"
# Группы в выпадающих списках клиентов/поставщиков.
# GSE → ROTABLE; consumable → EXPENDABLE.
CAT_BUCKETS = (
    (CAT_ROTABLE, ("ROTABLE", "GSE")),
    (CAT_EXPENDABLE, ("EXPENDABLE", "CONSUMABLE")),
)
MONTHS_RU = {
    1: "январь",
    2: "февраль",
    3: "март",
    4: "апрель",
    5: "май",
    6: "июнь",
    7: "июль",
    8: "август",
    9: "сентябрь",
    10: "октябрь",
    11: "ноябрь",
    12: "декабрь",
}


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


def clean_str(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip()


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


def as_money(value) -> float:
    """Число для денег/маржи: NaN/None → 0 (иначе nan truthy ломает `x or 0`)."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return 0.0
    try:
        v = float(value)
    except (TypeError, ValueError):
        return 0.0
    if math.isnan(v) or math.isinf(v):
        return 0.0
    return v


def fmt_money(v: float) -> str:
    return f"{as_money(v):,.0f}".replace(",", " ")


def fmt_int(v: int) -> str:
    return f"{v:,}".replace(",", " ")


def load_taz(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=0)
    df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]
    if COL_TRANSPORT_PLAN not in df.columns:
        for c in df.columns:
            if isinstance(c, str) and c.startswith("Стоимость доставки ПЛАН"):
                df = df.rename(columns={c: COL_TRANSPORT_PLAN})
                break
    if COL_MOVEMENT not in df.columns:
        for c in df.columns:
            if isinstance(c, str) and c.startswith("Дата начала движения"):
                df = df.rename(columns={c: COL_MOVEMENT})
                break
    return df


def prepare(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["_status"] = out[COL_STATUS].map(
        lambda v: "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip().upper()
    )
    out["_excluded"] = out["_status"].map(
        lambda s: any(tok in s for tok in EXCLUDED_STATUS_TOKENS)
    )
    out["_lead"] = out[COL_LEAD_TIME].astype(str).str.strip().str.upper()
    out["_stk"] = out["_lead"].eq("STK")
    out["_invoice"] = out[COL_INVOICE].map(clean_str)
    out["_pn"] = out[COL_PN].map(clean_str) if COL_PN in out.columns else ""
    out["_desc"] = out[COL_DESC].map(clean_str) if COL_DESC in out.columns else ""
    out["_client"] = out[COL_CUSTOMER].fillna("").map(
        lambda v: "" if (isinstance(v, float) and pd.isna(v)) else str(v).strip()
    )
    out.loc[out["_client"].str.lower().isin({"", "nan", "none", "<na>"}), "_client"] = "— без клиента —"
    out["_supplier"] = out[COL_SUPPLIER].fillna("").map(
        lambda v: "" if (isinstance(v, float) and pd.isna(v)) else str(v).strip()
    )
    out.loc[
        out["_supplier"].eq("") | out["_supplier"].str.lower().isin({"nan", "none", "<na>"}),
        "_supplier",
    ] = "— без поставщика —"
    out["_cat"] = [
        ""
        if v is None or (isinstance(v, float) and pd.isna(v)) or (isinstance(v, str) and not str(v).strip())
        else str(v).strip().upper()
        for v in out[COL_CATEGORY]
    ]
    out["_bucket"] = [category_bucket(c) for c in out["_cat"]]
    out["_q"] = out[COL_ORDER_DATE].map(parse_date)
    out["_w"] = out[COL_DELIVERY].map(parse_date)
    out["_aw"] = out[COL_PAY_DATE].map(parse_date) if COL_PAY_DATE in out.columns else None
    out["_ba"] = out[COL_MOVEMENT].map(parse_date) if COL_MOVEMENT in out.columns else None
    out["_days"] = [(w - q).days if w and q else None for w, q in zip(out["_w"], out["_q"])]
    out["_pay_days"] = [
        (aw - q).days if aw and q else None for aw, q in zip(out["_aw"], out["_q"])
    ]
    out["_sale"] = out[COL_SALE].map(parse_numeric) if COL_SALE in out.columns else 0.0
    out["_buy"] = out[COL_BUY].map(parse_numeric) if COL_BUY in out.columns else 0.0
    out["_fee"] = out[COL_FEE].map(parse_numeric) if COL_FEE in out.columns else 0.0
    plan_col = COL_TRANSPORT_PLAN if COL_TRANSPORT_PLAN in out.columns else None
    fact_col = COL_TRANSPORT_FACT if COL_TRANSPORT_FACT in out.columns else None
    out["_plan"] = out[plan_col].map(parse_numeric) if plan_col else 0.0
    if fact_col:
        # Пустой факт → берём план. Нельзя возвращать None через Series.map:
        # pandas превращает None в NaN, и тогда fallback ломается.
        def _fact_or_plan(v, plan: float) -> float:
            if v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == "":
                return float(plan)
            return parse_numeric(v)

        out["_fact"] = [
            _fact_or_plan(v, p) for v, p in zip(out[fact_col], out["_plan"])
        ]
    else:
        out["_fact"] = out["_plan"]
    out["_margin"] = out["_sale"] - out["_buy"] - out["_fact"] - out["_fee"]
    return out


def in_period(d: date | None, start: date, end: date) -> bool:
    return d is not None and start <= d <= end


def report_mask(df: pd.DataFrame, start: date, end: date) -> pd.Series:
    """Попадание: заполненная W в периоде; без cancel/refund/warranty/scrap/lost."""
    return (
        df["_w"].map(lambda d: in_period(d, start, end))
        & ~df["_excluded"]
    )


def lead_mask(df: pd.DataFrame, start: date, end: date) -> pd.Series:
    return (
        report_mask(df, start, end)
        & df["_stk"]
        & df["_q"].notna()
        & df["_days"].notna()
        & (df["_days"] >= 0)
        & df["_bucket"].isin({CAT_ROTABLE, CAT_EXPENDABLE})
    )


def pay_mask(df: pd.DataFrame, start: date, end: date) -> pd.Series:
    """IBERIA/JET TECHNIC: STK, AW−Q≥0, без постоплаты (движение раньше оплаты)."""
    suppliers = df["_supplier"].astype(str).str.strip().str.upper()
    ba = df["_ba"]
    aw = df["_aw"]
    postpay = ba.notna() & (aw.isna() | (ba < aw))
    return (
        lead_mask(df, start, end)
        & suppliers.isin(PAY_SUPPLIERS)
        & aw.notna()
        & df["_pay_days"].notna()
        & (df["_pay_days"] >= 0)
        & ~postpay
    )


def infer_taz_end(path: Path) -> date:
    m = re.search(r"(\d{2})\.(\d{2})\.(\d{4})", path.name)
    if m:
        d, mo, y = map(int, m.groups())
        try:
            return date(y, mo, d)
        except ValueError:
            pass
    return date.today()


def last_complete_month(as_of: date) -> tuple[date, date, str]:
    first_this = as_of.replace(day=1)
    last_prev = first_this - timedelta(days=1)
    first_prev = last_prev.replace(day=1)
    title = f"{MONTHS_RU[first_prev.month].capitalize()} {first_prev.year}"
    return first_prev, last_prev, title


def last_complete_quarter(as_of: date) -> tuple[date, date, str]:
    q = (as_of.month - 1) // 3 + 1
    if q == 1:
        year = as_of.year - 1
        start_m = 10
        q_label = 4
    else:
        year = as_of.year
        start_m = (q - 2) * 3 + 1
        q_label = q - 1
    start = date(year, start_m, 1)
    end_m = start_m + 2
    end = date(year, end_m, calendar.monthrange(year, end_m)[1])
    title = f"Q{q_label} {year} ({MONTHS_RU[start_m]}–{MONTHS_RU[end_m]})"
    return start, end, title


def category_bucket(cat) -> str:
    """ROTABLE (+GSE), EXPENDABLE (+consumable); остальное → Прочее."""
    if cat is None or (isinstance(cat, float) and pd.isna(cat)):
        return "Прочее"
    text = str(cat).strip().upper()
    if text in {"", "NAN", "NONE", "<NA>"}:
        return "Прочее"
    for bucket, tokens in CAT_BUCKETS:
        if any(tok in text for tok in tokens):
            return bucket
    return "Прочее"


@dataclass
class LineRow:
    pn: str
    description: str
    revenue: float
    margin: float
    lead_days: float | None
    is_stk: bool
    transport_plan: float
    transport_fact: float
    invoice: str

    @property
    def transport_delta(self) -> float:
        return self.transport_fact - self.transport_plan


@dataclass
class CategoryRow:
    name: str
    orders: int
    lines: int
    revenue: float
    margin: float
    lead_avg: float | None
    lead_n: int
    positions: list[LineRow] = field(default_factory=list)


@dataclass
class EntityRow:
    name: str
    orders: int
    lines: int
    revenue: float
    margin: float
    lead_avg: float | None
    lead_n: int
    transport_plan: float
    transport_fact: float
    categories: list[CategoryRow] = field(default_factory=list)

    @property
    def transport_delta(self) -> float:
        return self.transport_fact - self.transport_plan


@dataclass
class PeriodBlock:
    title: str
    start: date
    end: date
    orders: int
    lines: int
    revenue: float
    margin: float
    lead_avg: float | None
    lead_n: int
    pay_avg: float | None
    pay_n: int
    transport_plan: float
    transport_fact: float
    clients: list[EntityRow] = field(default_factory=list)
    suppliers: list[EntityRow] = field(default_factory=list)

    @property
    def transport_delta(self) -> float:
        return self.transport_fact - self.transport_plan


def line_from_row(row: pd.Series) -> LineRow:
    days = row["_days"]
    lead_days = None
    if days is not None and not (isinstance(days, float) and pd.isna(days)):
        lead_days = float(days)
    return LineRow(
        pn=str(row["_pn"] or "—"),
        description=str(row["_desc"] or "—"),
        revenue=as_money(row["_sale"]),
        margin=as_money(row["_margin"]),
        lead_days=lead_days,
        is_stk=bool(row["_stk"]),
        transport_plan=as_money(row["_plan"]),
        transport_fact=as_money(row["_fact"]),
        invoice=str(row["_invoice"] or ""),
    )


def _summarize_part(part: pd.DataFrame, lead_part: pd.DataFrame) -> tuple[int, int, float, float, float | None, int, list[LineRow]]:
    days = lead_part["_days"].astype(float) if not lead_part.empty else pd.Series(dtype=float)
    invoices = {inv for inv in part["_invoice"] if inv}
    positions = [line_from_row(r) for _, r in part.sort_values("_sale", ascending=False).iterrows()]
    return (
        len(invoices) if invoices else int(len(part)),
        int(len(part)),
        as_money(part["_sale"].sum()),
        as_money(part["_margin"].sum()),
        float(days.mean()) if not days.empty else None,
        int(len(days)),
        positions,
    )


def summarize_group(rows: pd.DataFrame, lead_rows: pd.DataFrame, key: str) -> list[EntityRow]:
    names = sorted(rows[key].unique(), key=lambda s: str(s).casefold())
    out: list[EntityRow] = []
    bucket_order = [CAT_ROTABLE, CAT_EXPENDABLE, "Прочее"]
    for name in names:
        part = rows[rows[key] == name]
        lead_part = lead_rows[lead_rows[key] == name]
        orders, lines, revenue, margin, lead_avg, lead_n, _ = _summarize_part(part, lead_part)

        categories: list[CategoryRow] = []
        for bucket in bucket_order:
            b_part = part[part["_bucket"] == bucket]
            if b_part.empty:
                continue
            b_lead = lead_part[lead_part["_bucket"] == bucket]
            b_orders, b_lines, b_rev, b_margin, b_lead_avg, b_lead_n, positions = _summarize_part(
                b_part, b_lead
            )
            categories.append(
                CategoryRow(
                    name=bucket,
                    orders=b_orders,
                    lines=b_lines,
                    revenue=b_rev,
                    margin=b_margin,
                    lead_avg=b_lead_avg,
                    lead_n=b_lead_n,
                    positions=positions,
                )
            )

        out.append(
            EntityRow(
                name=str(name),
                orders=orders,
                lines=lines,
                revenue=revenue,
                margin=margin,
                lead_avg=lead_avg,
                lead_n=lead_n,
                transport_plan=as_money(part["_plan"].sum()),
                transport_fact=as_money(part["_fact"].sum()),
                categories=categories,
            )
        )
    out.sort(key=lambda e: (-e.revenue, -e.orders, e.name.casefold()))
    return out


def build_period(df: pd.DataFrame, title: str, start: date, end: date) -> PeriodBlock:
    rows = df.loc[report_mask(df, start, end)].copy()
    lead_rows = df.loc[lead_mask(df, start, end)].copy()
    pay_rows = df.loc[pay_mask(df, start, end)].copy()
    invoices = {inv for inv in rows["_invoice"] if inv}
    days = lead_rows["_days"].astype(float)
    pay_days = pay_rows["_pay_days"].astype(float)
    return PeriodBlock(
        title=title,
        start=start,
        end=end,
        orders=len(invoices) if invoices else int(len(rows)),
        lines=int(len(rows)),
        revenue=as_money(rows["_sale"].sum()),
        margin=as_money(rows["_margin"].sum()),
        lead_avg=float(days.mean()) if not days.empty else None,
        lead_n=int(len(days)),
        pay_avg=float(pay_days.mean()) if not pay_days.empty else None,
        pay_n=int(len(pay_days)),
        transport_plan=as_money(rows["_plan"].sum()),
        transport_fact=as_money(rows["_fact"].sum()),
        clients=summarize_group(rows, lead_rows, "_client"),
        suppliers=summarize_group(rows, lead_rows, "_supplier"),
    )


def render_positions_table(positions: list[LineRow], table_id: str) -> str:
    body = []
    for r in positions:
        delta = r.transport_delta
        delta_cls = "bad" if delta > 0 else "ok"
        body.append(
            "<tr>"
            f"<td>{html_escape(r.pn)}</td>"
            f"<td>{html_escape(r.description)}</td>"
            f"<td class='num'>{fmt_money(r.revenue)}</td>"
            f"<td class='num'>{fmt_money(r.margin)}</td>"
            f"<td class='num'>{fmt_days(r.lead_days) if r.is_stk else '—'}</td>"
            f"<td class='num {delta_cls}'>{fmt_money(delta)}</td>"
            f"<td class='num'>{html_escape(r.invoice or '—')}</td>"
            "</tr>"
        )
    if not body:
        body.append("<tr><td colspan='7' class='empty'>Нет позиций</td></tr>")
    return f"""
    <div class="table-scroll">
    <table data-sortable id="{html_escape(table_id)}">
      <thead><tr>
        <th class="label-col">P/N <span class="arrow">↕</span></th>
        <th class="label-col">Description <span class="arrow">↕</span></th>
        <th class="num">Выручка $ <span class="arrow">↕</span></th>
        <th class="num">Маржа $ <span class="arrow">↕</span></th>
        <th class="num">Срок дн. <span class="arrow">↕</span></th>
        <th class="num">Δ тр. $ <span class="arrow">↕</span></th>
        <th class="num">Счёт <span class="arrow">↕</span></th>
      </tr></thead>
      <tbody>{''.join(body)}</tbody>
    </table>
    </div>
    """


def render_entity_dropdowns(entities: list[EntityRow], table_prefix: str) -> str:
    blocks = []
    for i, ent in enumerate(entities):
        cat_blocks = []
        for j, cat in enumerate(ent.categories):
            hint = ""
            if cat.name == CAT_ROTABLE:
                hint = " · вкл. GSE"
            elif cat.name == CAT_EXPENDABLE:
                hint = " · вкл. consumable"
            cat_blocks.append(
                f"""
<details class="catblock" open>
  <summary>
    <span class="cat-name">{html_escape(cat.name)}{html_escape(hint)}</span>
    <span class="ent-meta">{fmt_int(cat.orders)} зак. · {fmt_int(cat.lines)} поз. · ${fmt_money(cat.revenue)} · маржа ${fmt_money(cat.margin)} · срок {fmt_days(cat.lead_avg)} дн.</span>
  </summary>
  <div class="block-body">
    {render_positions_table(cat.positions, f"{table_prefix}-{i}-{j}")}
  </div>
</details>
"""
            )
        inner = "\n".join(cat_blocks) if cat_blocks else "<p class='empty'>Нет позиций</p>"
        blocks.append(
            f"""
<details class="subblock">
  <summary>
    <span class="ent-name">{html_escape(ent.name)}</span>
    <span class="ent-meta">{fmt_int(ent.orders)} зак. · {fmt_int(ent.lines)} поз. · ${fmt_money(ent.revenue)} · маржа ${fmt_money(ent.margin)} · срок {fmt_days(ent.lead_avg)} дн.</span>
  </summary>
  <div class="block-body">
    {inner}
  </div>
</details>
"""
        )
    if not blocks:
        return "<p class='empty'>Нет данных</p>"
    return "\n".join(blocks)


def render_period(period: PeriodBlock, idx: int, *, opened: bool) -> str:
    open_attr = " open" if opened else ""
    range_s = f"{period.start.strftime('%d.%m.%Y')} – {period.end.strftime('%d.%m.%Y')}"
    pay_line = (
        f"оплата IBERIA+JT: {fmt_days(period.pay_avg)} дн. · {fmt_int(period.pay_n)} поз."
        if period.pay_n
        else "оплата IBERIA+JT: нет данных"
    )
    return f"""
<details class="period"{open_attr}>
  <summary>
    <span class="period-title">{idx}. {html_escape(period.title)}</span>
    <span class="period-meta">{range_s}</span>
    <span class="period-meta">{fmt_int(period.orders)} зак. · ${fmt_money(period.revenue)} выр. · маржа ${fmt_money(period.margin)}</span>
  </summary>
  <div class="period-body">
    <div class="kpis">
      <div class="highlight">
        <div class="label">Выполнено заказов</div>
        <div class="value">{fmt_int(period.orders)}</div>
        <div class="muted">{fmt_int(period.lines)} позиций с датой W</div>
      </div>
      <div>
        <div class="label">Выручка общ.</div>
        <div class="value">${fmt_money(period.revenue)}</div>
        <div class="muted">Σ продажная итого</div>
      </div>
      <div>
        <div class="label">Маржа общ.</div>
        <div class="value">${fmt_money(period.margin)}</div>
        <div class="muted">продажа − закупка − тр. − fee</div>
      </div>
      <div>
        <div class="label">Сроки поставок факт</div>
        <div class="value">{fmt_days(period.lead_avg)} дн.</div>
        <div class="muted">ср. W−Q · только STK · {fmt_int(period.lead_n)} поз.</div>
        <div class="muted pay">{html_escape(pay_line)}</div>
      </div>
      <div class="wide">
        <div class="label">Транспорт план − факт = дельта</div>
        <div class="value transport">
          <span>{fmt_money(period.transport_plan)}</span>
          <span class="op">−</span>
          <span>{fmt_money(period.transport_fact)}</span>
          <span class="op">=</span>
          <span class="{'bad' if period.transport_delta > 0 else 'ok'}">{fmt_money(period.transport_delta)}</span>
        </div>
        <div class="muted">USD · по всем позициям с W</div>
      </div>
    </div>

    <details class="block" open>
      <summary>По клиентам ({len(period.clients)})</summary>
      <div class="block-body">{render_entity_dropdowns(period.clients, f"c{idx}")}</div>
    </details>
    <details class="block" open>
      <summary>По поставщикам ({len(period.suppliers)})</summary>
      <div class="block-body">{render_entity_dropdowns(period.suppliers, f"s{idx}")}</div>
    </details>
  </div>
</details>
"""


def render_html(periods: list[PeriodBlock], source_name: str, as_of: date) -> str:
    sections = "\n".join(
        render_period(p, i + 1, opened=(i == 0)) for i, p in enumerate(periods)
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
}}
* {{ box-sizing:border-box; }}
body {{
  margin:0; font-family: Arial, Calibri, Helvetica, sans-serif;
  color:var(--ink); background: linear-gradient(180deg, #022f40 0 140px, var(--bg) 140px);
  line-height:1.45;
}}
.wrap {{ max-width:1180px; margin:0 auto; padding:28px 20px 64px; }}
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
details.block, details.subblock, details.catblock {{
  border:1px solid #d7e2e6; border-radius:8px; margin:10px 0; background:#fff;
}}
details.catblock {{
  margin:8px 0; border-color:#c5d4da; background:#fafcfd;
}}
details.block > summary, details.subblock > summary, details.catblock > summary {{
  cursor:pointer; padding:10px 12px; font-weight:700; color:var(--navy);
  list-style:none; background:#f7fbfc;
  display:flex; flex-wrap:wrap; gap:8px 14px; align-items:baseline;
}}
details.catblock > summary {{
  background:#eef5f7; font-size:13px;
}}
details.block > summary::-webkit-details-marker,
details.subblock > summary::-webkit-details-marker,
details.catblock > summary::-webkit-details-marker {{ display:none; }}
details.block > summary::before,
details.subblock > summary::before,
details.catblock > summary::before {{ content:"▸ "; color:#7a8a90; }}
details.block[open] > summary::before,
details.subblock[open] > summary::before,
details.catblock[open] > summary::before {{ content:"▾ "; }}
.block-body {{ padding:12px; }}
.ent-name, .cat-name {{ font-weight:700; }}
.ent-meta {{ font-size:12px; color:var(--muted); font-weight:600; }}
.kpis {{
  display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:10px; margin-bottom:14px;
}}
.kpis > div {{ background:#f7fbfc; border:1px solid #d7e2e6; border-radius:8px; padding:12px 14px; }}
.kpis > div.highlight {{ background:#e8f6f8; border-color:#9ed7e0; }}
.kpis > div.wide {{ grid-column:1 / -1; }}
.label {{ font-size:11px; text-transform:uppercase; letter-spacing:.04em; color:#757575; }}
.value {{ font-size:22px; font-weight:700; margin-top:4px; color:var(--navy); font-variant-numeric:tabular-nums; }}
.value.transport {{ display:flex; flex-wrap:wrap; gap:8px 12px; align-items:baseline; font-size:20px; }}
.value.transport .op {{ color:var(--muted); font-weight:600; }}
.bad {{ color:#9b2c2c; font-weight:700; }}
.ok {{ color:#0a5c4c; font-weight:700; }}
.muted {{ color:var(--muted); font-size:12px; margin-top:4px; }}
.muted.pay {{ color:#0a5c4c; font-weight:600; }}
.table-scroll {{ overflow-x:auto; }}
table {{
  width:100%; border-collapse:collapse; font-size:13px;
  border:1px solid #b7c8cf;
}}
th, td {{
  border:1px solid #c5d4da; padding:9px 10px; text-align:left; vertical-align:middle;
}}
th {{
  font-size:11px; text-transform:uppercase; letter-spacing:.03em; color:#fff;
  background:var(--navy); cursor:pointer; user-select:none; white-space:nowrap;
  text-align:center;
}}
th.label-col {{ text-align:left; }}
th:hover {{ background:#03425a; }}
th .arrow {{ opacity:.45; margin-left:6px; font-size:11px; }}
th.sorted .arrow {{ opacity:1; }}
td.num, th.num {{
  text-align:center; font-variant-numeric:tabular-nums; white-space:nowrap;
}}
td.empty, .empty {{ color:#9a9a9a; text-align:center; }}
tbody tr:nth-child(even) {{ background:var(--zebra); }}
tbody tr:hover {{ background:#eef7f9; }}
.hint {{ margin-top:12px; color:var(--muted); font-size:12px; }}
.rules {{
  background:#e8f6f8; border:1px solid #9ed7e0; border-radius:8px;
  padding:12px 14px; margin-bottom:16px; font-size:13px; color:var(--navy);
}}
@media (max-width:900px) {{
  .kpis {{ grid-template-columns:repeat(2,minmax(0,1fr)); }}
}}
</style>
</head>
<body>
<div class="wrap">
  <div class="brand">FASTAIR</div>
  <h1>Выполненные заказы</h1>
  <div class="sub">Срез ТАЗ {html_escape(as_of.strftime('%d.%m.%Y'))} · источник {html_escape(source_name)}</div>
  <div class="rules">
    <strong>Попадание:</strong> заполненная факт. дата поставки <strong>W</strong>; период по W.
    Исключаем статусы <strong>CANCELLED / REFUND / WARRANTY / SCRAPPED / LOST</strong>.
    Фильтр FINISHED не используем (это оформление отгрузочных документов, не факт выполнения).
    <br/>
    <strong>Периоды:</strong> 1) прошедшая неделя · 2) последний <em>полный</em> месяц ·
    3) последний <em>полный</em> квартал · 4) текущий год с 01.01 по дату среза.
    <br/>
    <strong>Сроки:</strong> среднее W−Q, только <strong>STK</strong> (лид-таймы исключены).
    Для <strong>IBERIA / JET TECHNIC</strong> в шапке дополнительно ср. срок оплаты AW−Q (STK, без постоплаты).
    <br/>
    <strong>Маржа:</strong> продажная итого − закупка итого − транспорт (факт, иначе план) − transaction fee.
    В списках клиентов/поставщиков позиции разбиты на <strong>ROTABLE</strong> (вкл. GSE)
    и <strong>EXPENDABLE</strong> (вкл. consumable); по транспорту только Δ (красно/зелёно).
  </div>
  {sections}
  <p class="hint">Скачивайте / открывайте HTML напрямую. ZIP+HTML Windows часто помечает ложно.</p>
</div>
<script>
document.querySelectorAll('table[data-sortable]').forEach((table) => {{
  const headers = table.querySelectorAll('th');
  headers.forEach((th, idx) => {{
    th.addEventListener('click', () => {{
      const tbody = table.tBodies[0];
      if (!tbody) return;
      const rows = Array.from(tbody.querySelectorAll('tr'));
      const asc = th.dataset.asc !== '1';
      headers.forEach(h => {{ h.classList.remove('sorted'); h.dataset.asc = ''; }});
      th.classList.add('sorted');
      th.dataset.asc = asc ? '1' : '0';
      const parse = (td) => {{
        const t = (td?.textContent || '').trim().replace(/\\s/g, '').replace('%','').replace('$','');
        const n = Number(t.replace(',', '.'));
        return Number.isFinite(n) && t !== '' && t !== '—' ? n : t;
      }};
      rows.sort((a, b) => {{
        const av = parse(a.children[idx]);
        const bv = parse(b.children[idx]);
        if (typeof av === 'number' && typeof bv === 'number') return asc ? av - bv : bv - av;
        return asc ? String(av).localeCompare(String(bv), 'ru') : String(bv).localeCompare(String(av), 'ru');
      }});
      rows.forEach(r => tbody.appendChild(r));
    }});
  }});
}});
</script>
</body>
</html>
"""


def build_periods(df: pd.DataFrame, as_of: date) -> list[PeriodBlock]:
    week_start = as_of - timedelta(days=6)
    month_start, month_end, month_title = last_complete_month(as_of)
    q_start, q_end, q_title = last_complete_quarter(as_of)
    year_start = date(as_of.year, 1, 1)
    return [
        build_period(df, "Прошедшая неделя", week_start, as_of),
        build_period(df, f"Прошедший месяц · {month_title}", month_start, month_end),
        build_period(df, f"Прошедший квартал · {q_title}", q_start, q_end),
        build_period(df, f"Прошедший год · {as_of.year}", year_start, as_of),
    ]


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--taz",
        type=Path,
        default=Path("/tmp/taz_history/ТАЗ 02.10.2026.xlsx"),
    )
    parser.add_argument("--out-dir", type=Path, default=Path("output"))
    parser.add_argument("--stem", default="completed_orders")
    args = parser.parse_args()

    as_of = infer_taz_end(args.taz)
    df = prepare(load_taz(args.taz))
    periods = build_periods(df, as_of)

    for p in periods:
        print(
            f"{p.title}: orders={p.orders} lines={p.lines} "
            f"rev={p.revenue:.0f} margin={p.margin:.0f} "
            f"lead_avg={p.lead_avg} lead_n={p.lead_n} "
            f"pay_avg={p.pay_avg} pay_n={p.pay_n} "
            f"tr_delta={p.transport_delta:.0f}"
        )

    html = render_html(periods, args.taz.name, as_of)
    args.out_dir.mkdir(parents=True, exist_ok=True)
    out_html = args.out_dir / f"{args.stem}.html"
    out_html.write_text(html, encoding="utf-8")
    print(f"wrote {out_html}")


if __name__ == "__main__":
    main()
