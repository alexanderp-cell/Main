#!/usr/bin/env python3
"""TIME TO PAY — supplier share + JT/KT payment-speed HTML report."""

from __future__ import annotations

import argparse
import math
import re
import zipfile
from dataclasses import dataclass
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

COL_STATUS = "Status"
COL_SUPPLIER = "Поставщик"
COL_ORDER_DATE = "ЗАКАЗ ВЗЯТ В РАБОТУ (ДАТА) ОТ КЛИЕНТА"
COL_SALE = "Продажная, итого"
COL_PURCHASE = "Закупка, итого"
COL_LEAD_TIME = "Lead time"
COL_PAY_DATE = "Дата оплаты поставщику"
COL_PAY_AMOUNT = "Сумма оплаты поставщику"
COL_MOVEMENT = "Дата начала движения"

JT_LABEL = "JET TECHNIC"
KT_LABEL = "KT MNT"


def parse_numeric(value) -> float:
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


def parse_date(value) -> date | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    ts = pd.to_datetime(value, errors="coerce", dayfirst=True)
    if pd.isna(ts):
        return None
    return ts.date()


def parse_lead_days(value) -> int | None:
    """Return 0 for STK, positive int for numeric LT, None if unknown."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if math.isnan(float(value)):
            return None
        return max(0, int(round(float(value))))
    text = str(value).strip().upper().replace("\xa0", " ")
    if not text or text in {"NAN", "NONE", "NAT", "TBA", "?", "???"}:
        return None
    if "STK" in text:
        return 0
    # dates in Lead time are not usable as days
    if re.match(r"\d{4}-\d{2}-\d{2}", text):
        return None
    m = re.search(r"(\d+)", text)
    return int(m.group(1)) if m else None


def classify_channel(name: str) -> str | None:
    u = re.sub(r"\s+", " ", str(name or "").strip().upper())
    if "JET TECHNIC" in u or u in {"JT", "JETTECHNIC", "JET-TECHNIC"}:
        return "JT"
    if (
        "KT MNT" in u
        or "KT MAINTENANCE" in u
        or "KT MAINTEN" in u
        or u in {"KT", "KTMNT", "KT-MNT", "KT_MNT"}
    ):
        return "KT"
    return None


def load_taz(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=0)
    df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]
    # movement column may have trailing spaces in older exports
    for c in list(df.columns):
        if isinstance(c, str) and c.startswith("Дата начала движения"):
            if c != COL_MOVEMENT:
                df = df.rename(columns={c: COL_MOVEMENT})
            break
    return df


def prepare(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["_status"] = out[COL_STATUS].astype(str).str.strip().str.upper()
    out["_supplier"] = (
        out[COL_SUPPLIER]
        .fillna("")
        .map(lambda v: "" if (isinstance(v, float) and pd.isna(v)) else str(v).strip())
    )
    out.loc[
        out["_supplier"].eq("") | out["_supplier"].str.lower().isin({"nan", "none", "<na>"}),
        "_supplier",
    ] = "— без поставщика —"
    out["_channel"] = out["_supplier"].map(classify_channel)
    out["_q"] = out[COL_ORDER_DATE].map(parse_date)
    out["_aw"] = out[COL_PAY_DATE].map(parse_date)
    if COL_MOVEMENT in out.columns:
        out["_ba"] = out[COL_MOVEMENT].map(parse_date)
    else:
        out["_ba"] = None
    out["_sale"] = out[COL_SALE].map(parse_numeric)
    out["_purchase"] = out[COL_PURCHASE].map(parse_numeric)
    out["_margin"] = out["_sale"] - out["_purchase"]
    out["_ax"] = out[COL_PAY_AMOUNT].map(parse_numeric)
    out["_lt_days"] = out[COL_LEAD_TIME].map(parse_lead_days)
    out["_days_pay"] = [
        (aw - q).days if aw and q else None for aw, q in zip(out["_aw"], out["_q"])
    ]
    # drop cancelled-like
    bad = out["_status"].str.contains("CANCELLED|REFUND|WARRANTY", na=False)
    return out.loc[~bad].copy()


@dataclass
class SupplierShare:
    name: str
    channel: str | None
    revenue: float
    revenue_pct: float
    margin: float
    margin_pct: float
    n: int


@dataclass
class PaySpeed:
    label: str
    mean_days: float | None
    median_days: float | None
    n: int
    excluded_postpay: int
    excluded_lead: int


@dataclass
class DayPay:
    day: date
    jt: float
    kt: float


@dataclass
class PeriodReport:
    title: str
    start: date
    end: date
    total_revenue: float
    total_margin: float
    total_n: int
    shares: list[SupplierShare]
    jt_speed: PaySpeed
    kt_speed: PaySpeed
    both_speed: PaySpeed
    payments: list[DayPay]
    kt_found: bool


def in_period(d: date | None, start: date, end: date) -> bool:
    return d is not None and start <= d <= end


def build_shares(rows: pd.DataFrame) -> tuple[list[SupplierShare], float, float, int]:
    total_rev = float(rows["_sale"].sum())
    total_m = float(rows["_margin"].sum())
    total_n = int(len(rows))
    grouped = (
        rows.groupby("_supplier", dropna=False)
        .agg(revenue=("_sale", "sum"), margin=("_margin", "sum"), n=("_sale", "size"))
        .reset_index()
        .sort_values("revenue", ascending=False)
    )
    shares: list[SupplierShare] = []
    seen_channels = set()
    for _, rec in grouped.iterrows():
        name = str(rec["_supplier"]).strip()
        if name.lower() in {"nan", "none", "", "<na>"}:
            name = "— без поставщика —"
        ch = classify_channel(name)
        if ch:
            seen_channels.add(ch)
        shares.append(
            SupplierShare(
                name=name,
                channel=ch,
                revenue=float(rec["revenue"]),
                revenue_pct=(float(rec["revenue"]) / total_rev * 100.0) if total_rev else 0.0,
                margin=float(rec["margin"]),
                margin_pct=(float(rec["margin"]) / total_m * 100.0) if total_m else 0.0,
                n=int(rec["n"]),
            )
        )
    # ensure KT row exists even if absent in TAZ
    if "KT" not in seen_channels:
        shares.append(
            SupplierShare(
                name=KT_LABEL,
                channel="KT",
                revenue=0.0,
                revenue_pct=0.0,
                margin=0.0,
                margin_pct=0.0,
                n=0,
            )
        )
    # keep JT/KT near top after majors: sort by revenue but pin channels after sort
    shares.sort(key=lambda s: (-s.revenue, s.name.casefold()))
    return shares, total_rev, total_m, total_n


def is_postpay(row) -> bool:
    """Shipped/moved before supplier was paid (or moved and still unpaid)."""
    ba = row["_ba"]
    aw = row["_aw"]
    if ba is None:
        return False
    if aw is None:
        return True
    return ba < aw


def is_lead_time_order(row) -> bool:
    """Non-stock lead time: payment after manufacturing LT is expected."""
    lt = row["_lt_days"]
    return lt is not None and lt > 0


def build_speed(rows: pd.DataFrame, channel: str | None) -> PaySpeed:
    pool = rows if channel is None else rows[rows["_channel"] == channel]
    if pool.empty:
        label = {"JT": JT_LABEL, "KT": KT_LABEL, None: "JT + KT"}[channel]
        return PaySpeed(label, None, None, 0, 0, 0)

    ba = pool["_ba"]
    aw = pool["_aw"]
    postpay_mask = ba.notna() & (aw.isna() | (ba < aw))
    lead_mask = pool["_lt_days"].notna() & (pool["_lt_days"] > 0)
    excluded_postpay = int(postpay_mask.sum())
    excluded_lead = int(lead_mask.sum())

    usable = pool[
        aw.notna()
        & pool["_q"].notna()
        & pool["_days_pay"].notna()
        & (pool["_days_pay"] >= 0)
        & ~postpay_mask
        & (pool["_lt_days"] == 0)  # STK only; also drops unknown LT
    ]
    days = usable["_days_pay"].astype(float)
    label = {"JT": JT_LABEL, "KT": KT_LABEL, None: "JT + KT"}[channel]
    if days.empty:
        return PaySpeed(label, None, None, 0, excluded_postpay, excluded_lead)
    return PaySpeed(
        label=label,
        mean_days=float(days.mean()),
        median_days=float(days.median()),
        n=int(len(days)),
        excluded_postpay=excluded_postpay,
        excluded_lead=excluded_lead,
    )


def build_payments(df: pd.DataFrame, start: date, end: date) -> list[DayPay]:
    """Daily supplier payment amounts for JT/KT by AW date in period."""
    part = df[
        df["_channel"].isin(["JT", "KT"])
        & df["_aw"].notna()
        & df["_ax"].notna()
        & (df["_ax"] > 0)
    ].copy()
    part = part[part["_aw"].map(lambda d: in_period(d, start, end))]
    by_day: dict[date, DayPay] = {}
    for _, row in part.iterrows():
        day = row["_aw"]
        bucket = by_day.get(day) or DayPay(day=day, jt=0.0, kt=0.0)
        if row["_channel"] == "JT":
            bucket.jt += float(row["_ax"])
        else:
            bucket.kt += float(row["_ax"])
        by_day[day] = bucket
    # fill all calendar days for continuous weekend visibility
    if start > end:
        return []
    out: list[DayPay] = []
    cur = start
    while cur <= end:
        out.append(by_day.get(cur) or DayPay(day=cur, jt=0.0, kt=0.0))
        cur += timedelta(days=1)
    return out


def build_period(df: pd.DataFrame, title: str, start: date, end: date) -> PeriodReport:
    q_ok = df["_q"].notna() & df["_q"].map(lambda d: start <= d <= end)
    by_q = df.loc[q_ok].copy()
    shares, total_rev, total_m, total_n = build_shares(by_q)
    channels = by_q[by_q["_channel"].isin(["JT", "KT"])].copy()
    kt_found = bool((df["_channel"] == "KT").any())
    return PeriodReport(
        title=title,
        start=start,
        end=end,
        total_revenue=total_rev,
        total_margin=total_m,
        total_n=total_n,
        shares=shares,
        jt_speed=build_speed(channels, "JT"),
        kt_speed=build_speed(channels, "KT"),
        both_speed=build_speed(channels, None),
        payments=build_payments(df, start, end),
        kt_found=kt_found,
    )


def fmt_money(v: float) -> str:
    return f"{v:,.0f}".replace(",", " ")


def fmt_pct(v: float) -> str:
    return f"{v:.1f}%"


def fmt_days(v: float | None) -> str:
    if v is None:
        return "—"
    return f"{v:.1f}"


def html_escape(text: str) -> str:
    return (
        str(text)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def render_histogram(payments: list[DayPay], chart_id: str) -> str:
    if not payments:
        return '<p class="hint">Нет платежей JT/KT за период.</p>'

    max_amt = max((p.jt + p.kt) for p in payments) or 1.0
    # SVG geometry
    bar_w = 6
    gap = 1
    left = 48
    top = 16
    height = 180
    width = left + len(payments) * (bar_w + gap) + 16
    plot_h = height - 36

    bars = []
    weekend_bg = []
    for i, p in enumerate(payments):
        x = left + i * (bar_w + gap)
        is_we = p.day.weekday() >= 5
        if is_we:
            weekend_bg.append(
                f'<rect x="{x - 0.5}" y="{top}" width="{bar_w + gap}" height="{plot_h}" fill="#f3e8e8"/>'
            )
        total = p.jt + p.kt
        if total <= 0:
            continue
        h = max(1.0, total / max_amt * (plot_h - 2))
        y = top + plot_h - h
        if p.kt > 0 and p.jt > 0:
            h_kt = h * (p.kt / total)
            h_jt = h - h_kt
            bars.append(
                f'<rect x="{x}" y="{y + h_jt}" width="{bar_w}" height="{h_kt}" fill="#c45c26">'
                f'<title>{p.day.isoformat()}: KT {fmt_money(p.kt)} · JT {fmt_money(p.jt)}</title></rect>'
            )
            bars.append(
                f'<rect x="{x}" y="{y}" width="{bar_w}" height="{h_jt}" fill="#022f40">'
                f'<title>{p.day.isoformat()}: KT {fmt_money(p.kt)} · JT {fmt_money(p.jt)}</title></rect>'
            )
        elif p.kt > 0:
            bars.append(
                f'<rect x="{x}" y="{y}" width="{bar_w}" height="{h}" fill="#c45c26">'
                f'<title>{p.day.isoformat()}: KT {fmt_money(p.kt)}</title></rect>'
            )
        else:
            bars.append(
                f'<rect x="{x}" y="{y}" width="{bar_w}" height="{h}" fill="#022f40">'
                f'<title>{p.day.isoformat()}: JT {fmt_money(p.jt)}</title></rect>'
            )

    # y-axis labels
    y_labels = []
    for frac in (0, 0.5, 1.0):
        yy = top + plot_h - frac * (plot_h - 2)
        val = max_amt * frac
        y_labels.append(
            f'<text x="{left - 6}" y="{yy + 3}" text-anchor="end" class="axis">{fmt_money(val)}</text>'
            f'<line x1="{left}" y1="{yy}" x2="{width - 8}" y2="{yy}" stroke="#e5e5e5" stroke-width="1"/>'
        )

    # month ticks on x
    x_labels = []
    last_month = None
    for i, p in enumerate(payments):
        if p.day.day == 1 or (last_month is None):
            if p.day.month != last_month:
                last_month = p.day.month
                x = left + i * (bar_w + gap)
                label = p.day.strftime("%b")
                x_labels.append(
                    f'<text x="{x}" y="{height - 6}" class="axis">{html_escape(label)}</text>'
                )

    return f"""
<div class="chart-wrap" id="{html_escape(chart_id)}">
  <div class="legend">
    <span><i class="swatch jt"></i>JET TECHNIC</span>
    <span><i class="swatch kt"></i>KT MNT</span>
    <span class="muted">розовый фон = сб/вс</span>
  </div>
  <div class="chart-scroll">
    <svg viewBox="0 0 {width} {height}" width="{width}" height="{height}" role="img" aria-label="Платежи поставщикам по дням">
      {''.join(weekend_bg)}
      {''.join(y_labels)}
      {''.join(bars)}
      {''.join(x_labels)}
    </svg>
  </div>
</div>
"""


def render_period_section(period: PeriodReport, idx: int) -> str:
    shares_rows = []
    for s in period.shares:
        cls = "channel-jt" if s.channel == "JT" else ("channel-kt" if s.channel == "KT" else "")
        shares_rows.append(
            f'<tr class="{cls}">'
            f'<td>{html_escape(s.name)}</td>'
            f'<td class="num" data-value="{s.revenue}">{fmt_money(s.revenue)}</td>'
            f'<td class="num" data-value="{s.revenue_pct:.3f}">{fmt_pct(s.revenue_pct)}</td>'
            f'<td class="num" data-value="{s.margin}">{fmt_money(s.margin)}</td>'
            f'<td class="num" data-value="{s.margin_pct:.3f}">{fmt_pct(s.margin_pct)}</td>'
            f'<td class="num" data-value="{s.n}">{s.n}</td>'
            f"</tr>"
        )

    def speed_card(sp: PaySpeed) -> str:
        return f"""
        <div>
          <div class="label">{html_escape(sp.label)}</div>
          <div class="value">{fmt_days(sp.mean_days)} дн.</div>
          <div class="muted">медиана {fmt_days(sp.median_days)} · n={sp.n}</div>
        </div>"""

    kt_note = ""
    if not period.kt_found:
        kt_note = (
            '<p class="hint warn">KT MNT не найден в столбце Z (Поставщик) в этом ТАЗ — '
            "в долях и скорости оплаты по KT показатели нулевые. Учтён только JET TECHNIC.</p>"
        )

    return f"""
<section class="card period" id="period-{idx}">
  <h2>{html_escape(period.title)}</h2>
  <p class="period-range">{period.start.strftime('%d.%m.%Y')} — {period.end.strftime('%d.%m.%Y')} · заказы по дате Q · {period.total_n} поз.</p>
  {kt_note}

  <h3>1. Доли поставщиков</h3>
  <div class="kpis">
    <div class="highlight"><div class="label">Выручка всего</div><div class="value">{fmt_money(period.total_revenue)}</div><div class="muted">USD · продажная итого</div></div>
    <div><div class="label">Маржа всего</div><div class="value">{fmt_money(period.total_margin)}</div><div class="muted">продажа − закупка</div></div>
    <div><div class="label">JT выручка</div><div class="value">{fmt_pct(next((s.revenue_pct for s in period.shares if s.channel=='JT'), 0.0))}</div><div class="muted">доля от выручки</div></div>
    <div><div class="label">KT выручка</div><div class="value">{fmt_pct(next((s.revenue_pct for s in period.shares if s.channel=='KT'), 0.0))}</div><div class="muted">доля от выручки</div></div>
  </div>
  <div class="table-scroll">
  <table class="sortable" data-table="shares-{idx}">
    <thead>
      <tr>
        <th data-type="str" data-col="0">Поставщик <span class="arrow">↕</span></th>
        <th data-type="num" data-col="1">Выручка, USD <span class="arrow">↕</span></th>
        <th data-type="num" data-col="2">% выручки <span class="arrow">↕</span></th>
        <th data-type="num" data-col="3">Маржа, USD <span class="arrow">↕</span></th>
        <th data-type="num" data-col="4">% маржи <span class="arrow">↕</span></th>
        <th data-type="num" data-col="5">n <span class="arrow">↕</span></th>
      </tr>
    </thead>
    <tbody>
      {''.join(shares_rows)}
    </tbody>
  </table>
  </div>

  <h3>2. Time to pay (только JT + KT, STK)</h3>
  <div class="kpis three">
    {speed_card(period.both_speed)}
    {speed_card(period.jt_speed)}
    {speed_card(period.kt_speed)}
  </div>
  <p class="hint">
    Срок = дата оплаты поставщику (AW) − дата взятия в работу (Q).
    В среднее: только Lead time = STK; исключены постоплата (есть дата начала движения BA раньше оплаты AW)
    и заказы с ненулевым lead time (оплата после LT — норма).
    Исключено как постоплата: JT {period.jt_speed.excluded_postpay}, KT {period.kt_speed.excluded_postpay};
    как lead-time: JT {period.jt_speed.excluded_lead}, KT {period.kt_speed.excluded_lead}.
  </p>

  <h3>3. Платежи по дням (AW)</h3>
  {render_histogram(period.payments, f"chart-{idx}")}
  <p class="hint">Суммы оплат поставщику (AX) по дате AW. Несколько заказов в один день суммируются; JT и KT в одном дне — сегменты одной колонки.</p>
</section>
"""


def render_html(periods: list[PeriodReport], source_name: str) -> str:
    sections = "\n".join(render_period_section(p, i) for i, p in enumerate(periods))
    return f"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>TIME TO PAY — JT &amp; KT</title>
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
  padding:6px 10px; font-weight:700; margin-bottom:14px;
  font-size:13px; letter-spacing:.12em; text-transform:uppercase;
}}
h1 {{ color:#fff; font-size:28px; margin:0 0 6px; }}
.sub {{ color:#d5fbff; margin:0 0 18px; font-size:14px; }}
.card {{
  background:var(--card); border:1px solid var(--line); border-radius:8px;
  padding:18px; margin-bottom:18px;
}}
h2 {{ margin:0 0 6px; color:var(--navy); font-size:22px; }}
h3 {{ margin:22px 0 10px; color:var(--navy); font-size:16px; }}
.period-range {{ margin:0 0 14px; color:var(--muted); font-size:13px; }}
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
tbody tr:nth-child(even) {{ background:var(--zebra); }}
tbody tr:hover {{ background:#eef7f9; }}
tr.channel-jt td {{ background:#e8f2f5 !important; font-weight:700; }}
tr.channel-kt td {{ background:#f8ebe3 !important; font-weight:700; }}
.hint {{ margin-top:12px; color:var(--muted); font-size:12px; }}
.hint.warn {{ color:#8a3b12; background:#fff4ec; border:1px solid #f0c7a8; padding:10px 12px; border-radius:8px; }}
.chart-wrap {{ margin-top:8px; }}
.legend {{ display:flex; gap:16px; align-items:center; margin-bottom:8px; font-size:13px; }}
.swatch {{ display:inline-block; width:12px; height:12px; border-radius:2px; margin-right:6px; vertical-align:middle; }}
.swatch.jt {{ background:var(--jt); }}
.swatch.kt {{ background:var(--kt); }}
.chart-scroll {{ overflow-x:auto; border:1px solid var(--line); border-radius:8px; background:#fff; padding:8px; }}
svg .axis {{ font-size:10px; fill:#757575; font-family: Arial, sans-serif; }}
@media (max-width:800px) {{
  .kpis, .kpis.three {{ grid-template-columns:repeat(2,minmax(0,1fr)); }}
}}
</style>
</head>
<body>
<div class="wrap">
  <div class="brand">not so fastair</div>
  <h1>TIME TO PAY</h1>
  <div class="sub">Доли поставщиков · скорость оплаты JT / KT · платежи по дням · источник {html_escape(source_name)}</div>
  {sections}
  <p class="hint">
    Выручка = «Продажная, итого»; маржа = продажа − закупка.
    Каналы JT/KT определяются по столбцу Z (Поставщик). Root supplier (AA) — контрагент, которому платят JT/KT.
  </p>
</div>
<script>
(function () {{
  function cellValue(td, type) {{
    const raw = td.getAttribute('data-value');
    if (raw === null || raw === '') {{
      const t = td.textContent.trim();
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
    const headers = [...table.tHead.rows[0].cells];
    let sortCol = -1;
    let asc = true;
    headers.forEach(th => {{
      th.addEventListener('click', () => {{
        const col = Number(th.dataset.col);
        const type = th.dataset.type;
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


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--taz",
        type=Path,
        default=Path("/tmp/taz_history/ТАЗ 18.09.2026.xlsx"),
    )
    parser.add_argument("--out-dir", type=Path, default=Path("output"))
    parser.add_argument("--stem", default="time_to_pay")
    args = parser.parse_args()

    df = prepare(load_taz(args.taz))
    periods = [
        build_period(df, "Весь ТАЗ 2026", date(2026, 1, 1), date(2026, 12, 31)),
        build_period(df, "01.06 — 18.09.2026", date(2026, 6, 1), date(2026, 9, 18)),
    ]

    for p in periods:
        print(
            f"{p.title}: n={p.total_n} rev={p.total_revenue:.0f} "
            f"TTP JT mean={p.jt_speed.mean_days} n={p.jt_speed.n} "
            f"KT n={p.kt_speed.n} kt_found={p.kt_found}"
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
