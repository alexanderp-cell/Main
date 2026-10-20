#!/usr/bin/env python3
"""FASTAIR — Выполненные заказы: Utair + Белавиа, доля BLUE OCEAN.

Те же правила попадания и периоды, что в completed_orders:
  · только заполненная W; без CANCEL/REFUND/WARRANTY/SCRAP/LOST
  · FINISHED не фильтруем
  · маржа = продажа − закупка − транспорт − fee

Структура:
  период → клиент (Utair / Белавиа) → поставщики (выпадающие)
  + сравнение BLUE OCEAN vs остальные (доля, сроки STK, маржа, Δ тр.).
"""

from __future__ import annotations

import argparse
import re
from dataclasses import dataclass, field
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

from generate_completed_orders_report import (
    CAT_EXPENDABLE,
    CAT_ROTABLE,
    as_money,
    build_periods,
    fmt_days,
    fmt_int,
    fmt_money,
    fmt_transport_delta_html,
    html_escape,
    infer_taz_end,
    lead_mask,
    line_from_row,
    load_taz,
    prepare,
    render_positions_table,
    report_mask,
)

TARGET_CLIENTS = ("Utair", "Белавиа")
BO_TOKEN = "BLUE OCEAN"


def normalize_client(name: str) -> str:
    text = (name or "").strip()
    low = text.casefold()
    if "utair" in low or "ютэ" in low or "ютейр" in low:
        return "Utair"
    if "белавиа" in low or "belavia" in low:
        return "Белавиа"
    return text


def is_blue_ocean(supplier: str) -> bool:
    return BO_TOKEN in (supplier or "").upper()


def filter_target_clients(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["_client_norm"] = out["_client"].map(normalize_client)
    return out.loc[out["_client_norm"].isin(TARGET_CLIENTS)].copy()


@dataclass
class SliceKPI:
    orders: int
    lines: int
    revenue: float
    margin: float
    lead_avg: float | None
    lead_n: int
    transport_plan: float
    transport_fact: float

    @property
    def transport_delta(self) -> float:
        return self.transport_fact - self.transport_plan

    @property
    def margin_pct(self) -> float | None:
        if self.revenue <= 0:
            return None
        return 100.0 * self.margin / self.revenue


@dataclass
class SupplierBlock:
    name: str
    is_bo: bool
    kpi: SliceKPI
    positions_rotable: list
    positions_expendable: list
    positions_other: list


@dataclass
class ClientBlock:
    name: str
    kpi: SliceKPI
    bo: SliceKPI
    other: SliceKPI
    suppliers: list[SupplierBlock] = field(default_factory=list)


@dataclass
class PeriodView:
    title: str
    start: date
    end: date
    total: SliceKPI
    bo: SliceKPI
    other: SliceKPI
    clients: list[ClientBlock] = field(default_factory=list)


def _kpi(rows: pd.DataFrame, lead_rows: pd.DataFrame) -> SliceKPI:
    if rows.empty:
        return SliceKPI(0, 0, 0.0, 0.0, None, 0, 0.0, 0.0)
    invoices = {inv for inv in rows["_invoice"] if inv}
    days = lead_rows["_days"].astype(float) if not lead_rows.empty else pd.Series(dtype=float)
    return SliceKPI(
        orders=len(invoices) if invoices else int(len(rows)),
        lines=int(len(rows)),
        revenue=as_money(rows["_sale"].sum()),
        margin=as_money(rows["_margin"].sum()),
        lead_avg=float(days.mean()) if not days.empty else None,
        lead_n=int(len(days)),
        transport_plan=as_money(rows["_plan"].sum()),
        transport_fact=as_money(rows["_fact"].sum()),
    )


def _positions_by_bucket(part: pd.DataFrame) -> tuple[list, list, list]:
    def pos(bucket: str):
        sub = part[part["_bucket"] == bucket].sort_values("_sale", ascending=False)
        return [line_from_row(r) for _, r in sub.iterrows()]

    other = part[~part["_bucket"].isin({CAT_ROTABLE, CAT_EXPENDABLE})].sort_values(
        "_sale", ascending=False
    )
    return (
        pos(CAT_ROTABLE),
        pos(CAT_EXPENDABLE),
        [line_from_row(r) for _, r in other.iterrows()],
    )


def build_client(rows: pd.DataFrame, lead_rows: pd.DataFrame, name: str) -> ClientBlock:
    part = rows[rows["_client_norm"] == name]
    lead_part = lead_rows[lead_rows["_client_norm"] == name]
    bo_mask = part["_supplier"].map(is_blue_ocean)
    bo_rows = part.loc[bo_mask]
    other_rows = part.loc[~bo_mask]
    bo_lead = lead_part.loc[lead_part["_supplier"].map(is_blue_ocean)]
    other_lead = lead_part.loc[~lead_part["_supplier"].map(is_blue_ocean)]

    suppliers: list[SupplierBlock] = []
    for sup in sorted(part["_supplier"].unique(), key=lambda s: (not is_blue_ocean(str(s)), -float(part.loc[part["_supplier"] == s, "_sale"].sum()), str(s).casefold())):
        s_part = part[part["_supplier"] == sup]
        s_lead = lead_part[lead_part["_supplier"] == sup]
        rot, exp, oth = _positions_by_bucket(s_part)
        suppliers.append(
            SupplierBlock(
                name=str(sup),
                is_bo=is_blue_ocean(str(sup)),
                kpi=_kpi(s_part, s_lead),
                positions_rotable=rot,
                positions_expendable=exp,
                positions_other=oth,
            )
        )

    return ClientBlock(
        name=name,
        kpi=_kpi(part, lead_part),
        bo=_kpi(bo_rows, bo_lead),
        other=_kpi(other_rows, other_lead),
        suppliers=suppliers,
    )


def build_period_view(df: pd.DataFrame, title: str, start: date, end: date) -> PeriodView:
    rows = df.loc[report_mask(df, start, end)].copy()
    lead_rows = df.loc[lead_mask(df, start, end)].copy()
    bo_mask = rows["_supplier"].map(is_blue_ocean)
    bo_lead_mask = lead_rows["_supplier"].map(is_blue_ocean)
    clients = [build_client(rows, lead_rows, name) for name in TARGET_CLIENTS]
    clients = [c for c in clients if c.kpi.lines > 0]
    return PeriodView(
        title=title,
        start=start,
        end=end,
        total=_kpi(rows, lead_rows),
        bo=_kpi(rows.loc[bo_mask], lead_rows.loc[bo_lead_mask]),
        other=_kpi(rows.loc[~bo_mask], lead_rows.loc[~bo_lead_mask]),
        clients=clients,
    )


def share_pct(part: float, total: float) -> str:
    if total <= 0:
        return "—"
    return f"{100.0 * part / total:.1f}%"


def render_kpi_card(label: str, value: str, muted: str = "", highlight: bool = False) -> str:
    cls = " highlight" if highlight else ""
    muted_html = f'<div class="muted">{html_escape(muted)}</div>' if muted else ""
    return f"""
      <div class="{cls.strip()}">
        <div class="label">{html_escape(label)}</div>
        <div class="value">{value}</div>
        {muted_html}
      </div>
    """


def render_compare(bo: SliceKPI, other: SliceKPI, total: SliceKPI) -> str:
    return f"""
    <div class="compare">
      <div class="compare-card bo">
        <div class="compare-title">BLUE OCEAN</div>
        <div class="compare-grid">
          <div><span class="k">Доля заказов</span><span class="v">{share_pct(bo.orders, total.orders)}</span></div>
          <div><span class="k">Доля выручки</span><span class="v">{share_pct(bo.revenue, total.revenue)}</span></div>
          <div><span class="k">Доля маржи</span><span class="v">{share_pct(bo.margin, total.margin)}</span></div>
          <div><span class="k">Заказы / поз.</span><span class="v">{fmt_int(bo.orders)} / {fmt_int(bo.lines)}</span></div>
          <div><span class="k">Выручка</span><span class="v">${fmt_money(bo.revenue)}</span></div>
          <div><span class="k">Маржа</span><span class="v">${fmt_money(bo.margin)}{(' · ' + f'{bo.margin_pct:.1f}%') if bo.margin_pct is not None else ''}</span></div>
          <div><span class="k">Срок STK</span><span class="v">{fmt_days(bo.lead_avg)} дн. · {fmt_int(bo.lead_n)} поз.</span></div>
          <div><span class="k">Δ тр.</span><span class="v">{fmt_transport_delta_html(bo.transport_delta)}</span></div>
        </div>
      </div>
      <div class="compare-card">
        <div class="compare-title">Остальные поставщики</div>
        <div class="compare-grid">
          <div><span class="k">Доля заказов</span><span class="v">{share_pct(other.orders, total.orders)}</span></div>
          <div><span class="k">Доля выручки</span><span class="v">{share_pct(other.revenue, total.revenue)}</span></div>
          <div><span class="k">Доля маржи</span><span class="v">{share_pct(other.margin, total.margin)}</span></div>
          <div><span class="k">Заказы / поз.</span><span class="v">{fmt_int(other.orders)} / {fmt_int(other.lines)}</span></div>
          <div><span class="k">Выручка</span><span class="v">${fmt_money(other.revenue)}</span></div>
          <div><span class="k">Маржа</span><span class="v">${fmt_money(other.margin)}{(' · ' + f'{other.margin_pct:.1f}%') if other.margin_pct is not None else ''}</span></div>
          <div><span class="k">Срок STK</span><span class="v">{fmt_days(other.lead_avg)} дн. · {fmt_int(other.lead_n)} поз.</span></div>
          <div><span class="k">Δ тр.</span><span class="v">{fmt_transport_delta_html(other.transport_delta)}</span></div>
        </div>
      </div>
    </div>
    """


def render_supplier(sup: SupplierBlock, table_prefix: str) -> str:
    badge = ' <span class="bo-badge">BLUE OCEAN</span>' if sup.is_bo else ""
    blocks = []
    for label, positions, j in (
        ("ROTABLE · вкл. GSE", sup.positions_rotable, 0),
        ("EXPENDABLE · вкл. consumable", sup.positions_expendable, 1),
        ("Прочее", sup.positions_other, 2),
    ):
        if not positions:
            continue
        blocks.append(
            f"""
<details class="catblock" open>
  <summary><span class="cat-name">{html_escape(label)}</span>
  <span class="ent-meta">{fmt_int(len(positions))} поз.</span></summary>
  <div class="block-body">{render_positions_table(positions, f"{table_prefix}-{j}")}</div>
</details>
"""
        )
    inner = "\n".join(blocks) if blocks else "<p class='empty'>Нет позиций</p>"
    open_attr = " open" if sup.is_bo else ""
    return f"""
<details class="subblock{' bo-block' if sup.is_bo else ''}"{open_attr}>
  <summary>
    <span class="ent-name">{html_escape(sup.name)}{badge}</span>
    <span class="ent-meta">{fmt_int(sup.kpi.orders)} зак. · {fmt_int(sup.kpi.lines)} поз. · ${fmt_money(sup.kpi.revenue)} · маржа ${fmt_money(sup.kpi.margin)} · срок {fmt_days(sup.kpi.lead_avg)} дн. · {fmt_transport_delta_html(sup.kpi.transport_delta)}</span>
  </summary>
  <div class="block-body">{inner}</div>
</details>
"""


def render_client(client: ClientBlock, idx: int, pidx: int) -> str:
    suppliers_html = "\n".join(
        render_supplier(s, f"p{pidx}-c{idx}-s{i}") for i, s in enumerate(client.suppliers)
    )
    return f"""
<details class="block" open>
  <summary>
    <span class="ent-name">{html_escape(client.name)}</span>
    <span class="ent-meta">{fmt_int(client.kpi.orders)} зак. · {fmt_int(client.kpi.lines)} поз. · ${fmt_money(client.kpi.revenue)} · маржа ${fmt_money(client.kpi.margin)} · BO доля выр. {share_pct(client.bo.revenue, client.kpi.revenue)} · срок {fmt_days(client.kpi.lead_avg)} дн. · {fmt_transport_delta_html(client.kpi.transport_delta)}</span>
  </summary>
  <div class="block-body">
    {render_compare(client.bo, client.other, client.kpi)}
    <h3 class="section-h">Поставщики ({len(client.suppliers)})</h3>
    {suppliers_html}
  </div>
</details>
"""


def render_period(period: PeriodView, idx: int, *, opened: bool) -> str:
    open_attr = " open" if opened else ""
    range_s = f"{period.start.strftime('%d.%m.%Y')} – {period.end.strftime('%d.%m.%Y')}"
    clients_html = "\n".join(render_client(c, i, idx) for i, c in enumerate(period.clients))
    if not clients_html:
        clients_html = "<p class='empty'>Нет выполненных позиций Utair / Белавиа</p>"
    return f"""
<details class="period"{open_attr}>
  <summary>
    <span class="period-title">{idx}. {html_escape(period.title)}</span>
    <span class="period-meta">{range_s}</span>
    <span class="period-meta">{fmt_int(period.total.orders)} зак. · ${fmt_money(period.total.revenue)} выр. · маржа ${fmt_money(period.total.margin)} · BO {share_pct(period.bo.revenue, period.total.revenue)} выр.</span>
  </summary>
  <div class="period-body">
    <div class="kpis">
      {render_kpi_card("Выполнено заказов", fmt_int(period.total.orders), f"{fmt_int(period.total.lines)} позиций с W", True)}
      {render_kpi_card("Выручка", f"${fmt_money(period.total.revenue)}", "Utair + Белавиа")}
      {render_kpi_card("Маржа", f"${fmt_money(period.total.margin)}", "продажа − закупка − тр. − fee")}
      {render_kpi_card("Сроки STK", f"{fmt_days(period.total.lead_avg)} дн.", f"{fmt_int(period.total.lead_n)} поз. · W−Q")}
      <div class="wide">
        <div class="label">Доля BLUE OCEAN</div>
        <div class="value transport">
          <span>зак. {share_pct(period.bo.orders, period.total.orders)}</span>
          <span class="op">·</span>
          <span>выр. {share_pct(period.bo.revenue, period.total.revenue)}</span>
          <span class="op">·</span>
          <span>маржа {share_pct(period.bo.margin, period.total.margin)}</span>
          <span class="op">·</span>
          <span>срок BO {fmt_days(period.bo.lead_avg)} / ост. {fmt_days(period.other.lead_avg)} дн.</span>
        </div>
        <div class="muted">сравнение качества ниже по каждому клиенту</div>
      </div>
    </div>
    {render_compare(period.bo, period.other, period.total)}
    {clients_html}
  </div>
</details>
"""


def render_html(periods: list[PeriodView], source_name: str, as_of: date) -> str:
    sections = "\n".join(
        render_period(p, i + 1, opened=(i == 0)) for i, p in enumerate(periods)
    )
    return f"""<!doctype html>
<html lang="ru">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>FASTAIR — Utair / Белавиа · BLUE OCEAN</title>
<style>
:root {{
  --navy:#04384d; --line:#c5d4da; --muted:#6b7c83; --zebra:#f4f8f9;
  --bo:#0b5c6b; --bo-bg:#e7f5f7;
}}
* {{ box-sizing:border-box; }}
body {{
  margin:0; font-family: Arial, Calibri, Helvetica, sans-serif;
  background:linear-gradient(180deg,#eef5f7 0%,#f7fbfc 40%,#fff 100%);
  color:#1a2a30;
}}
.wrap {{ max-width:1180px; margin:0 auto; padding:28px 20px 64px; }}
.hero {{
  background:linear-gradient(135deg,#04384d 0%,#0a5c6b 55%,#12707f 100%);
  padding:22px 22px 18px; border-radius:14px; margin-bottom:16px;
}}
.eyebrow {{
  display:inline-block; background:rgba(255,255,255,.14); color:#d5fbff;
  padding:6px 10px; font-weight:700; margin-bottom:10px; border-radius:6px; font-size:12px;
}}
h1 {{ color:#fff; font-size:28px; margin:0 0 6px; font-weight:700; }}
.sub {{ color:#d5fbff; margin:0 0 0; font-size:14px; }}
.rules {{
  background:#e8f6f8; border:1px solid #9ed7e0; border-radius:8px;
  padding:12px 14px; margin-bottom:16px; font-size:13px; color:var(--navy);
}}
details.period {{
  border:1px solid var(--line); border-radius:10px; margin:0 0 14px; background:#fff;
  overflow:hidden;
}}
details.period > summary {{
  cursor:pointer; padding:14px 16px; list-style:none; background:#f7fbfc;
  display:flex; flex-wrap:wrap; gap:8px 14px; align-items:baseline;
  border-bottom:1px solid transparent;
}}
details.period[open] > summary {{ border-bottom-color:var(--line); }}
details.period > summary::-webkit-details-marker {{ display:none; }}
.period-title {{ font-size:17px; font-weight:700; color:var(--navy); }}
.period-meta {{ font-size:13px; color:var(--muted); }}
.period-body {{ padding:12px 16px 16px; }}
details.block, details.subblock, details.catblock {{
  border:1px solid #d7e2e6; border-radius:8px; margin:10px 0; background:#fff;
}}
details.catblock {{ margin:8px 0; border-color:#c5d4da; background:#fafcfd; }}
details.block > summary, details.subblock > summary, details.catblock > summary {{
  cursor:pointer; padding:10px 12px; font-weight:700; color:var(--navy);
  list-style:none; background:#f7fbfc;
  display:flex; flex-wrap:wrap; gap:8px 14px; align-items:baseline;
}}
details.catblock > summary {{ background:#eef5f7; font-size:13px; }}
details.subblock.bo-block {{ border-color:#7eb8c2; }}
details.subblock.bo-block > summary {{ background:var(--bo-bg); }}
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
.bo-badge {{
  display:inline-block; margin-left:8px; padding:2px 8px; border-radius:4px;
  background:var(--bo); color:#fff; font-size:11px; font-weight:700; vertical-align:middle;
}}
.section-h {{ margin:16px 0 8px; font-size:14px; color:var(--navy); }}
.kpis {{
  display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:10px; margin-bottom:14px;
}}
.kpis > div {{ background:#f7fbfc; border:1px solid #d7e2e6; border-radius:8px; padding:12px 14px; }}
.kpis > div.highlight {{ background:#e8f6f8; border-color:#9ed7e0; }}
.kpis > div.wide {{ grid-column:1 / -1; }}
.label {{ font-size:11px; text-transform:uppercase; letter-spacing:.04em; color:#757575; }}
.value {{ font-size:22px; font-weight:700; margin-top:4px; color:var(--navy); font-variant-numeric:tabular-nums; }}
.value.transport {{ display:flex; flex-wrap:wrap; gap:8px 12px; align-items:baseline; font-size:18px; }}
.value.transport .op {{ color:var(--muted); font-weight:600; }}
.bad {{ color:#9b2c2c; font-weight:700; }}
.ok {{ color:#0a5c4c; font-weight:700; }}
.muted {{ color:var(--muted); font-size:12px; margin-top:4px; }}
.compare {{
  display:grid; grid-template-columns:1fr 1fr; gap:12px; margin:12px 0 16px;
}}
.compare-card {{
  border:1px solid #d7e2e6; border-radius:8px; padding:12px 14px; background:#fff;
}}
.compare-card.bo {{ background:var(--bo-bg); border-color:#9ed7e0; }}
.compare-title {{ font-weight:700; color:var(--navy); margin-bottom:10px; }}
.compare-grid {{
  display:grid; grid-template-columns:1fr 1fr; gap:8px 12px; font-size:13px;
}}
.compare-grid .k {{ display:block; color:var(--muted); font-size:11px; text-transform:uppercase; letter-spacing:.03em; }}
.compare-grid .v {{ display:block; font-weight:700; color:var(--navy); margin-top:2px; font-variant-numeric:tabular-nums; }}
.table-scroll {{ overflow-x:auto; }}
table {{
  width:100%; border-collapse:collapse; font-size:13px; border:1px solid #b7c8cf;
}}
th, td {{
  border:1px solid #c5d4da; padding:9px 10px; text-align:left; vertical-align:middle;
}}
th {{
  font-size:11px; text-transform:uppercase; letter-spacing:.03em; color:#fff;
  background:var(--navy); cursor:pointer; user-select:none; white-space:nowrap; text-align:center;
}}
th.label-col {{ text-align:left; }}
th:hover {{ background:#03425a; }}
th .arrow {{ opacity:.45; margin-left:6px; font-size:11px; }}
td.num, th.num {{ text-align:center; font-variant-numeric:tabular-nums; white-space:nowrap; }}
td.empty, .empty {{ color:#9a9a9a; text-align:center; }}
tbody tr:nth-child(even) {{ background:var(--zebra); }}
tbody tr:hover {{ background:#eef7f9; }}
.hint {{ margin-top:12px; color:var(--muted); font-size:12px; }}
@media (max-width:900px) {{
  .kpis, .compare, .compare-grid {{ grid-template-columns:1fr; }}
}}
</style>
</head>
<body>
<div class="wrap">
  <div class="hero">
    <div class="eyebrow">FASTAIR</div>
    <h1>Utair + Белавиа · доля BLUE OCEAN</h1>
    <div class="sub">Срез ТАЗ {html_escape(as_of.strftime('%d.%m.%Y'))} · источник {html_escape(source_name)}</div>
  </div>
  <div class="rules">
    <strong>Клиенты:</strong> только <strong>Utair</strong> (вкл. UTair) и <strong>Белавиа</strong>.
    Внутри клиента — выпадающие <strong>поставщики</strong>; BLUE OCEAN выделен.
    <br/>
    <strong>Попадание:</strong> заполненная W; без CANCELLED / REFUND / WARRANTY / SCRAPPED / LOST.
    FINISHED не фильтруем.
    <br/>
    <strong>Качество:</strong> сравниваем долю заказов/выручки/маржи, срок STK (W−Q) и Δ транспорта
    у BLUE OCEAN vs остальные поставщики.
    <br/>
    <strong>Маржа:</strong> продажная итого − закупка итого − транспорт (факт, иначе план) − fee.
  </div>
  {sections}
  <p class="hint">Открывайте HTML напрямую. ZIP+HTML Windows часто помечает ложно.</p>
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


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--taz",
        type=Path,
        default=Path("/tmp/taz_history/ТАЗ 09.10.2026.xlsx"),
    )
    parser.add_argument("--out-dir", type=Path, default=Path("output"))
    parser.add_argument("--stem", default="utair_belavia_blue_ocean")
    args = parser.parse_args()

    as_of = infer_taz_end(args.taz)
    df = filter_target_clients(prepare(load_taz(args.taz)))
    # reuse period windows from base builder, then rebuild focused views
    base_periods = build_periods(df, as_of)
    periods = [
        build_period_view(df, p.title, p.start, p.end) for p in base_periods
    ]

    for p in periods:
        print(
            f"{p.title}: orders={p.total.orders} lines={p.total.lines} "
            f"rev={p.total.revenue:.0f} margin={p.total.margin:.0f} "
            f"BO_rev_share={share_pct(p.bo.revenue, p.total.revenue)} "
            f"BO_lead={p.bo.lead_avg} other_lead={p.other.lead_avg}"
        )
        for c in p.clients:
            print(
                f"  {c.name}: orders={c.kpi.orders} rev={c.kpi.revenue:.0f} "
                f"BO_share={share_pct(c.bo.revenue, c.kpi.revenue)} "
                f"suppliers={len(c.suppliers)}"
            )

    html = render_html(periods, args.taz.name, as_of)
    args.out_dir.mkdir(parents=True, exist_ok=True)
    out_html = args.out_dir / f"{args.stem}.html"
    out_html.write_text(html, encoding="utf-8")
    print(f"wrote {out_html}")


if __name__ == "__main__":
    main()
