#!/usr/bin/env python3
"""FASTAIR — Выполненные заказы: только EXPENDABLE (вкл. consumable), все клиенты.

Попадание / периоды — как в completed_orders (W, без cancel/refund/warranty…).

В каждом периоде:
  1) По поставщикам — шапка KPI + выпадающий список по каждому поставщику
     (метрики + позиции)
  2) По клиентам — выпадающий клиент; внутри — мини-отчёты по поставщикам
     (та же логика, что в разделе 1)
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

import pandas as pd

from generate_completed_orders_report import (
    CAT_EXPENDABLE,
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
    kpi: SliceKPI
    positions: list = field(default_factory=list)


@dataclass
class ClientBlock:
    name: str
    kpi: SliceKPI
    suppliers: list[SupplierBlock] = field(default_factory=list)


@dataclass
class PeriodView:
    title: str
    start: date
    end: date
    total: SliceKPI
    suppliers: list[SupplierBlock] = field(default_factory=list)
    clients: list[ClientBlock] = field(default_factory=list)


def filter_expendables(df: pd.DataFrame) -> pd.DataFrame:
    return df.loc[df["_bucket"].eq(CAT_EXPENDABLE)].copy()


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


def _positions(part: pd.DataFrame) -> list:
    return [
        line_from_row(r)
        for _, r in part.sort_values("_sale", ascending=False).iterrows()
    ]


def build_suppliers(rows: pd.DataFrame, lead_rows: pd.DataFrame) -> list[SupplierBlock]:
    names = sorted(
        rows["_supplier"].unique(),
        key=lambda s: (-float(rows.loc[rows["_supplier"] == s, "_sale"].sum()), str(s).casefold()),
    )
    out: list[SupplierBlock] = []
    for name in names:
        part = rows[rows["_supplier"] == name]
        lead_part = lead_rows[lead_rows["_supplier"] == name]
        out.append(
            SupplierBlock(
                name=str(name),
                kpi=_kpi(part, lead_part),
                positions=_positions(part),
            )
        )
    return out


def build_clients(rows: pd.DataFrame, lead_rows: pd.DataFrame) -> list[ClientBlock]:
    names = sorted(
        rows["_client"].unique(),
        key=lambda s: (-float(rows.loc[rows["_client"] == s, "_sale"].sum()), str(s).casefold()),
    )
    out: list[ClientBlock] = []
    for name in names:
        part = rows[rows["_client"] == name]
        lead_part = lead_rows[lead_rows["_client"] == name]
        out.append(
            ClientBlock(
                name=str(name),
                kpi=_kpi(part, lead_part),
                suppliers=build_suppliers(part, lead_part),
            )
        )
    return out


def build_period_view(df: pd.DataFrame, title: str, start: date, end: date) -> PeriodView:
    rows = df.loc[report_mask(df, start, end)].copy()
    lead_rows = df.loc[lead_mask(df, start, end)].copy()
    return PeriodView(
        title=title,
        start=start,
        end=end,
        total=_kpi(rows, lead_rows),
        suppliers=build_suppliers(rows, lead_rows),
        clients=build_clients(rows, lead_rows),
    )


def kpi_meta(kpi: SliceKPI) -> str:
    mp = f" · {kpi.margin_pct:.1f}%" if kpi.margin_pct is not None else ""
    return (
        f"{fmt_int(kpi.orders)} зак. · {fmt_int(kpi.lines)} поз. · "
        f"${fmt_money(kpi.revenue)} · маржа ${fmt_money(kpi.margin)}{mp} · "
        f"срок {fmt_days(kpi.lead_avg)} дн. · {fmt_transport_delta_html(kpi.transport_delta)}"
    )


def render_header_kpis(kpi: SliceKPI) -> str:
    return f"""
    <div class="kpis">
      <div class="highlight">
        <div class="label">Выполнено заказов</div>
        <div class="value">{fmt_int(kpi.orders)}</div>
        <div class="muted">{fmt_int(kpi.lines)} позиций EXPENDABLE с W</div>
      </div>
      <div>
        <div class="label">Выручка</div>
        <div class="value">${fmt_money(kpi.revenue)}</div>
        <div class="muted">Σ продажная итого</div>
      </div>
      <div>
        <div class="label">Маржа</div>
        <div class="value">${fmt_money(kpi.margin)}</div>
        <div class="muted">продажа − закупка − тр. − fee{(' · ' + f'{kpi.margin_pct:.1f}%') if kpi.margin_pct is not None else ''}</div>
      </div>
      <div>
        <div class="label">Сроки STK</div>
        <div class="value">{fmt_days(kpi.lead_avg)} дн.</div>
        <div class="muted">ср. W−Q · {fmt_int(kpi.lead_n)} поз.</div>
      </div>
      <div class="wide">
        <div class="label">Транспорт план − факт = дельта</div>
        <div class="value transport">
          <span>{fmt_money(kpi.transport_plan)}</span>
          <span class="op">−</span>
          <span>{fmt_money(kpi.transport_fact)}</span>
          <span class="op">=</span>
          {fmt_transport_delta_html(kpi.transport_delta)}
        </div>
        <div class="muted">USD · только EXPENDABLE (вкл. consumable)</div>
      </div>
    </div>
    """


def render_supplier_dropdown(sup: SupplierBlock, table_id: str, *, opened: bool = False) -> str:
    open_attr = " open" if opened else ""
    body = (
        render_positions_table(sup.positions, table_id)
        if sup.positions
        else "<p class='empty'>Нет позиций</p>"
    )
    return f"""
<details class="subblock"{open_attr}>
  <summary>
    <span class="ent-name">{html_escape(sup.name)}</span>
    <span class="ent-meta">{kpi_meta(sup.kpi)}</span>
  </summary>
  <div class="block-body">
    {render_header_kpis(sup.kpi)}
    {body}
  </div>
</details>
"""


def render_client_dropdown(client: ClientBlock, prefix: str) -> str:
    suppliers = "\n".join(
        render_supplier_dropdown(s, f"{prefix}-s{i}", opened=False)
        for i, s in enumerate(client.suppliers)
    )
    if not suppliers:
        suppliers = "<p class='empty'>Нет поставщиков</p>"
    return f"""
<details class="block">
  <summary>
    <span class="ent-name">{html_escape(client.name)}</span>
    <span class="ent-meta">{kpi_meta(client.kpi)} · поставщиков {fmt_int(len(client.suppliers))}</span>
  </summary>
  <div class="block-body">
    {render_header_kpis(client.kpi)}
    <h3 class="section-h">Поставщики ({fmt_int(len(client.suppliers))})</h3>
    {suppliers}
  </div>
</details>
"""


def render_period(period: PeriodView, idx: int, *, opened: bool) -> str:
    open_attr = " open" if opened else ""
    range_s = f"{period.start.strftime('%d.%m.%Y')} – {period.end.strftime('%d.%m.%Y')}"
    suppliers_html = "\n".join(
        render_supplier_dropdown(s, f"p{idx}-sup-{i}")
        for i, s in enumerate(period.suppliers)
    ) or "<p class='empty'>Нет поставщиков</p>"
    clients_html = "\n".join(
        render_client_dropdown(c, f"p{idx}-cli-{i}")
        for i, c in enumerate(period.clients)
    ) or "<p class='empty'>Нет клиентов</p>"
    return f"""
<details class="period"{open_attr}>
  <summary>
    <span class="period-title">{idx}. {html_escape(period.title)}</span>
    <span class="period-meta">{range_s}</span>
    <span class="period-meta">{fmt_int(period.total.orders)} зак. · ${fmt_money(period.total.revenue)} выр. · маржа ${fmt_money(period.total.margin)}</span>
  </summary>
  <div class="period-body">
    {render_header_kpis(period.total)}

    <details class="section" open>
      <summary>1. По поставщикам ({fmt_int(len(period.suppliers))})</summary>
      <div class="block-body">
        {suppliers_html}
      </div>
    </details>

    <details class="section" open>
      <summary>2. По клиентам ({fmt_int(len(period.clients))})</summary>
      <div class="block-body">
        <p class="section-note">Внутри клиента — мини-отчёты по поставщикам (как в разделе 1).</p>
        {clients_html}
      </div>
    </details>
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
<title>FASTAIR — EXPENDABLE · поставщики и клиенты</title>
<style>
:root {{
  --navy:#04384d; --line:#c5d4da; --muted:#6b7c83; --zebra:#f4f8f9;
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
.sub {{ color:#d5fbff; margin:0; font-size:14px; }}
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
details.section, details.block, details.subblock {{
  border:1px solid #d7e2e6; border-radius:8px; margin:10px 0; background:#fff;
}}
details.section > summary {{
  cursor:pointer; padding:12px 14px; list-style:none; background:#e8f6f8;
  font-weight:700; color:var(--navy); font-size:15px;
}}
details.block > summary, details.subblock > summary {{
  cursor:pointer; padding:10px 12px; font-weight:700; color:var(--navy);
  list-style:none; background:#f7fbfc;
  display:flex; flex-wrap:wrap; gap:8px 14px; align-items:baseline;
}}
details.section > summary::-webkit-details-marker,
details.block > summary::-webkit-details-marker,
details.subblock > summary::-webkit-details-marker {{ display:none; }}
details.section > summary::before,
details.block > summary::before,
details.subblock > summary::before {{ content:"▸ "; color:#7a8a90; }}
details.section[open] > summary::before,
details.block[open] > summary::before,
details.subblock[open] > summary::before {{ content:"▾ "; }}
.block-body {{ padding:12px; }}
.ent-name {{ font-weight:700; }}
.ent-meta {{ font-size:12px; color:var(--muted); font-weight:600; }}
.section-h {{ margin:16px 0 8px; font-size:14px; color:var(--navy); }}
.section-note {{ margin:0 0 12px; color:var(--muted); font-size:13px; }}
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
  .kpis {{ grid-template-columns:1fr 1fr; }}
}}
</style>
</head>
<body>
<div class="wrap">
  <div class="hero">
    <div class="eyebrow">FASTAIR</div>
    <h1>EXPENDABLE · поставщики и клиенты</h1>
    <div class="sub">Срез ТАЗ {html_escape(as_of.strftime('%d.%m.%Y'))} · {html_escape(source_name)} · только расходка (вкл. consumable)</div>
  </div>
  <div class="rules">
    <strong>Срез:</strong> только <strong>EXPENDABLE</strong> (consumable сюда же), все клиенты.
    <br/>
    <strong>Попадание:</strong> заполненная W; без CANCELLED / REFUND / WARRANTY / SCRAPPED / LOST. FINISHED не фильтруем.
    <br/>
    <strong>Структура периода:</strong> 1) по поставщикам (шапка + выпадающие списки с метриками) ·
    2) по клиентам (внутри — мини-отчёты по поставщикам).
    <br/>
    <strong>Метрики:</strong> заказы, позиции, выручка, маржа, срок STK (W−Q), Δ транспорта.
  </div>
  {sections}
  <p class="hint">Открывайте HTML напрямую.</p>
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
    parser.add_argument("--stem", default="expendables_suppliers_clients")
    args = parser.parse_args()

    as_of = infer_taz_end(args.taz)
    df = filter_expendables(prepare(load_taz(args.taz)))
    base = build_periods(df, as_of)
    periods = [build_period_view(df, p.title, p.start, p.end) for p in base]

    for p in periods:
        print(
            f"{p.title}: orders={p.total.orders} lines={p.total.lines} "
            f"rev={p.total.revenue:.0f} margin={p.total.margin:.0f} "
            f"suppliers={len(p.suppliers)} clients={len(p.clients)}"
        )

    html = render_html(periods, args.taz.name, as_of)
    args.out_dir.mkdir(parents=True, exist_ok=True)
    out = args.out_dir / f"{args.stem}.html"
    out.write_text(html, encoding="utf-8")
    print(f"wrote {out}")


if __name__ == "__main__":
    main()
