#!/usr/bin/env python3
"""FASTAIR — Выполненные заказы по категории (EXPENDABLE / ROTABLE), все клиенты.

Попадание / периоды — как в completed_orders (W, без cancel/refund/scrap/lost).
WARRANTY входит в выполнение (поставка уже была) и дублируется в разделе 3.

В каждом периоде:
  1) По поставщикам — шапка KPI + выпадающий поставщик;
     внутри — выпадающие клиенты с мини-отчётами
  2) По клиентам — выпадающий клиент;
     внутри — выпадающие поставщики с мини-отчётами
  3) WARRANTY — те же позиции со статусом гарантии (не доп. выручка)
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass, field
from datetime import date
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
    warranty_mask,
)


@dataclass(frozen=True)
class BucketSpec:
    key: str
    title: str
    short: str
    include_note: str
    stem: str


BUCKETS: dict[str, BucketSpec] = {
    CAT_EXPENDABLE: BucketSpec(
        key=CAT_EXPENDABLE,
        title="EXPENDABLE",
        short="расходка (вкл. consumable)",
        include_note="consumable сюда же",
        stem="expendables_suppliers_clients",
    ),
    CAT_ROTABLE: BucketSpec(
        key=CAT_ROTABLE,
        title="ROTABLE",
        short="ротейблы (вкл. GSE/TOOL)",
        include_note="GSE/TOOL сюда же",
        stem="rotables_suppliers_clients",
    ),
}


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
class EntityLeaf:
    """Мини-отчёт: метрики + позиции (клиент внутри поставщика или наоборот)."""

    name: str
    kpi: SliceKPI
    positions: list = field(default_factory=list)


@dataclass
class SupplierBlock:
    name: str
    kpi: SliceKPI
    clients: list[EntityLeaf] = field(default_factory=list)


@dataclass
class ClientBlock:
    name: str
    kpi: SliceKPI
    suppliers: list[EntityLeaf] = field(default_factory=list)


@dataclass
class PeriodView:
    title: str
    start: date
    end: date
    total: SliceKPI
    suppliers: list[SupplierBlock] = field(default_factory=list)
    clients: list[ClientBlock] = field(default_factory=list)
    warranty_total: SliceKPI = field(default_factory=lambda: SliceKPI(0, 0, 0.0, 0.0, None, 0, 0.0, 0.0))
    warranty_suppliers: list[SupplierBlock] = field(default_factory=list)
    warranty_clients: list[ClientBlock] = field(default_factory=list)


def filter_bucket(df: pd.DataFrame, bucket: str) -> pd.DataFrame:
    return df.loc[df["_bucket"].eq(bucket)].copy()


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
    out = []
    for _, r in part.sort_values("_sale", ascending=False).iterrows():
        line = line_from_row(r)
        if bool(r.get("_warranty", False)):
            desc = line.description or ""
            if not desc.upper().startswith("[WARRANTY]"):
                line.description = f"[WARRANTY] {desc}".strip()
        out.append(line)
    return out


def _sorted_names(rows: pd.DataFrame, key: str) -> list:
    return sorted(
        rows[key].unique(),
        key=lambda s: (-float(rows.loc[rows[key] == s, "_sale"].sum()), str(s).casefold()),
    )


def build_leaves(
    rows: pd.DataFrame, lead_rows: pd.DataFrame, key: str
) -> list[EntityLeaf]:
    out: list[EntityLeaf] = []
    for name in _sorted_names(rows, key):
        part = rows[rows[key] == name]
        lead_part = lead_rows[lead_rows[key] == name]
        out.append(
            EntityLeaf(
                name=str(name),
                kpi=_kpi(part, lead_part),
                positions=_positions(part),
            )
        )
    return out


def build_suppliers(rows: pd.DataFrame, lead_rows: pd.DataFrame) -> list[SupplierBlock]:
    out: list[SupplierBlock] = []
    for name in _sorted_names(rows, "_supplier"):
        part = rows[rows["_supplier"] == name]
        lead_part = lead_rows[lead_rows["_supplier"] == name]
        out.append(
            SupplierBlock(
                name=str(name),
                kpi=_kpi(part, lead_part),
                clients=build_leaves(part, lead_part, "_client"),
            )
        )
    return out


def build_clients(rows: pd.DataFrame, lead_rows: pd.DataFrame) -> list[ClientBlock]:
    out: list[ClientBlock] = []
    for name in _sorted_names(rows, "_client"):
        part = rows[rows["_client"] == name]
        lead_part = lead_rows[lead_rows["_client"] == name]
        out.append(
            ClientBlock(
                name=str(name),
                kpi=_kpi(part, lead_part),
                suppliers=build_leaves(part, lead_part, "_supplier"),
            )
        )
    return out


def build_period_view(df: pd.DataFrame, title: str, start: date, end: date) -> PeriodView:
    rows = df.loc[report_mask(df, start, end)].copy()
    lead_rows = df.loc[lead_mask(df, start, end)].copy()
    w_rows = df.loc[warranty_mask(df, start, end)].copy()
    w_lead = df.loc[lead_mask(df, start, end) & df["_warranty"]].copy()
    return PeriodView(
        title=title,
        start=start,
        end=end,
        total=_kpi(rows, lead_rows),
        suppliers=build_suppliers(rows, lead_rows),
        clients=build_clients(rows, lead_rows),
        warranty_total=_kpi(w_rows, w_lead),
        warranty_suppliers=build_suppliers(w_rows, w_lead),
        warranty_clients=build_clients(w_rows, w_lead),
    )


def kpi_meta(kpi: SliceKPI) -> str:
    mp = f" · {kpi.margin_pct:.1f}%" if kpi.margin_pct is not None else ""
    return (
        f"{fmt_int(kpi.orders)} зак. · {fmt_int(kpi.lines)} поз. · "
        f"${fmt_money(kpi.revenue)} · маржа ${fmt_money(kpi.margin)}{mp} · "
        f"срок {fmt_days(kpi.lead_avg)} дн. · {fmt_transport_delta_html(kpi.transport_delta)}"
    )


def render_header_kpis(kpi: SliceKPI, spec: BucketSpec) -> str:
    return f"""
    <div class="kpis">
      <div class="highlight">
        <div class="label">Выполнено заказов</div>
        <div class="value">{fmt_int(kpi.orders)}</div>
        <div class="muted">{fmt_int(kpi.lines)} позиций {spec.title} с W</div>
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
        <div class="muted">USD · только {spec.title} ({spec.include_note})</div>
      </div>
    </div>
    """


def render_leaf_dropdown(
    leaf: EntityLeaf, table_id: str, kind_label: str, spec: BucketSpec
) -> str:
    """Мини-отчёт внутри поставщика/клиента."""
    body = (
        render_positions_table(leaf.positions, table_id)
        if leaf.positions
        else "<p class='empty'>Нет позиций</p>"
    )
    return f"""
<details class="leaf">
  <summary>
    <span class="ent-name">{html_escape(leaf.name)}</span>
    <span class="ent-meta">{kpi_meta(leaf.kpi)}</span>
  </summary>
  <div class="block-body">
    {render_header_kpis(leaf.kpi, spec)}
    <h4 class="leaf-h">Позиции · {html_escape(kind_label)}</h4>
    {body}
  </div>
</details>
"""


def render_supplier_dropdown(sup: SupplierBlock, prefix: str, spec: BucketSpec) -> str:
    clients = "\n".join(
        render_leaf_dropdown(c, f"{prefix}-c{i}", "клиент", spec)
        for i, c in enumerate(sup.clients)
    ) or "<p class='empty'>Нет клиентов</p>"
    return f"""
<details class="drop">
  <summary>
    <span class="ent-name">{html_escape(sup.name)}</span>
    <span class="ent-meta">{kpi_meta(sup.kpi)} · клиентов {fmt_int(len(sup.clients))}</span>
  </summary>
  <div class="block-body">
    {render_header_kpis(sup.kpi, spec)}
    <h3 class="section-h">Клиенты ({fmt_int(len(sup.clients))})</h3>
    {clients}
  </div>
</details>
"""


def render_client_dropdown(client: ClientBlock, prefix: str, spec: BucketSpec) -> str:
    suppliers = "\n".join(
        render_leaf_dropdown(s, f"{prefix}-s{i}", "поставщик", spec)
        for i, s in enumerate(client.suppliers)
    ) or "<p class='empty'>Нет поставщиков</p>"
    return f"""
<details class="drop">
  <summary>
    <span class="ent-name">{html_escape(client.name)}</span>
    <span class="ent-meta">{kpi_meta(client.kpi)} · поставщиков {fmt_int(len(client.suppliers))}</span>
  </summary>
  <div class="block-body">
    {render_header_kpis(client.kpi, spec)}
    <h3 class="section-h">Поставщики ({fmt_int(len(client.suppliers))})</h3>
    {suppliers}
  </div>
</details>
"""


def render_warranty_header_kpis(kpi: SliceKPI, spec: BucketSpec) -> str:
    return f"""
    <div class="kpis warranty-kpis">
      <div class="highlight">
        <div class="label">Warranty-кейсы</div>
        <div class="value">{fmt_int(kpi.orders)}</div>
        <div class="muted">{fmt_int(kpi.lines)} позиций {spec.title} со статусом WARRANTY</div>
      </div>
      <div>
        <div class="label">Сумма оригинала (выручка)</div>
        <div class="value">${fmt_money(kpi.revenue)}</div>
        <div class="muted">не добавляется к итогу периода повторно</div>
      </div>
      <div>
        <div class="label">Маржа оригинала</div>
        <div class="value">${fmt_money(kpi.margin)}</div>
        <div class="muted">как в выполнении · под риском гарантии</div>
      </div>
      <div>
        <div class="label">Сроки STK (оригинал)</div>
        <div class="value">{fmt_days(kpi.lead_avg)} дн.</div>
        <div class="muted">ср. W−Q · {fmt_int(kpi.lead_n)} поз.</div>
      </div>
    </div>
    """


def render_period(period: PeriodView, idx: int, spec: BucketSpec, *, opened: bool) -> str:
    open_attr = " open" if opened else ""
    range_s = f"{period.start.strftime('%d.%m.%Y')} – {period.end.strftime('%d.%m.%Y')}"
    suppliers_html = "\n".join(
        render_supplier_dropdown(s, f"p{idx}-sup-{i}", spec)
        for i, s in enumerate(period.suppliers)
    ) or "<p class='empty'>Нет поставщиков</p>"
    clients_html = "\n".join(
        render_client_dropdown(c, f"p{idx}-cli-{i}", spec)
        for i, c in enumerate(period.clients)
    ) or "<p class='empty'>Нет клиентов</p>"
    w_sup_html = "\n".join(
        render_supplier_dropdown(s, f"p{idx}-wsup-{i}", spec)
        for i, s in enumerate(period.warranty_suppliers)
    ) or "<p class='empty'>Нет warranty-кейсов у поставщиков</p>"
    w_cli_html = "\n".join(
        render_client_dropdown(c, f"p{idx}-wcli-{i}", spec)
        for i, c in enumerate(period.warranty_clients)
    ) or "<p class='empty'>Нет warranty-кейсов у клиентов</p>"
    w_meta = (
        f" · warranty {fmt_int(period.warranty_total.orders)} зак. / "
        f"${fmt_money(period.warranty_total.revenue)}"
        if period.warranty_total.lines
        else ""
    )
    return f"""
<details class="period"{open_attr}>
  <summary>
    <span class="period-title">{idx}. {html_escape(period.title)}</span>
    <span class="period-meta">{range_s}</span>
    <span class="period-meta">{fmt_int(period.total.orders)} зак. · ${fmt_money(period.total.revenue)} выр. · маржа ${fmt_money(period.total.margin)}{w_meta}</span>
  </summary>
  <div class="period-body">
    {render_header_kpis(period.total, spec)}

    <details class="section" open>
      <summary>1. По поставщикам ({fmt_int(len(period.suppliers))})</summary>
      <div class="block-body">
        <p class="section-note">Внутри поставщика — мини-отчёты по клиентам. Позиции со статусом WARRANTY помечены в описании.</p>
        {suppliers_html}
      </div>
    </details>

    <details class="section" open>
      <summary>2. По клиентам ({fmt_int(len(period.clients))})</summary>
      <div class="block-body">
        <p class="section-note">Внутри клиента — мини-отчёты по поставщикам.</p>
        {clients_html}
      </div>
    </details>

    <details class="section warranty-section"{" open" if period.warranty_total.lines else ""}>
      <summary>3. WARRANTY ({fmt_int(period.warranty_total.orders)} зак. · {fmt_int(period.warranty_total.lines)} поз. · ${fmt_money(period.warranty_total.revenue)})</summary>
      <div class="block-body">
        <p class="section-note warranty-note">
          Те же позиции, что уже учтены выше как выполненные (есть W).
          Здесь — отдельный срез по статусу WARRANTY: сумма оригинала <strong>не</strong> прибавляется к выручке периода повторно.
        </p>
        {render_warranty_header_kpis(period.warranty_total, spec)}
        <details class="section" open>
          <summary>3.1. Warranty · по поставщикам ({fmt_int(len(period.warranty_suppliers))})</summary>
          <div class="block-body">{w_sup_html}</div>
        </details>
        <details class="section" open>
          <summary>3.2. Warranty · по клиентам ({fmt_int(len(period.warranty_clients))})</summary>
          <div class="block-body">{w_cli_html}</div>
        </details>
      </div>
    </details>
  </div>
</details>
"""


def render_html(
    periods: list[PeriodView], source_name: str, as_of: date, spec: BucketSpec
) -> str:
    sections = "\n".join(
        render_period(p, i + 1, spec, opened=(i == 0)) for i, p in enumerate(periods)
    )
    return f"""<!doctype html>
<html lang="ru">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>FASTAIR — {spec.title} · поставщики и клиенты</title>
<style>
:root {{
  --navy:#04384d; --line:#c5d4da; --muted:#6b7c83; --zebra:#f4f8f9;
  --drop:#6b3f1f; --drop-bg:#f3e6d4; --drop-border:#d2a66a;
  --leaf:#3d4a22; --leaf-bg:#eef0e2; --leaf-border:#b7c08a;
  --section-bg:#e8f0f3;
  --warranty:#6b2430; --warranty-bg:#f7e8eb; --warranty-border:#c98a96;
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

/* Разделы 1/2 — холодный фон, но не как выпадающие сущности */
details.section {{
  border:1px solid #b7c8cf; border-radius:8px; margin:14px 0; background:#fff;
}}
details.section > summary {{
  cursor:pointer; padding:12px 14px; list-style:none; background:var(--section-bg);
  font-weight:700; color:var(--navy); font-size:15px;
}}
details.warranty-section {{
  border-color: var(--warranty-border);
}}
details.warranty-section > summary {{
  background: var(--warranty-bg);
  color: var(--warranty);
}}
.warranty-note {{ color: var(--warranty); }}
.warranty-kpis > div {{ background:#fdf6f7; border-color: var(--warranty-border); }}
.warranty-kpis > div.highlight {{ background: var(--warranty-bg); border-color: var(--warranty-border); }}
.warranty-kpis .value {{ color: var(--warranty); }}

/* Основные выпадающие (поставщик / клиент) — тёплый песочный, контраст к голубому */
details.drop {{
  border:1px solid var(--drop-border); border-radius:8px; margin:10px 0; background:#fff;
}}
details.drop > summary {{
  cursor:pointer; padding:11px 12px; font-weight:700; color:var(--drop);
  list-style:none; background:var(--drop-bg);
  display:flex; flex-wrap:wrap; gap:8px 14px; align-items:baseline;
}}

/* Вложенные мини-отчёты — оливковый, чтобы отличались от песочных */
details.leaf {{
  border:1px solid var(--leaf-border); border-radius:8px; margin:8px 0; background:#fff;
}}
details.leaf > summary {{
  cursor:pointer; padding:10px 12px; font-weight:700; color:var(--leaf);
  list-style:none; background:var(--leaf-bg);
  display:flex; flex-wrap:wrap; gap:8px 14px; align-items:baseline;
}}

details.section > summary::-webkit-details-marker,
details.drop > summary::-webkit-details-marker,
details.leaf > summary::-webkit-details-marker {{ display:none; }}
details.section > summary::before,
details.drop > summary::before,
details.leaf > summary::before {{ content:"▸ "; opacity:.7; }}
details.section[open] > summary::before,
details.drop[open] > summary::before,
details.leaf[open] > summary::before {{ content:"▾ "; }}

.block-body {{ padding:12px; background:#fff; }}
.ent-name {{ font-weight:700; }}
.ent-meta {{ font-size:12px; color:#5c5348; font-weight:600; }}
details.leaf .ent-meta {{ color:#4a5240; }}
.section-h {{ margin:16px 0 8px; font-size:14px; color:var(--navy); }}
.leaf-h {{ margin:8px 0; font-size:13px; color:var(--muted); font-weight:600; }}
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
    <h1>{spec.title} · поставщики и клиенты</h1>
    <div class="sub">Срез ТАЗ {html_escape(as_of.strftime('%d.%m.%Y'))} · {html_escape(source_name)} · только {spec.short}</div>
  </div>
  <div class="rules">
    <strong>Срез:</strong> только <strong>{spec.title}</strong> ({spec.include_note}), все клиенты.
    KIT / FEET / пустой Category в эти отчёты не входят (бакет «Прочее»).
    <br/>
    <strong>Попадание:</strong> заполненная W; без CANCELLED / REFUND / SCRAPPED / LOST;
    без типа поставки «Гарантийная поставка» (замена/возврат из ремонта).
    Статус <strong>WARRANTY</strong> на обычной продаже входит в выполнение и дублируется в разделе 3 — без повторного сложения выручки.
    FINISHED не фильтруем.
    <br/>
    <strong>Структура периода:</strong>
    1) по поставщикам → клиенты ·
    2) по клиентам → поставщики ·
    3) WARRANTY → те же кейсы отдельно.
    <br/>
    <strong>Метрики:</strong> заказы, позиции, выручка, маржа, срок STK (W−Q), Δ транспорта.
    Из сроков исключаем битые даты заказа Q (год &lt; 2020, напр. 1970-01-01).
  </div>
  {sections}
  <p class="hint">Открывайте HTML напрямую. Выпадающие списки — песочный/оливковый, чтобы отличались от голубой шапки отчёта.</p>
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


def generate(
    taz: Path,
    out_dir: Path,
    bucket: str,
    stem: str | None = None,
) -> Path:
    if bucket not in BUCKETS:
        raise SystemExit(f"Unknown bucket {bucket!r}; choose from {sorted(BUCKETS)}")
    spec = BUCKETS[bucket]
    as_of = infer_taz_end(taz)
    df = filter_bucket(prepare(load_taz(taz)), bucket)
    base = build_periods(df, as_of)
    periods = [build_period_view(df, p.title, p.start, p.end) for p in base]

    for p in periods:
        print(
            f"{p.title}: orders={p.total.orders} lines={p.total.lines} "
            f"rev={p.total.revenue:.0f} margin={p.total.margin:.0f} "
            f"suppliers={len(p.suppliers)} clients={len(p.clients)} "
            f"warranty={p.warranty_total.orders}/{p.warranty_total.lines} "
            f"w_rev={p.warranty_total.revenue:.0f}"
        )

    html = render_html(periods, taz.name, as_of, spec)
    out_dir.mkdir(parents=True, exist_ok=True)
    out = out_dir / f"{stem or spec.stem}.html"
    out.write_text(html, encoding="utf-8")
    print(f"wrote {out}")
    return out


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--bucket",
        choices=sorted(BUCKETS),
        default=CAT_EXPENDABLE,
        help="Категория: EXPENDABLE (consumable) или ROTABLE (GSE/TOOL)",
    )
    parser.add_argument(
        "--taz",
        type=Path,
        default=Path("/tmp/taz_history/ТАЗ 23.10.2026.xlsx"),
    )
    parser.add_argument("--out-dir", type=Path, default=Path("output"))
    parser.add_argument("--stem", default=None, help="Имя файла без .html")
    args = parser.parse_args()
    generate(args.taz, args.out_dir, args.bucket, args.stem)


if __name__ == "__main__":
    main()
