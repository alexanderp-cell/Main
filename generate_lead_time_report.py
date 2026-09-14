#!/usr/bin/env python3
"""Average delivery lead time by client for a delivery-date window.

Rules:
- Status must be FINISHED (e.g. «4 FINISHED»).
- Column S («Lead time») must be exactly STK (other lead-time values excluded).
- Keep rows whose actual delivery date (W) falls in [start, end] inclusive.
- Lead time days = W − Q (order taken into work).
- Require both dates; drop negative lead times.
- Split ROTABLE / EXPENDABLE averages and counts per client.
"""

from __future__ import annotations

import argparse
import zipfile
from datetime import date
from pathlib import Path

import pandas as pd

COL_STATUS = "Status"
COL_CUSTOMER = "Customer"
COL_CATEGORY = "Category"
COL_LEAD_TIME = "Lead time"  # column S
COL_ORDER_DATE = "ЗАКАЗ ВЗЯТ В РАБОТУ (ДАТА) ОТ КЛИЕНТА"
COL_DELIVERY = "ФАКТИЧЕСКАЯ ДАТА ПОСТАВКИ (СОГЛАСНО УСЛОВИЯМ ПОСТАВКИ)"

CAT_ROTABLE = "ROTABLE"
CAT_EXPENDABLE = "EXPENDABLE"


def load_taz(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=0)
    df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]
    return df


def prepare(df: pd.DataFrame, start: date, end: date) -> pd.DataFrame:
    out = df.copy()
    out["_q"] = pd.to_datetime(out[COL_ORDER_DATE], errors="coerce")
    out["_w"] = pd.to_datetime(out[COL_DELIVERY], errors="coerce")
    out["_days"] = (out["_w"] - out["_q"]).dt.days
    out["_cat"] = out[COL_CATEGORY].astype(str).str.strip().str.upper()
    out["_client"] = out[COL_CUSTOMER].astype(str).str.strip()
    out["_status"] = out[COL_STATUS].astype(str).str.strip().str.upper()
    out["_lead"] = out[COL_LEAD_TIME].astype(str).str.strip().str.upper()

    mask = (
        out["_status"].str.contains("FINISHED", na=False)
        & out["_lead"].eq("STK")
        & out["_w"].notna()
        & out["_q"].notna()
        & (out["_w"].dt.date >= start)
        & (out["_w"].dt.date <= end)
        & (out["_days"] >= 0)
        & out["_client"].ne("")
        & out["_client"].str.lower().ne("nan")
        & out["_cat"].isin({CAT_ROTABLE, CAT_EXPENDABLE})
    )
    return out.loc[mask].copy()


def agg_side(series: pd.Series) -> tuple[float | None, int]:
    if series.empty:
        return None, 0
    return float(series.mean()), int(len(series))


def build_table(rows: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    clients = sorted(rows["_client"].unique(), key=lambda s: s.casefold())
    records = []
    for client in clients:
        part = rows[rows["_client"] == client]
        r_avg, r_n = agg_side(part.loc[part["_cat"] == CAT_ROTABLE, "_days"])
        e_avg, e_n = agg_side(part.loc[part["_cat"] == CAT_EXPENDABLE, "_days"])
        records.append(
            {
                "client": client,
                "rotable_avg": r_avg,
                "rotable_n": r_n,
                "expendable_avg": e_avg,
                "expendable_n": e_n,
            }
        )
    table = pd.DataFrame(records)

    r_all = rows.loc[rows["_cat"] == CAT_ROTABLE, "_days"]
    e_all = rows.loc[rows["_cat"] == CAT_EXPENDABLE, "_days"]
    r_avg, r_n = agg_side(r_all)
    e_avg, e_n = agg_side(e_all)
    overall = {
        "avg": float(rows["_days"].mean()) if len(rows) else None,
        "median": float(rows["_days"].median()) if len(rows) else None,
        "rotable_avg": r_avg,
        "rotable_n": r_n,
        "expendable_avg": e_avg,
        "expendable_n": e_n,
        "clients": int(table.shape[0]),
        "positions": int(len(rows)),
    }
    return table, overall


def fmt_avg(value: float | None) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return "—"
    return f"{float(value):.1f}"


def fmt_n(value: int) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)) or int(value) == 0:
        return "—"
    return str(int(value))


def html_escape(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def empty_num_cell() -> str:
    return '<td class="num empty" data-value="">—</td>'


def avg_cell(value: float | None, count: int) -> str:
    if count == 0 or value is None or (isinstance(value, float) and pd.isna(value)):
        return empty_num_cell()
    shown = f"{float(value):.1f}"
    return f'<td class="num" data-value="{shown}">{shown}</td>'


def n_cell(count: int) -> str:
    if not count:
        return empty_num_cell()
    return f'<td class="num" data-value="{int(count)}">{int(count)}</td>'


def render_html(table: pd.DataFrame, overall: dict, source_name: str) -> str:
    rows_html = []
    for rec in table.itertuples(index=False):
        client = html_escape(rec.client)
        sort_key = html_escape(rec.client.casefold())
        r_n = int(rec.rotable_n or 0)
        e_n = int(rec.expendable_n or 0)
        rows_html.append(
            f'<tr><td class="client" data-value="{sort_key}">{client}</td>'
            f"{avg_cell(rec.rotable_avg, r_n)}{n_cell(r_n)}"
            f"{avg_cell(rec.expendable_avg, e_n)}{n_cell(e_n)}</tr>"
        )

    o_r = fmt_avg(overall["rotable_avg"])
    o_e = fmt_avg(overall["expendable_avg"])
    o_avg = fmt_avg(overall["avg"])
    o_med = fmt_avg(overall["median"])

    return f"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>Средний срок поставки — лето 2026</title>
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
.wrap {{ max-width:980px; margin:0 auto; padding:28px 20px 64px; }}
.brand {{ color:#fff; font-size:13px; letter-spacing:.12em; text-transform:uppercase;
  background:var(--cyan); color:var(--navy); display:inline-block; padding:6px 10px; font-weight:700; margin-bottom:14px; }}
h1 {{ color:#fff; font-size:28px; margin:0 0 6px; }}
.sub {{ color:#d5fbff; margin:0 0 18px; font-size:14px; }}
.card {{ background:var(--card); border:1px solid var(--line); border-radius:8px; padding:18px; }}
.kpis {{ display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:10px; margin-bottom:18px; }}
.kpis > div {{ background:#f7fbfc; border:1px solid #d7e2e6; border-radius:8px; padding:12px 14px; }}
.kpis > div.highlight {{ background:#e8f6f8; border-color:#9ed7e0; }}
.label {{ font-size:11px; text-transform:uppercase; letter-spacing:.04em; color:#757575; }}
.value {{ font-size:22px; font-weight:700; margin-top:4px; color:var(--navy); font-variant-numeric:tabular-nums; }}
.muted {{ color:var(--muted); font-size:12px; margin-top:4px; }}
table {{ width:100%; border-collapse:collapse; font-size:14px; }}
th, td {{ border-bottom:1px solid #e8e8e8; padding:10px 8px; text-align:left; vertical-align:middle; }}
th {{
  font-size:12px; text-transform:uppercase; letter-spacing:.03em; color:#fff;
  background:var(--navy); cursor:pointer; user-select:none; white-space:nowrap;
}}
th:hover {{ background:#03425a; }}
th .arrow {{ opacity:.45; margin-left:6px; font-size:11px; }}
th.sorted .arrow {{ opacity:1; }}
td.num {{ text-align:right; font-variant-numeric:tabular-nums; white-space:nowrap; }}
td.empty {{ color:#9a9a9a; }}
tbody#dataBody tr:nth-child(even) {{ background:var(--zebra); }}
tbody#dataBody tr:hover {{ background:#eef7f9; }}
tbody.summary-top td, tfoot td {{ background:#e8f2f5; font-weight:700; border-top:2px solid var(--navy); }}
.hint {{ margin-top:14px; color:var(--muted); font-size:12px; }}
@media (max-width:800px) {{ .kpis {{ grid-template-columns:repeat(2,minmax(0,1fr)); }} }}
</style>
</head>
<body>
<div class="wrap">
  <div class="brand">not so fastair</div>
  <h1>Средний срок поставки</h1>
  <div class="sub">Лето 2026 · FINISHED · Lead time = STK · факт поставки июнь–август · клик по заголовку — сортировка</div>

  <section class="card">
    <div class="kpis">
      <div class="highlight"><div class="label">Средний по всем клиентам</div><div class="value">{o_avg} дн.</div><div class="muted">медиана {o_med} · все категории</div></div>
      <div><div class="label">Ротабл</div><div class="value">{o_r} дн.</div><div class="muted">{overall['rotable_n']} поз.</div></div>
      <div><div class="label">Расходка</div><div class="value">{o_e} дн.</div><div class="muted">{overall['expendable_n']} поз.</div></div>
      <div><div class="label">Клиентов</div><div class="value">{overall['clients']}</div><div class="muted">{overall['positions']} позиций в выборке</div></div>
    </div>

    <table id="leadTable">
      <thead>
        <tr>
          <th data-type="str" data-col="0">Клиент <span class="arrow">↕</span></th>
          <th data-type="num" data-col="1">Ротабл, дн. <span class="arrow">↕</span></th>
          <th data-type="num" data-col="2">Ротабл n <span class="arrow">↕</span></th>
          <th data-type="num" data-col="3">Расходка, дн. <span class="arrow">↕</span></th>
          <th data-type="num" data-col="4">Расходка n <span class="arrow">↕</span></th>
        </tr>
      </thead>
      <tbody class="summary-top">
        <tr>
          <td>Средний по всем клиентам</td>
          <td class="num">{o_r}</td>
          <td class="num">{overall['rotable_n']}</td>
          <td class="num">{o_e}</td>
          <td class="num">{overall['expendable_n']}</td>
        </tr>
      </tbody>
      <tbody id="dataBody">
        {''.join(rows_html)}
      </tbody>
      <tfoot>
        <tr>
          <td>Средний по всем клиентам</td>
          <td class="num">{o_r}</td>
          <td class="num">{overall['rotable_n']}</td>
          <td class="num">{o_e}</td>
          <td class="num">{overall['expendable_n']}</td>
        </tr>
      </tfoot>
    </table>

    <p class="hint">
      Срок = факт. дата поставки (столбец W) − дата взятия в работу (столбец Q), в днях.
      Выборка: статус FINISHED, столбец S (Lead time) = STK, факт поставки в июне–августе 2026,
      обе даты заполнены, срок ≥ 0. n = число позиций в среднем. Источник: {html_escape(source_name)}.
      Нажмите на заголовок столбца, чтобы отсортировать.
    </p>
  </section>
</div>
<script>
(function () {{
  const table = document.getElementById('leadTable');
  const tbody = document.getElementById('dataBody');
  const headers = [...table.tHead.rows[0].cells];
  let sortCol = 0;
  let asc = true;

  function cellValue(td, type) {{
    const raw = td.getAttribute('data-value');
    if (raw === '' || raw === null) return type === 'num' ? Number.POSITIVE_INFINITY : '';
    return type === 'num' ? Number(raw) : raw;
  }}

  function sortBy(col, type) {{
    if (sortCol === col) asc = !asc;
    else {{ sortCol = col; asc = true; }}

    headers.forEach((th, i) => {{
      th.classList.toggle('sorted', i === col);
      th.querySelector('.arrow').textContent = i === col ? (asc ? '↑' : '↓') : '↕';
    }});

    const rows = [...tbody.rows];
    rows.sort((a, b) => {{
      const av = cellValue(a.cells[col], type);
      const bv = cellValue(b.cells[col], type);
      let cmp = 0;
      if (type === 'num') cmp = av - bv;
      else cmp = String(av).localeCompare(String(bv), 'ru');
      if (type === 'num') {{
        const aEmpty = !isFinite(av);
        const bEmpty = !isFinite(bv);
        if (aEmpty !== bEmpty) return aEmpty ? 1 : -1;
      }}
      return asc ? cmp : -cmp;
    }});
    rows.forEach(r => tbody.appendChild(r));
  }}

  headers.forEach(th => {{
    th.addEventListener('click', () => sortBy(Number(th.dataset.col), th.dataset.type));
  }});
}})();
</script>
</body>
</html>
"""


def render_markdown(table: pd.DataFrame, overall: dict) -> str:
    lines = [
        "# Средний срок поставки — лето 2026",
        "",
        "Заказы со статусом **FINISHED**, столбец S (**Lead time**) = **STK**, "
        "факт поставки в **июнь–август 2026**. Дни = факт поставки (W) − взятие в работу (Q).",
        "",
        "## Сводка",
        "",
        f"- **Средний по всем клиентам:** {fmt_avg(overall['avg'])} дн. (медиана {fmt_avg(overall['median'])} · все категории)",
        f"- **Ротабл:** {fmt_avg(overall['rotable_avg'])} дн. ({overall['rotable_n']} поз.)",
        f"- **Расходка:** {fmt_avg(overall['expendable_avg'])} дн. ({overall['expendable_n']} поз.)",
        f"- **Клиентов:** {overall['clients']} ({overall['positions']} позиций в выборке)",
        "",
        "## По клиентам",
        "",
        "| Клиент | Ротабл, дн. | Ротабл n | Расходка, дн. | Расходка n |",
        "| --- | --- | --- | --- | --- |",
        f"| Средний по всем клиентам | {fmt_avg(overall['rotable_avg'])} | {overall['rotable_n']} | {fmt_avg(overall['expendable_avg'])} | {overall['expendable_n']} |",
    ]
    for rec in table.itertuples(index=False):
        lines.append(
            f"| {rec.client} | {fmt_avg(rec.rotable_avg)} | {fmt_n(rec.rotable_n)} | "
            f"{fmt_avg(rec.expendable_avg)} | {fmt_n(rec.expendable_n)} |"
        )
    lines.append("")
    return "\n".join(lines)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--taz",
        type=Path,
        default=Path("/tmp/taz_history/ТАЗ 11.09.2026.xlsx"),
    )
    parser.add_argument("--out-dir", type=Path, default=Path("output"))
    parser.add_argument("--start", default="2026-06-01")
    parser.add_argument("--end", default="2026-08-31")
    args = parser.parse_args()

    start = date.fromisoformat(args.start)
    end = date.fromisoformat(args.end)

    df = load_taz(args.taz)
    rows = prepare(df, start, end)
    table, overall = build_table(rows)

    # Sanity: Азур
    azur = table[table["client"].str.contains("Азур", case=False, na=False)]
    if not azur.empty:
        a = azur.iloc[0]
        print(
            f"Азур Эйр: rotable {a.rotable_avg:.1f}/{a.rotable_n}, "
            f"expendable {a.expendable_avg:.1f}/{a.expendable_n}, "
            f"total n={a.rotable_n + a.expendable_n}"
        )

    print(
        f"overall avg={overall['avg']:.1f} rot={overall['rotable_n']} "
        f"exp={overall['expendable_n']} clients={overall['clients']} n={overall['positions']}"
    )

    args.out_dir.mkdir(parents=True, exist_ok=True)
    html = render_html(table, overall, args.taz.name)
    md = render_markdown(table, overall)

    ascii_html = args.out_dir / "lead_time_summer_2026.html"
    ascii_md = args.out_dir / "lead_time_summer_2026.md"
    ru_html = args.out_dir / "средний_срок_поставки_лето_2026.html"
    ascii_html.write_text(html, encoding="utf-8")
    ru_html.write_text(html, encoding="utf-8")
    ascii_md.write_text(md, encoding="utf-8")

    zip_path = args.out_dir / "lead_time_summer_2026.zip"
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("lead_time_summer_2026.html", html)
    ru_zip = args.out_dir / "средний_срок_поставки_лето_2026.zip"
    with zipfile.ZipFile(ru_zip, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("средний_срок_поставки_лето_2026.html", html)

    print(f"wrote {ascii_html}")
    print(f"wrote {ascii_md}")


if __name__ == "__main__":
    main()
