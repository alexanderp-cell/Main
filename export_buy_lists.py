#!/usr/bin/env python3
"""
Рекомендации к выкупу со склада АТИ:
- 1 внутренний файл (формат как liquidity assessment)
- 3 клиентских файла (формат как склад АТИ: partno/serialno/...)
  по состоянию: сервисное / ансервис / неизвестное
"""
from __future__ import annotations

from collections import defaultdict
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import analyze_liquidity as al

OUT_DIR = Path("/workspace/output")
ART_DIR = Path("/opt/cursor/artifacts")

INTERNAL_NAME = "ATI_buyout_recommendations_internal.xlsx"
CLIENT_NAMES = {
    3: "ATI_buyout_client_serviceable.xlsx",
    2: "ATI_buyout_client_unserviceable.xlsx",
    1: "ATI_buyout_client_unknown.xlsx",
}
CLIENT_SHEET_TITLES = {
    3: "Сервисное состояние",
    2: "Ансервисное состояние",
    1: "Неизвестное состояние",
}


def is_buy_candidate(r: dict) -> bool:
    """Позиции, которые имеет смысл рекомендовать к выкупу."""
    g = r["liquidity_grade"]
    if g in {"A", "B"}:
        return True
    if g == "C":
        demand = (r.get("taz_orders") or 0) + (r.get("requests") or 0)
        rev = r.get("potential_revenue_usd") or 0
        # сильный C: повторный спрос или заметная потенц. выручка
        return demand >= 2 or rev >= 3000
    return False


def sort_key(r: dict):
    """Для выкупа важнее деньги, не буква ликвидности.

    Иначе расходка с оценкой A ($10–200) уезжает выше ротабля B на $50–100k.
    """
    return (
        -(r.get("potential_revenue_usd") or 0),
        al.GRADE_ORDER[r["liquidity_grade"]],
        -r["liquidity_score"],
        -r["qty"],
        r["partno"],
    )


def load_all_market_and_score():
    """Повторно собирает scored-позиции теми же правилами, что основной отчёт."""
    taz_files = [
        al.DATA / "TAZ_17.07.2026.xlsx",
        al.DATA / "TA3 2025.xlsx",
        al.DATA / "TA3-Архив 2024-01-26.xlsx",
    ]
    tuz_files = [
        al.DATA / "TUZ_17.07.2026.xlsx",
        al.DATA / "ТУЗ 2025 Jan - June.xlsx",
        al.DATA / "ТУЗ 6_19.xlsx",
    ]

    print("Loading market + ATI for buyout lists...")
    taz: list[al.Event] = []
    for path in taz_files:
        if path.exists():
            part = al.load_taz(path)
            print(f"  TAZ {path.name}: {len(part)}")
            taz.extend(part)

    tuz: list[al.Event] = []
    for path in tuz_files:
        if path.exists():
            part = al.load_tuz(path)
            print(f"  TUZ {path.name}: {len(part)}")
            tuz.extend(part)

    exp: list[al.Event] = []
    exp_xlsx = al.DATA / "EXPENDABLES.xlsx"
    if exp_xlsx.exists():
        part = al.load_exp(exp_xlsx)
        print(f"  EXP {exp_xlsx.name}: {len(part)}")
        exp.extend(part)
    for path in sorted(al.DATA.glob("*.csv")):
        part = al.load_exp_csv(path)
        print(f"  EXP CSV {path.name}: {len(part)}")
        exp.extend(part)

    ati = al.load_ati(al.DATA / "ATI.xlsx")
    print(f"  ATI rows: {len(ati)}")

    def event_global_key(e: al.Event):
        if e.kind == "order":
            inv = al.normalize_invoice(e.request_no) or (
                f"NO_INV|{al.client_key(e.client)}|{al.day_key(e.date)}|{round(e.qty or 0, 4)}"
            )
            return ("O", e.pn, inv)
        return (
            "R",
            e.pn,
            al.client_key(e.client),
            al.day_key(e.date),
            round(e.qty or 0, 4),
            round(e.price or 0, 2),
            (e.description or "")[:40],
        )

    seen = set()
    all_events = []
    for e in taz + tuz + exp:
        k = event_global_key(e)
        if k in seen:
            continue
        seen.add(k)
        all_events.append(e)

    by_pn, soft_to_pns, alt_to_pns = al.build_market_index(all_events)
    stock = al.aggregate_ati_stock(ati)
    scored = []
    for item in stock:
        market = al.resolve_market(item["pn"], by_pn, soft_to_pns, alt_to_pns)
        scored.append(al.build_row_from_stock(item, market))
    return ati, scored


def write_internal(path: Path, scored_buy: list[dict], ati_n: int):
    sec = {
        3: sorted([r for r in scored_buy if r["section"] == 3], key=sort_key),
        2: sorted([r for r in scored_buy if r["section"] == 2], key=sort_key),
        1: sorted([r for r in scored_buy if r["section"] == 1], key=sort_key),
    }
    wb = openpyxl.Workbook()
    ws0 = wb.active
    ws0.title = "0. Сводка"
    ws0["A1"] = "Рекомендации к выкупу — склад АТИ (внутренний)"
    ws0["A1"].font = Font(bold=True, size=16, name="Calibri")
    ws0["A2"] = (
        f"Дата: {datetime.now():%Y-%m-%d %H:%M} | "
        "Критерий: ликвидность A/B или сильный C (спрос ≥2 или потенц. выручка ≥ $3 000)"
    )
    ws0.merge_cells("A2:F2")

    total_rev = sum(r.get("potential_revenue_usd") or 0 for r in scored_buy)
    rows_sum = [
        ("Строк в исходном АТИ", ati_n),
        ("Рекомендовано позиций (P/N×состояние)", len(scored_buy)),
        ("Сервисное состояние", len(sec[3])),
        ("Ансервисное состояние", len(sec[2])),
        ("Неизвестное состояние", len(sec[1])),
        ("Суммарное кол-во Utair в рекомендациях, шт.", sum(r["qty"] for r in scored_buy)),
        ("Суммарная потенц. выручка (где есть цена), USD", round(total_rev, 2)),
        ("A / B / C среди рекомендаций", f"{sum(1 for r in scored_buy if r['liquidity_grade']=='A')} / "
         f"{sum(1 for r in scored_buy if r['liquidity_grade']=='B')} / "
         f"{sum(1 for r in scored_buy if r['liquidity_grade']=='C')}"),
    ]
    ws0["A4"] = "Сводка"
    ws0["A4"].font = Font(bold=True, size=13)
    for i, (k, v) in enumerate(rows_sum, 5):
        ws0.cell(i, 1, k)
        cell = ws0.cell(i, 2, v)
        cell.font = Font(bold=True)
        if "выручка" in k.lower() and isinstance(v, (int, float)):
            cell.number_format = '"$"#,##0.00'

    ws0["A14"] = "Методика"
    ws0["A14"].font = Font(bold=True, size=13)
    method = [
        "Сервисное: Condition не пустое и не US/NA (N, SV, OH, R, S, IT, …).",
        "Ансервис: Condition US или NA.",
        "Неизвестное: Condition пустое.",
        "В клиентские файлы попадают исходные строки склада (с serialno) по отобранным P/N и разделу состояния.",
        "Сортировка внутри листа: потенц. выручка → ликвидность → балл → qty "
        "(чтобы дорогие позиции не уезжали ниже расходки с буквой A).",
        "D (нет спроса) в рекомендации не входят.",
    ]
    for i, t in enumerate(method, 15):
        ws0.cell(i, 1, t)
        ws0.merge_cells(start_row=i, start_column=1, end_row=i, end_column=6)
    ws0.column_dimensions["A"].width = 70
    ws0.column_dimensions["B"].width = 22

    al.write_rows(wb.create_sheet("1. Сервисное"), sec[3], "2F5D9F")
    al.write_rows(wb.create_sheet("2. Ансервис"), sec[2], "B33B3B")
    al.write_rows(wb.create_sheet("3. Неизвестное"), sec[1], "6B4C9A")

    wsl = wb.create_sheet("Легенда")
    wsl["A1"] = "Ликвидность"
    wsl["A1"].font = Font(bold=True, size=13)
    for i, (g, name) in enumerate([("A", "Высокая"), ("B", "Средняя"), ("C", "Низкая")], 3):
        c = wsl.cell(i, 1, g)
        c.fill = al.GRADE_FILL[g]
        c.font = al.GRADE_FONT
        wsl.cell(i, 2, name)
    wsl["A7"] = "Этот файл — внутренний. Клиенту отправляются 3 отдельных файла в формате склада АТИ."
    wsl.column_dimensions["A"].width = 12
    wsl.column_dimensions["B"].width = 40

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    print(f"Saved internal: {path}")
    return {k: len(v) for k, v in sec.items()}


def write_client_ati(path: Path, ati_rows: list[dict], keys: set[tuple], title_note: str):
    """Файл как склад АТИ: partno, serialno, description, ata_chapter, ac_typ, condition, qty."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Лист2"
    headers = ["partno", "serialno", "description", "ata_chapter", "ac_typ", "condition", "qty"]
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
    thin = Border(
        left=Side(style="thin", color="D0D0D0"),
        right=Side(style="thin", color="D0D0D0"),
        top=Side(style="thin", color="D0D0D0"),
        bottom=Side(style="thin", color="D0D0D0"),
    )
    for col, h in enumerate(headers, 1):
        cell = ws.cell(1, col, h)
        cell.fill = header_fill
        cell.font = header_font
        cell.border = thin
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # сохраняем порядок исходного склада
    n = 0
    for r in ati_rows:
        sec = al.condition_section(r.get("condition") or "")
        key = (r["pn"], sec)
        if key not in keys:
            continue
        vals = [
            r.get("partno") or "",
            r.get("serialno") or "",
            r.get("description") or "",
            r.get("ata") or "",
            r.get("ac_typ") or "",
            r.get("condition") or "",
            r.get("qty") if r.get("qty") is not None else 1,
        ]
        for col, val in enumerate(vals, 1):
            cell = ws.cell(n + 2, col, val)
            cell.border = thin
            cell.alignment = Alignment(vertical="center")
        n += 1

    widths = [18, 16, 36, 10, 12, 12, 8]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    if n:
        ws.auto_filter.ref = f"A1:G{n + 1}"

    # скрытый/служебный комментарий на отдельном листе не нужен клиенту —
    # только формат склада. title_note пишем в свойствах через doc props? skip.
    _ = title_note

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    print(f"Saved client ({n} rows): {path}")
    return n


def main():
    ati, scored = load_all_market_and_score()
    buy = [r for r in scored if is_buy_candidate(r)]
    print(
        f"Buy candidates: {len(buy)} "
        f"(svc={sum(1 for r in buy if r['section']==3)}, "
        f"us={sum(1 for r in buy if r['section']==2)}, "
        f"unk={sum(1 for r in buy if r['section']==1)})"
    )

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    ART_DIR.mkdir(parents=True, exist_ok=True)

    internal = OUT_DIR / INTERNAL_NAME
    counts = write_internal(internal, buy, len(ati))
    art_internal = ART_DIR / INTERNAL_NAME
    openpyxl.load_workbook(internal).save(art_internal)
    print(f"Saved internal copy: {art_internal}")

    for section, fname in CLIENT_NAMES.items():
        keys = {(r["pn"], r["section"]) for r in buy if r["section"] == section}
        path = OUT_DIR / fname
        n = write_client_ati(path, ati, keys, CLIENT_SHEET_TITLES[section])
        openpyxl.load_workbook(path).save(ART_DIR / fname)
        print(f"  section {section}: {len(keys)} PN-groups, {n} warehouse lines → {fname}")

    print("Done.", counts)


if __name__ == "__main__":
    main()
