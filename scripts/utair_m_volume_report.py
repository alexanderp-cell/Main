#!/usr/bin/env python3
"""Объём работ ЮЛ М с клиентом Utair за 2025 и 2026.

Период строки — год столбца «ЗАКАЗ ВЗЯТ В РАБОТУ (ДАТА) ОТ КЛИЕНТА».
В объём входят только:
  Invoicer = М
  Client = Utair (без учёта регистра: Utair / UTair)
  Тип поставки = «Продажа от поставщика» или «Продажа со склада»
  статус не содержит cancel / refund / warranty

ТАЗ 2025 и ТАЗ 06.11.2026 пересекаются: поздний файл содержит те же заказы.
Каждая строка берётся один раз из среза 06.11.2026. Файл 2025 нужен, чтобы
проверить, что живые строки не потерялись и не попали в итог второй раз.
"""

from __future__ import annotations

import argparse
import html
import re
from pathlib import Path

import pandas as pd

BAD_STATUS = ("CANCEL", "REFUND", "WARRANTY")
KEEP_TYPES = {"продажа от поставщика", "продажа со склада"}
ORDER_DATE = "ЗАКАЗ ВЗЯТ В РАБОТУ (ДАТА) ОТ КЛИЕНТА"
# Счёт ещё не согласован с клиентом, в объём не входит.
HOLD_INVOICES = {"23102611345"}
# Заказ клиента в комментарии. Статус в ТАЗ — warranty, в 2026 объём он входит отдельно.
INCLUDE_CLIENT_ORDERS = {"P19807126"}
MONTHS = {
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


def norm_spaces(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\xa0", " ")).strip()


def norm_type(value) -> str:
    return norm_spaces(value).casefold()


def num(value) -> float:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return 0.0
    if isinstance(value, bool):
        return 0.0
    if isinstance(value, (int, float)):
        if isinstance(value, float) and pd.isna(value):
            return 0.0
        return float(value)
    text = str(value).strip().replace("\xa0", "").replace(" ", "").replace(",", ".")
    if not text or text.startswith("#"):
        return 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0


def money(value: float) -> str:
    return f"{value:,.2f}".replace(",", " ").replace(".", ",")


def integer(value: int) -> str:
    return f"{int(value):,}".replace(",", " ")


def load_orders(path: Path) -> pd.DataFrame:
    frame = pd.read_excel(path, sheet_name="ORDERS", engine="openpyxl")
    frame.columns = [" ".join(str(c).replace("\n", " ").split()) for c in frame.columns]
    return frame


def prepare(frame: pd.DataFrame, source: str) -> pd.DataFrame:
    client_col = "Client" if "Client" in frame.columns else "Customer"
    pn_col = "p/n" if "p/n" in frame.columns else "P/N"
    if "Продажная, итого" in frame.columns:
        sale_col = "Продажная, итого"
    else:
        sale_col = "Продажная, итого (без НДС)"
    if "Комментарии" in frame.columns:
        comment_col = "Комментарии"
    elif "Комментарий" in frame.columns:
        comment_col = "Комментарий"
    else:
        comment_col = None
    out = pd.DataFrame(
        {
            "source": source,
            "invoice": frame["Номер счета"].map(norm_spaces),
            "invoicer": frame["Invoicer"].map(norm_spaces),
            "client": frame[client_col].map(norm_spaces),
            "status": frame["Status"].map(lambda v: norm_spaces(v).upper()),
            "comment": frame[comment_col].map(norm_spaces) if comment_col else "",
            "delivery_type": frame["Тип поставки"].map(norm_spaces),
            "pn": frame[pn_col].map(lambda v: norm_spaces(v).upper()),
            "description": frame["DESCRIPTION"].map(norm_spaces),
            "qty": frame["QTY IN PO"].map(num),
            "category": frame["Category"].map(norm_spaces),
            "supplier": frame["Поставщик"].map(norm_spaces),
            "uuc": frame["UNIQUE UNIT CODE"].map(lambda v: norm_spaces(v).upper()),
            "order_date": pd.to_datetime(frame[ORDER_DATE], errors="coerce", dayfirst=True),
            "sale": frame[sale_col].map(num),
            "purchase": frame["Закупка, итого"].map(num),
        }
    )
    out["delivery_key"] = out["delivery_type"].map(norm_type)
    out["bad_status"] = out["status"].map(lambda s: any(token in s for token in BAD_STATUS))
    base_ok = (
        out["invoicer"].eq("М")
        & out["client"].str.casefold().eq("utair")
        & out["delivery_key"].isin(KEEP_TYPES)
        & out["order_date"].notna()
    )
    forced = out["comment"].map(comment_has_order) & base_ok
    out["kept"] = (base_ok & ~out["bad_status"]) | forced
    out["client_order"] = out["comment"].map(client_order_label)
    out["year"] = out["order_date"].dt.year
    out["month"] = out["order_date"].dt.month
    out["margin"] = out["sale"] - out["purchase"]
    out["real_invoice"] = ~out["invoice"].isin({"", "-", "—"})
    return out


def comment_has_order(value: str) -> bool:
    tokens = set(re.findall(r"[A-Za-z0-9]+", norm_spaces(value).upper()))
    return bool(tokens & INCLUDE_CLIENT_ORDERS)


def client_order_label(value: str) -> str:
    tokens = re.findall(r"[A-Za-z0-9]+", norm_spaces(value).upper())
    found = [token for token in tokens if token in INCLUDE_CLIENT_ORDERS]
    return found[0] if found else ""


def usable_uuc(value: str) -> str:
    text = norm_spaces(value).upper()
    if text in {"", "-", "—"}:
        return ""
    return text


def counterparts(row: pd.Series, later: pd.DataFrame, by_uuc: dict[str, list[int]]) -> pd.DataFrame:
    key = usable_uuc(row.uuc)
    if key and key in by_uuc:
        return later.loc[by_uuc[key]]
    same = later.index[(later["invoice"] == row.invoice) & (later["pn"] == row.pn)]
    return later.loc[same]


def audit_overlap(earlier: pd.DataFrame, later: pd.DataFrame) -> dict:
    early = earlier[earlier["kept"] & earlier["year"].eq(2025)].copy()
    by_uuc: dict[str, list[int]] = {}
    for idx, row in later.iterrows():
        key = usable_uuc(row.uuc)
        if key:
            by_uuc.setdefault(key, []).append(idx)

    missing = []
    retired_sale = 0.0
    retired_n = 0
    for _, row in early.iterrows():
        found = counterparts(row, later, by_uuc)
        if found.empty:
            missing.append(row.invoice)
            continue
        if not bool(found["kept"].any()):
            retired_n += 1
            retired_sale += float(row.sale)

    early_invoices = set(early["invoice"]) - {"", "-", "—"}
    later_kept = later[later["kept"]]
    later_invoices = set(later_kept["invoice"]) - {"", "-", "—"}
    shared = early_invoices & later_invoices
    shared_sale_early = float(early.loc[early["invoice"].isin(shared), "sale"].sum())
    return {
        "early_lines": int(len(early)),
        "early_sale": float(early["sale"].sum()),
        "missing": missing,
        "retired_n": retired_n,
        "retired_sale": retired_sale,
        "shared_invoices": len(shared),
        "shared_sale_early": shared_sale_early,
        "later_2026_in_early_file": int((earlier["kept"] & earlier["year"].eq(2026)).sum()),
    }


def period_stats(rows: pd.DataFrame) -> dict:
    real = rows[rows["real_invoice"]]
    return {
        "lines": int(len(rows)),
        "invoices": int(real["invoice"].nunique()),
        "placeholder_lines": int((~rows["real_invoice"]).sum()),
        "sale": float(rows["sale"].sum()),
        "purchase": float(rows["purchase"].sum()),
        "margin": float(rows["margin"].sum()),
        "date_min": rows["order_date"].min(),
        "date_max": rows["order_date"].max(),
    }


def esc(value) -> str:
    return html.escape(str(value))


def render_html(result: pd.DataFrame, audit: dict, excluded: pd.DataFrame) -> str:
    y2025 = result[result["year"].eq(2025)]
    y2026 = result[result["year"].eq(2026)]
    s2025 = period_stats(y2025)
    s2026 = period_stats(y2026)
    total_invoices = int(
        result.loc[result["real_invoice"], "invoice"].nunique()
    )
    both_years = sorted(
        set(y2025.loc[y2025["real_invoice"], "invoice"])
        & set(y2026.loc[y2026["real_invoice"], "invoice"])
    )

    def kpi_row(label: str, a, b, total) -> str:
        return (
            "<tr>"
            f"<td>{esc(label)}</td>"
            f"<td class='num'>{a}</td>"
            f"<td class='num'>{b}</td>"
            f"<td class='num total'>{total}</td>"
            "</tr>"
        )

    month_rows = []
    for year in (2025, 2026):
        part = result[result["year"].eq(year)]
        for month in range(1, 13):
            chunk = part[part["month"].eq(month)]
            if chunk.empty and not (year == 2025 and month == 1):
                if year == 2026 and month > 11:
                    continue
                if chunk.empty and year == 2025 and month == 1:
                    pass
                elif chunk.empty:
                    continue
            if chunk.empty and not (year == 2025 and month == 1):
                continue
            inv = int(chunk.loc[chunk["real_invoice"], "invoice"].nunique()) if len(chunk) else 0
            month_rows.append(
                "<tr>"
                f"<td>{year}</td>"
                f"<td>{MONTHS[month]}</td>"
                f"<td class='num'>{integer(len(chunk))}</td>"
                f"<td class='num'>{integer(inv)}</td>"
                f"<td class='num'>{money(float(chunk['sale'].sum())) if len(chunk) else '0,00'}</td>"
                f"<td class='num'>{money(float(chunk['purchase'].sum())) if len(chunk) else '0,00'}</td>"
                f"<td class='num'>{money(float(chunk['margin'].sum())) if len(chunk) else '0,00'}</td>"
                "</tr>"
            )

    def breakdown(frame: pd.DataFrame, column: str) -> str:
        rows = []
        for year in (2025, 2026):
            part = frame[frame["year"].eq(year)]
            grouped = (
                part.groupby(column, dropna=False)
                .agg(lines=("sale", "size"), sale=("sale", "sum"))
                .sort_values("sale", ascending=False)
            )
            for name, rec in grouped.iterrows():
                rows.append(
                    "<tr>"
                    f"<td>{year}</td>"
                    f"<td>{esc(name or '—')}</td>"
                    f"<td class='num'>{integer(rec['lines'])}</td>"
                    f"<td class='num'>{money(float(rec['sale']))}</td>"
                    "</tr>"
                )
        return "\n".join(rows)

    top_rows = []
    for year in (2025, 2026):
        grouped = (
            result[result["year"].eq(year)]
            .groupby("invoice")
            .agg(
                lines=("sale", "size"),
                sale=("sale", "sum"),
                pn=("pn", "first"),
                description=("description", "first"),
                supplier=("supplier", "first"),
                status=("status", lambda s: ", ".join(sorted(set(s)))),
            )
            .sort_values("sale", ascending=False)
            .head(5)
        )
        year_sale = float(result.loc[result["year"].eq(year), "sale"].sum())
        for invoice, rec in grouped.iterrows():
            share = float(rec["sale"]) / year_sale * 100 if year_sale else 0
            top_rows.append(
                "<tr>"
                f"<td>{year}</td>"
                f"<td>{esc(invoice)}</td>"
                f"<td class='num'>{money(float(rec['sale']))}</td>"
                f"<td class='num'>{share:.1f}%</td>"
                f"<td>{esc(rec['pn'])}</td>"
                f"<td>{esc(rec['description'])}</td>"
                f"<td>{esc(rec['supplier'])}</td>"
                f"<td>{esc(rec['status'])}</td>"
                "</tr>"
            )

    excluded_rows = []
    if len(excluded):
        grouped = (
            excluded.groupby(["year", "reason"])
            .agg(lines=("sale", "size"), sale=("sale", "sum"))
            .reset_index()
            .sort_values(["year", "sale"], ascending=[True, False])
        )
        for _, rec in grouped.iterrows():
            excluded_rows.append(
                "<tr>"
                f"<td>{int(rec['year'])}</td>"
                f"<td>{esc(rec['reason'])}</td>"
                f"<td class='num'>{integer(rec['lines'])}</td>"
                f"<td class='num'>{money(float(rec['sale']))}</td>"
                "</tr>"
            )

    span_note = ""
    if both_years:
        shown = ", ".join(both_years)
        span_note = (
            f" Номер счёта {esc(shown)} есть и в 2025, и в 2026: "
            "это разные строки с разными датами взятия в работу, каждая в своём году. "
            "В общем числе счетов такой номер посчитан один раз."
        )

    placeholder = result[~result["real_invoice"]]
    placeholder_note = ""
    if len(placeholder):
        bits = [
            f"{row.order_date.date().isoformat()} · {row.pn} · {money(float(row.sale))} USD · {row.status}"
            for row in placeholder.itertuples()
        ]
        placeholder_note = (
            " В 2025 есть "
            + integer(len(placeholder))
            + " строка без номера счёта (в ТАЗ стоит «-»): "
            + "; ".join(bits)
            + ". Она входит в позиции и в продажу, в число счетов не входит."
        )

    return f"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8">
<title>Utair · ЮЛ М · объём 2025 и 2026</title>
<style>
  body {{ font-family: "Segoe UI", Arial, sans-serif; color: #1c2430; margin: 32px auto; max-width: 1100px; line-height: 1.45; }}
  h1 {{ font-size: 28px; margin-bottom: 4px; }}
  h2 {{ font-size: 20px; margin-top: 32px; }}
  p, li {{ font-size: 15px; }}
  .muted {{ color: #5c6b7a; }}
  table {{ border-collapse: collapse; width: 100%; margin: 12px 0 8px; font-size: 14px; }}
  th, td {{ border-bottom: 1px solid #d5dde6; padding: 6px 8px; text-align: left; vertical-align: top; }}
  th {{ background: #1f3864; color: white; font-weight: 600; }}
  td.num, th.num {{ text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }}
  td.total {{ font-weight: 700; }}
  tr:nth-child(even) td {{ background: #f6f8fb; }}
  .note {{ background: #f4f7fb; border: 1px solid #d5dde6; padding: 12px 14px; }}
</style>
</head>
<body>
<h1>Utair · ЮЛ М · объём работ</h1>
<p class="muted">2025 год целиком и 2026 год по срез ТАЗ 06.11.2026. Суммы в USD, столбец «Продажная, итого» позднего файла.</p>

<div class="note">
  <strong>Как попала строка.</strong> Invoicer = М. Клиент только Utair. Тип поставки только «Продажа от поставщика» или «Продажа со склада». Статусы cancel, refund и warranty не входят. Год строки — «ЗАКАЗ ВЗЯТ В РАБОТУ (ДАТА) ОТ КЛИЕНТА».
  Валовая маржа здесь = продажа − закупка, без транспорта и transaction fee.
  Счёт 23102611345 снят с отчёта: он ещё окончательно не согласован.
  Заказ клиента P19807126 добавлен в 2026, хотя в ТАЗ у строки статус warranty. Это счёт 14052610285, HMU 442653, продажа 1 145 000 USD, дата взятия в работу 14 мая 2026.
</div>

<h2>Итог</h2>
<table>
  <thead>
    <tr><th></th><th class="num">2025</th><th class="num">2026, по 6 ноября</th><th class="num">Вместе</th></tr>
  </thead>
  <tbody>
    {kpi_row("Позиции", integer(s2025["lines"]), integer(s2026["lines"]), integer(s2025["lines"] + s2026["lines"]))}
    {kpi_row("Счета", integer(s2025["invoices"]), integer(s2026["invoices"]), integer(total_invoices))}
    {kpi_row("Продажа, USD", money(s2025["sale"]), money(s2026["sale"]), money(s2025["sale"] + s2026["sale"]))}
    {kpi_row("Закупка, USD", money(s2025["purchase"]), money(s2026["purchase"]), money(s2025["purchase"] + s2026["purchase"]))}
    {kpi_row("Валовая маржа, USD", money(s2025["margin"]), money(s2026["margin"]), money(s2025["margin"] + s2026["margin"]))}
  </tbody>
</table>
<p class="muted">
  2025: с {s2025["date_min"].date().isoformat()} по {s2025["date_max"].date().isoformat()}.
  Января 2025 в ТАЗ по Utair и ЮЛ М нет.
  2026: с {s2026["date_min"].date().isoformat()} по {s2026["date_max"].date().isoformat()}, это не полный год.
  {span_note}{placeholder_note}
</p>

<h2>По месяцу взятия в работу</h2>
<p class="muted">Число счетов в месяце — сколько номеров имеют хотя бы одну строку в этом месяце. Сумма по месяцам больше годового числа счетов, если один счёт растянут на несколько месяцев.</p>
<table>
  <thead>
    <tr>
      <th>Год</th><th>Месяц</th>
      <th class="num">Позиции</th><th class="num">Счета</th>
      <th class="num">Продажа, USD</th><th class="num">Закупка, USD</th><th class="num">Валовая маржа, USD</th>
    </tr>
  </thead>
  <tbody>
    {"".join(month_rows)}
  </tbody>
</table>

<h2>Тип поставки</h2>
<table>
  <thead><tr><th>Год</th><th>Тип поставки</th><th class="num">Позиции</th><th class="num">Продажа, USD</th></tr></thead>
  <tbody>{breakdown(result, "delivery_type")}</tbody>
</table>

<h2>Статус на срезе 06.11.2026</h2>
<table>
  <thead><tr><th>Год</th><th>Статус</th><th class="num">Позиции</th><th class="num">Продажа, USD</th></tr></thead>
  <tbody>{breakdown(result, "status")}</tbody>
</table>

<h2>Крупнейшие счета</h2>
<p class="muted">Пять счетов с наибольшей продажей внутри года. Доля — от продажи этого года.</p>
<table>
  <thead>
    <tr>
      <th>Год</th><th>Счёт</th><th class="num">Продажа, USD</th><th class="num">Доля</th>
      <th>P/N</th><th>Описание</th><th>Поставщик</th><th>Статус</th>
    </tr>
  </thead>
  <tbody>{"".join(top_rows)}</tbody>
</table>

<h2>Почему 2025 и 2026 не сложились дважды</h2>
<p>
  Файл «TA3 2025» и файл «ТАЗ 06.11.2026» описывают одну и ту же книгу заказов на разные даты.
  После фильтров в файле 2025 остаётся {integer(audit["early_lines"])} строк 2025 года на {money(audit["early_sale"])} USD.
  Общих номеров счетов с поздним файлом: {integer(audit["shared_invoices"])}.
  Если прибавить эти строки к срезу 06.11.2026, продажа 2025 задвоится на {money(audit["shared_sale_early"])} USD.
</p>
<p>
  Каждая отфильтрованная строка файла 2025 нашлась в файле 2026 (по UNIQUE UNIT CODE, а если его нет — по номеру счёта и P/N).
  Ненайденных строк: {integer(len(audit["missing"]))}.
  Из найденных {integer(audit["retired_n"])} строк на {money(audit["retired_sale"])} USD к 6 ноября 2026 уже не проходят фильтр: статус стал cancel, refund или warranty, либо сменились клиент и тип поставки.
  Строк с датой 2026 в файле 2025 нет ({integer(audit["later_2026_in_early_file"])}).
  Поэтому объём выше взят только из ТАЗ 06.11.2026 и разрезан по дате взятия в работу.
  Суммы в позднем срезе — актуальные: часть строк за год поменяла цену, часть разбита на несколько строк, часть закрыта.
</p>

<h2>Что рядом с Utair и ЮЛ М, но не вошло</h2>
<p class="muted">Те же клиент и Invoicer, дата взятия в работу в 2025 или 2026. Сюда же входит счёт 23102611345: он проходит фильтр, но ещё не согласован.</p>
<table>
  <thead><tr><th>Год</th><th>Причина</th><th class="num">Позиции</th><th class="num">Продажа, USD</th></tr></thead>
  <tbody>{"".join(excluded_rows)}</tbody>
</table>
</body>
</html>
"""


def exclusion_frame(later: pd.DataFrame) -> pd.DataFrame:
    base = later[
        later["invoicer"].eq("М")
        & later["client"].str.casefold().eq("utair")
        & later["year"].isin([2025, 2026])
        & later["order_date"].notna()
        & ~later["kept"]
    ].copy()

    def reason(row) -> str:
        if row.bad_status:
            return row.status
        if row.delivery_key not in KEEP_TYPES:
            return row.delivery_type or "пустой тип поставки"
        return "другой фильтр"

    base["reason"] = base.apply(reason, axis=1)
    return base


def write_xlsx(path: Path, result: pd.DataFrame, audit: dict, excluded: pd.DataFrame) -> None:
    y2025 = result[result["year"].eq(2025)]
    y2026 = result[result["year"].eq(2026)]
    s2025, s2026 = period_stats(y2025), period_stats(y2026)
    summary = pd.DataFrame(
        [
            ["Период", "2025", "2026 по 06.11.2026", "Вместе"],
            ["Позиции", s2025["lines"], s2026["lines"], s2025["lines"] + s2026["lines"]],
            [
                "Счета",
                s2025["invoices"],
                s2026["invoices"],
                int(result.loc[result["real_invoice"], "invoice"].nunique()),
            ],
            ["Продажа, USD", round(s2025["sale"], 2), round(s2026["sale"], 2), round(s2025["sale"] + s2026["sale"], 2)],
            ["Закупка, USD", round(s2025["purchase"], 2), round(s2026["purchase"], 2), round(s2025["purchase"] + s2026["purchase"], 2)],
            ["Валовая маржа, USD", round(s2025["margin"], 2), round(s2026["margin"], 2), round(s2025["margin"] + s2026["margin"], 2)],
            ["Правило периода", ORDER_DATE, "", ""],
            ["Invoicer", "М", "", ""],
            ["Клиент", "Utair", "", ""],
            ["Тип поставки", "Продажа от поставщика; Продажа со склада", "", ""],
            ["Исключённые статусы", "cancel, refund, warranty", "", ""],
            ["Снят вручную, ещё не согласован", "23102611345", "", ""],
            ["Добавлен вручную, заказ клиента", "P19807126 (счёт 14052610285, статус warranty)", "", ""],
            ["Источник сумм", "ТАЗ 06.11.2026, без повторного сложения файла 2025", "", ""],
            ["Строк 2025 в файле 2025 после фильтра", audit["early_lines"], "", ""],
            ["Общих счетов, которые нельзя сложить дважды", audit["shared_invoices"], "", ""],
            ["Продажа этих счетов в файле 2025, USD", round(audit["shared_sale_early"], 2), "", ""],
            ["Строк файла 2025, которые к 06.11.2026 вышли из фильтра", audit["retired_n"], "", ""],
            ["Их продажа в файле 2025, USD", round(audit["retired_sale"], 2), "", ""],
        ]
    )
    detail = result.sort_values(["year", "order_date", "invoice", "pn"]).copy()
    detail["order_date"] = detail["order_date"].dt.date
    detail = detail[
        [
            "year",
            "order_date",
            "invoice",
            "status",
            "delivery_type",
            "pn",
            "description",
            "qty",
            "category",
            "supplier",
            "sale",
            "purchase",
            "margin",
            "uuc",
            "client_order",
        ]
    ]
    detail.columns = [
        "Год",
        "Заказ взят в работу",
        "Номер счета",
        "Статус",
        "Тип поставки",
        "P/N",
        "Описание",
        "QTY",
        "Категория",
        "Поставщик",
        "Продажа, USD",
        "Закупка, USD",
        "Валовая маржа, USD",
        "UNIQUE UNIT CODE",
        "Заказ клиента",
    ]
    excluded_out = excluded.sort_values(["year", "reason", "invoice"]).copy()
    excluded_out["order_date"] = excluded_out["order_date"].dt.date
    excluded_out = excluded_out[
        ["year", "order_date", "invoice", "reason", "pn", "description", "sale", "supplier"]
    ]
    excluded_out.columns = [
        "Год",
        "Заказ взят в работу",
        "Номер счета",
        "Почему не вошло",
        "P/N",
        "Описание",
        "Продажа, USD",
        "Поставщик",
    ]
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Итог", index=False, header=False)
        detail.to_excel(writer, sheet_name="Позиции", index=False)
        excluded_out.to_excel(writer, sheet_name="Не вошло", index=False)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--taz-2025", type=Path, required=True)
    parser.add_argument("--taz-2026", type=Path, required=True)
    parser.add_argument("--html", type=Path, required=True)
    parser.add_argument("--xlsx", type=Path, required=True)
    args = parser.parse_args()

    earlier = prepare(load_orders(args.taz_2025), "2025")
    later = prepare(load_orders(args.taz_2026), "2026")
    audit = audit_overlap(earlier, later)
    if audit["missing"]:
        raise SystemExit(f"В файле 2026 не нашлись строки файла 2025: {audit['missing'][:10]}")
    if audit["later_2026_in_early_file"]:
        raise SystemExit("В файле 2025 есть строки 2026 года: их нужно разбирать отдельно, чтобы не задвоить")

    pool = later[later["kept"] & later["year"].isin([2025, 2026])].copy()
    held = pool[pool["invoice"].isin(HOLD_INVOICES)].copy()
    held["reason"] = "счёт ещё не согласован"
    result = pool[~pool["invoice"].isin(HOLD_INVOICES)].copy()
    excluded = pd.concat([exclusion_frame(later), held], ignore_index=True)
    args.html.parent.mkdir(parents=True, exist_ok=True)
    args.html.write_text(render_html(result, audit, excluded), encoding="utf-8")
    write_xlsx(args.xlsx, result, audit, excluded)

    for year in (2025, 2026):
        stats = period_stats(result[result["year"].eq(year)])
        print(
            year,
            "lines",
            stats["lines"],
            "invoices",
            stats["invoices"],
            "sale",
            round(stats["sale"], 2),
            "purchase",
            round(stats["purchase"], 2),
            "margin",
            round(stats["margin"], 2),
        )
    print("shared invoices", audit["shared_invoices"], "shared sale early", round(audit["shared_sale_early"], 2))
    print("retired", audit["retired_n"], round(audit["retired_sale"], 2))
    print("wrote", args.html)
    print("wrote", args.xlsx)


if __name__ == "__main__":
    main()
