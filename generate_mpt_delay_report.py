#!/usr/bin/env python3
"""HTML-отчёт: задержки получения МПТ по ROTABLE (все клиенты кроме Белавиа).

Источники:
  · ТАЗ — Category=ROTABLE, Status ∈ {1 NOT PAID, 2 PAID, 6 TROUBLE}, возраст >7 дн.
  · Брокерские → только вкладка «Списки на МПТ»
  · Перечни зап. частей 0% НДС → только листы *2025-2026 (ФЭИ/ЛР/МЕТЕОР)

Одна тревога на заказ (приоритет):
  ipc — Нет IPC / IPC не будет получен
  4 — МПТ есть в Перечне, но в ТАЗ ещё не «Есть МПТ»
  1 — нет заявки в Списках
  2 — вырезка всё ещё «запрошена», возраст >7 дн.
  3 — нет МПТ в Перечне, возраст >14 дн.

Разделы отчёта — по клиентам; внутри клиента заказы от ранних к свежим.
Исключение: МПТ в Перечне И в ТАЗ «Есть МПТ» → не кладём.
«Закупка на склад» — в общем отчёте с пометкой.
"""

from __future__ import annotations

import argparse
import html
import re
import zipfile
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

import pandas as pd

from generate_report import (
    COL_COMMENT,
    COL_CUSTOMER,
    COL_DESC,
    COL_INVOICE,
    COL_INVOICER,
    COL_ORDER_DATE,
    COL_PN,
    COL_STATUS,
    load_taz,
    normalize_invoice,
    parse_date,
)

COL_CATEGORY = "Category"
COL_MPT = "МПТ"
COL_MPT_ACTION = "Действия по МПТ"
COL_AC_TYPE = "ТИП ВС (Продажи)"
COL_ALT_PN = "ALT P/N"

OPEN_STATUSES = {"1 NOT PAID", "2 PAID", "6 TROUBLE"}
EXCLUDED_CUSTOMERS = {"Белавиа"}
WAREHOUSE_MARKERS = {"закупка на склад"}

INVOICER_TO_SHEET = {
    "ФЭ": "ФЭИ 2025-2026",
    "ФЭИ": "ФЭИ 2025-2026",
    "М": "МЕТЕОР 2025-2026",
    "МЕТЕОР": "МЕТЕОР 2025-2026",
    "METEOR": "МЕТЕОР 2025-2026",
    "ЛР": "ЛР 2025-2026",
    "LR": "ЛР 2025-2026",
}

# Normalized extract / MPT pipeline stages from «Статус по вырезке»
STAGE_REQUESTED = "вырезка_запрошена"
STAGE_RECEIVED = "вырезка_получена"
STAGE_MPT_DONE = "мпт_получен"
STAGE_OTHER = "прочее"
STAGE_EMPTY = "пусто"


def esc(value: Any) -> str:
    return html.escape("" if value is None else str(value))


def norm_text(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).strip()
    text = re.sub(r"\s+", " ", text)
    if text.lower() in {"nan", "none", "nat", "-"}:
        return ""
    return text


def norm_pn(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if float(value).is_integer():
            return str(int(value))
        return str(value).strip()
    text = str(value).strip().upper().replace(" ", "")
    if text.endswith(".0") and text[:-2].replace("-", "").isalnum():
        # keep as-is unless pure integer float artifact
        pass
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    return text


def norm_client(value: Any) -> str:
    raw = norm_text(value)
    key = raw.casefold()
    aliases = {
        "ютэйр": "Utair",
        "ютейр": "Utair",
        "utair": "Utair",
        "аэрофлот-техникс": "Аэрофлот Техникс",
        "аэрофлот техникс": "Аэрофлот Техникс",
        "атран": "АТРАН",
        "тарп авиэйшн": "ТАРП Авиэйшн",
        "компонентс (utg)": "UTG",
        "utg": "UTG",
    }
    return aliases.get(key, raw)


def normalize_extract_stage(status: str) -> str:
    s = norm_text(status).casefold().replace("ё", "е")
    if not s:
        return STAGE_EMPTY
    if "мпт получен" in s or s.startswith("мпт получ"):
        return STAGE_MPT_DONE
    if "ожидаем мпт" in s:
        return STAGE_OTHER
    if "запрос" in s:
        return STAGE_REQUESTED
    if "получил" in s:
        return STAGE_RECEIVED
    if any(k in s for k in ("почте", "рто", "диске", "на диске")):
        return STAGE_RECEIVED  # extract exists somewhere / in process after request
    return STAGE_OTHER


def stage_label(stage: str) -> str:
    return {
        STAGE_REQUESTED: "вырезка запрошена",
        STAGE_RECEIVED: "вырезку получили",
        STAGE_MPT_DONE: "мпт получен",
        STAGE_OTHER: "прочее",
        STAGE_EMPTY: "—",
    }.get(stage, stage)


def is_warehouse(customer: str) -> bool:
    return customer.casefold() in WAREHOUSE_MARKERS or "склад" in customer.casefold()


def utair_no_mpt_needed(comment: Any) -> bool:
    """Utair DAT MOW / DAT TEHRAN — МПТ не требуется."""
    s = re.sub(r"\s+", " ", norm_text(comment).upper())
    return "DAT MOW" in s or "DAT TEHRAN" in s


def ac_tokens(value: str) -> set[str]:
    """Aircraft-type tokens with common TAZ ↔ Перечень aliases."""
    s = norm_text(value).upper()
    if not s:
        return set()
    s = s.replace("САМОЛЕТЫ ТИПА", " ").replace("САМОЛЕТ ТИПА", " ")
    s = s.replace("САМОЛЕТЫ ТИПОВ", " ").replace("САМОЛЕТОВ ТИПА", " ")
    s = re.sub(r"[^A-Z0-9]+", " ", s)
    parts = [p for p in s.split() if len(p) >= 2]
    extra: set[str] = set(parts)
    joined = "".join(parts)

    for p in list(parts):
        if p.startswith("B") and p[1:].isdigit():
            extra.add(p[1:])  # B737 → 737
        if p.isdigit() and len(p) in {3, 4}:
            extra.add("B" + p)

    # Airbus family
    if any(x in extra or x in joined for x in ("A32S", "A320", "A321", "A319", "A318")):
        extra.update({"A320", "A321", "A319", "A318", "A32S", "320", "321", "319", "318"})

    # Boeing 737 family: B737 / B738 (=737-800) / B735·B7385 (=737-500) / CL
    if any(
        x in extra or x in joined
        for x in ("737", "B737", "B738", "738", "B735", "735", "B7385", "7385", "737CL", "B737CL", "737800", "737500")
    ):
        extra.update(
            {
                "737",
                "B737",
                "B738",
                "738",
                "B735",
                "735",
                "B7385",
                "7385",
                "B737CL",
                "737CL",
                "737800",
                "B737800",
                "737500",
                "B737500",
            }
        )

    # Boeing 767 family: B762 (=767-200), B763 (=767-300), …
    if any(x in extra or x in joined for x in ("767", "B767", "B762", "762", "B763", "763", "767200", "767300")):
        extra.update(
            {
                "767",
                "B767",
                "B762",
                "762",
                "B763",
                "763",
                "767200",
                "B767200",
                "767300",
                "B767300",
            }
        )

    # Boeing 777
    if any(x in extra or x in joined for x in ("777", "B777")):
        extra.update({"777", "B777"})

    # Boeing 757
    if any(x in extra or x in joined for x in ("757", "B757")):
        extra.update({"757", "B757"})

    return {p for p in extra if p}


def ac_types_compatible(taz_ac: str, perm_ac: str) -> str:
    """Return match|mismatch|unknown."""
    a, b = ac_tokens(taz_ac), ac_tokens(perm_ac)
    if not a or not b:
        return "unknown"
    if a & b:
        return "match"
    # substring fallback on compact forms
    ca, cb = "".join(sorted(a)), "".join(sorted(b))
    for t in a:
        if t in perm_ac.upper().replace(" ", ""):
            return "match"
    return "mismatch"


@dataclass
class ListHit:
    client: str
    pn: str
    status_raw: str
    stage: str
    batch: str
    matched_via: str  # pn|alt


@dataclass
class PermHit:
    sheet: str
    pn: str
    description: str
    ac_type: str
    doc: str
    matched_via: str  # pn|alt
    ac_match: str  # match|mismatch|unknown


@dataclass
class MptRow:
    invoice: str
    client: str
    pn: str
    alt_pn: str
    description: str
    status: str
    invoicer: str
    ac_type: str
    order_date: date | None
    age_days: int
    mpt_action: str
    mpt_col: str
    warehouse: bool
    list_hit: ListHit | None
    perm_hit: PermHit | None
    alert: str = ""  # ipc|1|2|3|4

    def row_key(self) -> str:
        return f"{self.client}|{self.invoice}|{self.pn}"


ALERT_LABELS = {
    "1": "нет заявки",
    "2": "нет вырезки >7д",
    "3": "нет МПТ >14д",
    "4": "ТАЗ не обновлён",
    "ipc": "IPC",
}


def pick_alert(
    *,
    ipc: bool,
    has_perm: bool,
    list_hit: ListHit | None,
    age: int,
) -> str | None:
    """Exactly one alert per order, pipeline priority."""
    if ipc:
        return "ipc"
    if has_perm:
        return "4"
    if list_hit is None:
        return "1"
    if list_hit.stage == STAGE_REQUESTED and age > 7:
        return "2"
    if age > 14:
        return "3"
    return None


def load_lists(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="Списки на МПТ")
    df = df.rename(
        columns={
            "Part number": "_pn_raw",
            "Alternative part number": "_alt_raw",
            "Клиент": "_client_raw",
            "Статус по вырезке": "_status_raw",
            "Дата": "_batch",
            "Description": "_desc",
        }
    )
    df["_pn"] = df["_pn_raw"].map(norm_pn)
    df["_alt"] = df.get("_alt_raw", pd.Series(dtype=object)).map(norm_pn)
    df["_client"] = df["_client_raw"].map(norm_client)
    df["_status_raw"] = df["_status_raw"].map(norm_text)
    df["_stage"] = df["_status_raw"].map(normalize_extract_stage)
    df["_batch"] = df.get("_batch", pd.Series(dtype=object)).map(norm_text)
    df = df[df["_pn"] != ""].copy()
    return df


def load_permissions(path: Path) -> dict[str, pd.DataFrame]:
    out: dict[str, pd.DataFrame] = {}
    for sheet in ("ФЭИ 2025-2026", "ЛР 2025-2026", "МЕТЕОР 2025-2026"):
        df = pd.read_excel(path, sheet_name=sheet)
        df["_pn"] = df["p/n"].map(norm_pn)
        df["_desc"] = df.get("Description", pd.Series(dtype=object)).map(norm_text)
        df["_ac"] = df.get("Тип Воздушного судна", pd.Series(dtype=object)).map(norm_text)
        doc_col = "№ и дата документа Минпромторга"
        df["_doc"] = df.get(doc_col, pd.Series(dtype=object)).map(norm_text)
        df = df[df["_pn"] != ""].copy()
        out[sheet] = df
    return out


def find_list_hit(lists: pd.DataFrame, client: str, pn: str, alt: str) -> ListHit | None:
    if not pn and not alt:
        return None
    # prefer client+pn, then pn-only, then alt variants
    candidates: list[tuple[int, pd.Series, str]] = []
    for via, code in (("pn", pn), ("alt", alt)):
        if not code:
            continue
        hits = lists[lists["_pn"] == code]
        if hits.empty:
            continue
        same = hits[hits["_client"].map(lambda c: c.casefold() == client.casefold())]
        pool = same if not same.empty else hits
        for _, row in pool.iterrows():
            rank = 0 if not same.empty and via == "pn" else (1 if via == "pn" else 2)
            candidates.append((rank, row, via))
    if not candidates:
        return None
    candidates.sort(key=lambda x: x[0])
    _, row, via = candidates[0]
    return ListHit(
        client=norm_text(row["_client"]),
        pn=norm_text(row["_pn"]),
        status_raw=norm_text(row["_status_raw"]),
        stage=row["_stage"],
        batch=norm_text(row["_batch"]),
        matched_via=via,
    )


def find_perm_hit(
    perms: dict[str, pd.DataFrame],
    invoicer: str,
    pn: str,
    alt: str,
    taz_ac: str,
) -> PermHit | None:
    sheet = INVOICER_TO_SHEET.get(invoicer.upper() if invoicer.isascii() else invoicer, None)
    # Cyrillic keys
    sheet = INVOICER_TO_SHEET.get(invoicer, sheet)
    search_sheets: list[str]
    if sheet:
        search_sheets = [sheet]
    else:
        search_sheets = list(perms.keys())

    def lookup(code: str, via: str) -> PermHit | None:
        if not code:
            return None
        for sh in search_sheets:
            df = perms[sh]
            hits = df[df["_pn"] == code]
            if hits.empty:
                continue
            # prefer AC match
            best = None
            best_rank = 99
            for _, row in hits.iterrows():
                ac_m = ac_types_compatible(taz_ac, row["_ac"])
                rank = {"match": 0, "unknown": 1, "mismatch": 2}[ac_m]
                if rank < best_rank:
                    best_rank = rank
                    best = PermHit(
                        sheet=sh,
                        pn=row["_pn"],
                        description=row["_desc"],
                        ac_type=row["_ac"],
                        doc=row["_doc"],
                        matched_via=via,
                        ac_match=ac_m,
                    )
            if best is not None:
                return best
        return None

    hit = lookup(pn, "pn")
    if hit:
        return hit
    return lookup(alt, "alt")


def build_rows(
    taz: pd.DataFrame,
    lists: pd.DataFrame,
    perms: dict[str, pd.DataFrame],
    report_date: date,
) -> list[MptRow]:
    df = taz.copy()
    df = df[df[COL_CATEGORY].astype(str).str.strip().str.upper() == "ROTABLE"]
    df = df[df[COL_STATUS].isin(OPEN_STATUSES)]
    df = df[~df[COL_CUSTOMER].astype(str).str.strip().isin(EXCLUDED_CUSTOMERS)]

    rows: list[MptRow] = []
    for _, r in df.iterrows():
        order_date = parse_date(r.get(COL_ORDER_DATE))
        if order_date is None:
            continue
        age = (report_date - order_date).days
        if age <= 7:
            continue

        client = norm_client(r.get(COL_CUSTOMER))
        if client == "Utair" and utair_no_mpt_needed(r.get(COL_COMMENT)):
            continue

        pn = norm_pn(r.get(COL_PN))
        alt = norm_pn(r.get(COL_ALT_PN))
        action = norm_text(r.get(COL_MPT_ACTION))
        invoicer = norm_text(r.get(COL_INVOICER))
        ac_type = norm_text(r.get(COL_AC_TYPE))
        mpt_col = norm_text(r.get(COL_MPT))

        list_hit = find_list_hit(lists, client, pn, alt)
        perm_hit = find_perm_hit(perms, invoicer, pn, alt, ac_type)

        has_perm = perm_hit is not None
        taz_has = action.casefold() == "есть мпт"
        if has_perm and taz_has:
            continue  # OK — skip

        ipc = action.casefold() in {"нет ipc от клиента", "ipc не будет получен"}
        alert = pick_alert(ipc=ipc, has_perm=has_perm, list_hit=list_hit, age=age)
        if alert is None:
            continue

        rows.append(
            MptRow(
                invoice=normalize_invoice(r.get(COL_INVOICE)),
                client=client,
                pn=pn or norm_text(r.get(COL_PN)),
                alt_pn=alt,
                description=norm_text(r.get(COL_DESC)),
                status=norm_text(r.get(COL_STATUS)),
                invoicer=invoicer,
                ac_type=ac_type,
                order_date=order_date,
                age_days=age,
                mpt_action=action or "—",
                mpt_col=mpt_col or "—",
                warehouse=is_warehouse(client),
                list_hit=list_hit,
                perm_hit=perm_hit,
                alert=alert,
            )
        )

    from collections import Counter

    # global order used only as fallback; HTML re-groups by client anyway
    client_n = Counter(r.client for r in rows)
    rows.sort(
        key=lambda x: (
            -client_n[x.client],
            x.client.casefold(),
            x.order_date or date.max,
            x.invoice,
            x.pn,
        )
    )
    return rows


def render_table(rows: list[MptRow]) -> str:
    if not rows:
        return '<p class="muted">Нет позиций.</p>'
    head = """
<table>
<thead><tr>
<th>Статус</th><th>№ счета</th><th>P/N</th><th>Description</th>
<th>Дней с даты заказа</th><th>Действия по МПТ</th><th>Списки по МПТ</th>
<th>Полученные МПТ</th><th>ТИП ВС в ТАЗ</th>
</tr></thead><tbody>
"""
    body: list[str] = []
    for row in rows:
        tr_cls = ' class="wh"' if row.warehouse else ""

        pn_cell = esc(row.pn)
        if row.list_hit and row.list_hit.matched_via == "alt":
            pn_cell += ' <span class="tag alt">матч ALT в списках</span>'
        if row.perm_hit and row.perm_hit.matched_via == "alt":
            pn_cell += ' <span class="tag alt">матч ALT в перечне</span>'
        if row.alt_pn:
            pn_cell += f'<div class="muted">ALT: {esc(row.alt_pn)}</div>'

        if row.list_hit:
            list_cell = (
                f'<div><b>{esc(stage_label(row.list_hit.stage))}</b></div>'
                f'<div class="muted">{esc(row.list_hit.status_raw)}'
                f'{(" · " + esc(row.list_hit.batch)) if row.list_hit.batch else ""}</div>'
            )
        else:
            list_cell = '<span class="bad">нет заявки</span>'

        if row.perm_hit:
            ac_tag = {
                "match": '<span class="tag ok">тип ВС ≈</span>',
                "mismatch": '<span class="tag bad">тип ВС ≠</span>',
                "unknown": '<span class="tag">тип ВС ?</span>',
            }[row.perm_hit.ac_match]
            perm_cell = (
                f'<div><b>есть</b> · {esc(row.perm_hit.sheet)} {ac_tag}</div>'
                f'<div class="muted">{esc(row.perm_hit.doc)}</div>'
                f'<div class="muted">тип ВС в перечне: {esc(row.perm_hit.ac_type)}</div>'
            )
        else:
            sheet = INVOICER_TO_SHEET.get(row.invoicer, "—")
            perm_cell = f'<span class="bad">нет в {esc(sheet)}</span>'

        alert_cell = (
            f'<span class="tag a{esc(row.alert)}">{esc(ALERT_LABELS.get(row.alert, row.alert))}</span>'
        )

        body.append(
            f"<tr{tr_cls}>"
            f"<td>{alert_cell}</td>"
            f"<td>{esc(row.invoice)}</td>"
            f"<td>{pn_cell}</td>"
            f"<td>{esc(row.description)}</td>"
            f'<td class="num">{row.age_days}</td>'
            f"<td>{esc(row.mpt_action)}</td>"
            f"<td>{list_cell}</td>"
            f"<td>{perm_cell}</td>"
            f"<td>{esc(row.ac_type) or '—'}</td>"
            f"</tr>"
        )
    return head + "\n".join(body) + "\n</tbody></table>"


def render_html(rows: list[MptRow], report_date: date) -> str:
    from collections import Counter

    counts = Counter(r.alert for r in rows)
    wh_n = sum(1 for r in rows if r.warehouse)
    by_client: dict[str, list[MptRow]] = {}
    for row in rows:
        by_client.setdefault(row.client, []).append(row)
    ordered_clients = sorted(
        by_client.items(), key=lambda kv: (-len(kv[1]), kv[0].casefold())
    )

    client_blocks: list[str] = []
    tabs: list[str] = []
    for idx, (client, items) in enumerate(ordered_clients):
        # earliest → newest
        items_sorted = sorted(
            items,
            key=lambda r: (r.order_date or date.max, r.invoice, r.pn),
        )
        anchor = f"c{idx}"
        wh = ' <span class="tag wh">на склад</span>' if any(i.warehouse for i in items_sorted) else ""
        alert_pills = "".join(
            f'<span class="tag a{a}">{esc(ALERT_LABELS[a])} · {n}</span>'
            for a, n in sorted(Counter(r.alert for r in items_sorted).items(), key=lambda x: x[0])
        )
        tabs.append(f'<a href="#{anchor}">{esc(client)} ({len(items_sorted)})</a>')
        client_blocks.append(
            f"""
<details class="group" id="{anchor}">
  <summary>
    <span class="group-name">{esc(client)}{wh}</span>
    <span class="pill">{len(items_sorted)} поз.</span>
    {alert_pills}
  </summary>
  <div class="group-body">
    {render_table(items_sorted)}
  </div>
</details>
"""
        )

    return f"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>Задержки МПТ — {report_date.strftime('%d.%m.%Y')}</title>
<style>
:root {{
  --bg:#e7ecef; --card:#ffffff; --ink:#202020; --muted:#4c4c4c;
  --navy:#022f40; --cyan:#d5fbff; --orange:#fe621d; --line:#c4c4c4;
  --red:#c0392b; --green:#1e7a46;
}}
* {{ box-sizing:border-box; }}
body {{
  margin:0;
  font-family: Arial, Calibri, Helvetica, sans-serif;
  color:var(--ink);
  background: linear-gradient(180deg, #022f40 0 160px, var(--bg) 160px);
  line-height:1.45;
}}
.wrap {{ max-width:1400px; margin:0 auto; padding:28px 16px 64px; }}
.brand {{
  display:flex; align-items:center; justify-content:space-between; gap:12px;
  color:#ffffff; margin-bottom:18px;
}}
.brand-mark {{
  font-size:13px; letter-spacing:.12em; text-transform:uppercase;
  background:var(--cyan); color:var(--navy); padding:6px 10px; font-weight:700;
}}
h1 {{ font-size:28px; margin:0 0 6px; color:#ffffff; font-weight:700; }}
.sub {{ color:#d5fbff; margin-bottom:18px; font-size:14px; }}
h2 {{ font-size:20px; margin:28px 0 12px; color:var(--navy); }}
.card {{
  background:var(--card); border:1px solid var(--line); border-radius:8px;
  padding:18px 18px 8px;
}}
.kpi-grid {{
  display:grid; grid-template-columns:repeat(6,minmax(0,1fr)); gap:12px; margin:12px 0 18px;
}}
.kpi {{
  background:#f7fbfc; border:1px solid #d7e2e6; border-radius:8px; padding:12px 14px;
}}
.label {{ font-size:11px; text-transform:uppercase; letter-spacing:.04em; color:#757575; }}
.value {{ font-size:22px; font-weight:700; margin-top:4px; font-variant-numeric:tabular-nums; color:var(--navy); }}
.muted {{ color:#757575; font-size:12px; margin-top:4px; }}
.tabs {{ display:flex; flex-wrap:wrap; gap:8px; margin:8px 0 14px; }}
.tabs a {{
  text-decoration:none; color:var(--navy); background:#ffffff; border:1px solid var(--line);
  padding:8px 12px; border-radius:6px; font-size:13px; font-weight:700;
}}
.tabs a:hover {{ background:var(--cyan); }}
details.group {{
  background:var(--card); border:1px solid var(--line); border-radius:8px;
  margin:10px 0; overflow:hidden;
}}
details.group summary {{
  cursor:pointer; list-style:none; padding:14px 16px; display:flex; flex-wrap:wrap; gap:8px; align-items:center;
  background:linear-gradient(90deg, #022f40 0%, #022f40 12px, #ffffff 12px);
}}
details.group summary::-webkit-details-marker {{ display:none; }}
.group-name {{ font-weight:700; font-size:16px; margin-right:auto; color:var(--navy); padding-left:8px; }}
.pill {{
  font-size:12px; background:var(--cyan); border:1px solid #b6e8ef; color:var(--navy);
  border-radius:4px; padding:4px 10px; font-weight:700;
}}
.group-body {{ padding:0 14px 16px; border-top:1px solid var(--line); overflow-x:auto; }}
table {{ width:100%; border-collapse:collapse; font-size:12.5px; border:1px solid #c4c4c4; }}
th, td {{
  border-bottom:1px solid #d9d9d9; border-right:1px solid #d0d0d0;
  padding:8px 8px; vertical-align:top; text-align:left;
}}
th:last-child, td:last-child {{ border-right:none; }}
th {{
  font-size:10px; text-transform:uppercase; letter-spacing:.03em; color:#4c4c4c;
  background:#f3f7f8; border-bottom:1px solid #b8b8b8;
}}
td.num {{ text-align:right; font-variant-numeric:tabular-nums; white-space:nowrap; }}
tr.wh {{ background:#f0fafc; }}
.tag {{
  display:inline-block; font-size:11px; font-weight:700; border-radius:4px;
  padding:2px 7px; margin:1px 2px; border:1px solid #d0d0d0; background:#f3f3f3; color:var(--navy);
}}
.tag.alt {{ background:#fff3cd; border-color:#e6d89c; }}
.tag.wh {{ background:#d5fbff; border-color:#b6e8ef; }}
.tag.ok {{ background:#e3f6ea; border-color:#b7e0c4; color:var(--green); }}
.tag.bad, .bad {{ background:#fde8e6; border-color:#f0b8b2; color:var(--red); }}
.tag.a1 {{ background:#fde8e6; border-color:#f0b8b2; color:var(--red); }}
.tag.a2 {{ background:#fff3cd; border-color:#e6d89c; }}
.tag.a3 {{ background:#ffe0cc; border-color:#f0b48a; color:var(--orange); }}
.tag.a4 {{ background:#e8f0fe; border-color:#b6c9f0; }}
.tag.aipc {{ background:#efe8f8; border-color:#c9b6e0; }}
.footer {{ margin-top:28px; color:#4c4c4c; font-size:12px; }}
@media (max-width:900px) {{
  .kpi-grid {{ grid-template-columns:repeat(2,minmax(0,1fr)); }}
}}
</style>
</head>
<body>
<div class="wrap">
  <div class="brand">
    <div class="brand-mark">FASTAIR</div>
    <div style="font-size:13px;opacity:.9">MPT delay report · ROTABLE</div>
  </div>
  <h1>Задержки получения МПТ</h1>
  <div class="sub">
    Срез {report_date.strftime('%d.%m.%Y')} · Category ROTABLE · Status 1/2/6 · заказ &gt;7 дн. · без Белавиа ·
    одна тревога на заказ · разделы по клиентам · внутри — от ранних к свежим ·
    Utair DAT MOW / DAT TEHRAN исключены (МПТ не нужен)
  </div>

  <section class="card">
    <h2 style="margin-top:0">Сводка</h2>
    <div class="kpi-grid">
      <div class="kpi"><div class="label">Всего</div><div class="value">{len(rows)}</div>
        <div class="muted">{len(by_client)} клиентов · склад {wh_n}</div></div>
      <div class="kpi"><div class="label">1 · Нет заявки</div><div class="value">{counts.get('1', 0)}</div></div>
      <div class="kpi"><div class="label">2 · Нет вырезки &gt;7д</div><div class="value">{counts.get('2', 0)}</div></div>
      <div class="kpi"><div class="label">3 · Нет МПТ &gt;14д</div><div class="value">{counts.get('3', 0)}</div></div>
      <div class="kpi"><div class="label">4 · ТАЗ не обновлён</div><div class="value">{counts.get('4', 0)}</div></div>
      <div class="kpi"><div class="label">IPC</div><div class="value">{counts.get('ipc', 0)}</div></div>
    </div>
    <p class="muted">
      Приоритет одной тревоги: IPC → ТАЗ не обновлён (есть в Перечне) → нет заявки → нет вырезки → нет МПТ &gt;14д.
      Исключены позиции, где МПТ есть и в Перечне, и в ТАЗ («Есть МПТ»).
      Тип ВС сверяется мягко («тип ВС ≠» при расхождении). Голубая строка — «на склад».
    </p>
  </section>

  <div class="tabs">
    {''.join(tabs)}
  </div>

  <h2>По клиентам</h2>
  {''.join(client_blocks)}

  <div class="footer">
    Источники: ТАЗ · Брокерские заявки (Списки на МПТ) · Перечни зап. частей 0% НДС (ФЭИ/ЛР/МЕТЕОР 2025–2026).
    Self-contained HTML — открывается локально без интернета.
  </div>
</div>
</body>
</html>
"""


def generate(
    taz_path: Path,
    broker_path: Path,
    permissions_path: Path,
    output_html: Path,
    report_date: date | None = None,
) -> tuple[Path, Path, list[MptRow]]:
    if report_date is None:
        m = re.search(r"(\d{2})\.(\d{2})\.(\d{4})", taz_path.name)
        if m:
            d, mo, y = map(int, m.groups())
            report_date = date(y, mo, d)
        else:
            report_date = date.today()

    taz = load_taz(taz_path, excluded_invoicers=set())
    lists = load_lists(broker_path)
    perms = load_permissions(permissions_path)
    rows = build_rows(taz, lists, perms, report_date)

    output_html.parent.mkdir(parents=True, exist_ok=True)
    html_text = render_html(rows, report_date)
    output_html.write_text(html_text, encoding="utf-8")

    zip_path = output_html.with_suffix(".zip")
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(output_html.name, html_text.encode("utf-8"))

    return output_html, zip_path, rows


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--taz", type=Path, required=True)
    parser.add_argument("--broker", type=Path, required=True)
    parser.add_argument("--permissions", type=Path, required=True)
    parser.add_argument("--output", "-o", type=Path, default=Path("output/mpt_delay.html"))
    parser.add_argument("--date", type=str, default=None, help="dd.mm.yyyy")
    args = parser.parse_args()

    report_date = None
    if args.date:
        d, m, y = map(int, args.date.split("."))
        report_date = date(y, m, d)

    html_path, zip_path, rows = generate(
        args.taz, args.broker, args.permissions, args.output, report_date
    )
    from collections import Counter

    c = Counter(r.alert for r in rows)
    print(f"rows in report: {len(rows)}")
    print(f"clients: {len({r.client for r in rows})}")
    print(f"alerts: {dict(c)}")
    print(f"wrote {html_path}")
    print(f"wrote {zip_path}")


if __name__ == "__main__":
    main()
