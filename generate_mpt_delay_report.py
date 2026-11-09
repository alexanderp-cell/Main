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
Исключение: действующее МПТ в Перечне (≤1 год от даты получения) И в ТАЗ «Есть МПТ» → не кладём.
Действующее МПТ в столбце «Списки по МПТ» показывается как «ЕСТЬ МПТ» (заявка не важна).
«Закупка на склад» — в общем отчёте с пометкой.
"""

from __future__ import annotations

import argparse
import base64
import html
import re
import zipfile
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from functools import lru_cache
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

ASSETS_DIR = Path(__file__).resolve().parent / "assets"


@lru_cache(maxsize=32)
def data_uri(name: str) -> str:
    """Embed local PNG as data URI so HTML stays self-contained."""
    path = ASSETS_DIR / name
    raw = path.read_bytes()
    mime = "image/png" if path.suffix.lower() == ".png" else "image/jpeg"
    return f"data:{mime};base64,{base64.b64encode(raw).decode('ascii')}"

COL_CATEGORY = "Category"
COL_MPT = "МПТ"
COL_MPT_ACTION = "Действия по МПТ"
COL_AC_TYPE = "ТИП ВС (Продажи)"
COL_ALT_PN = "ALT P/N"

OPEN_STATUSES = {"1 NOT PAID", "2 PAID", "6 TROUBLE"}
CLOSED_STATUSES = {"5 CANCELLED", "7 REFUND"}
EXCLUDED_CUSTOMERS = {"Белавиа"}
WAREHOUSE_MARKERS = {"закупка на склад"}
EMPTY_INVOICES = {"", "-", "—", "–", "nan", "none"}

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


DAP_MOW_CLIENTS = {"ак сибирь", "сибирь", "s7 engineering"}


def comment_has_dap_mow(comment: Any) -> bool:
    """DAP MOW / DAP Москва — МПТ не требуется."""
    s = re.sub(r"\s+", " ", norm_text(comment).upper().replace("Ё", "Е"))
    return "DAP MOW" in s or "DAP МОСКВА" in s


def no_mpt_needed(client: str, comment: Any) -> bool:
    if client == "Utair" and utair_no_mpt_needed(comment):
        return True
    if client.casefold() in DAP_MOW_CLIENTS and comment_has_dap_mow(comment):
        return True
    return False


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

    # Embraer 170/175: E170 / E175 ↔ Embraer 170/175
    # (не кладём голый EMBRAER — иначе склеится с семейством 145/Legacy)
    if (
        any(x in extra for x in ("E170", "E175", "ERJ170", "ERJ175"))
        or "E170" in joined
        or "E175" in joined
        or ("EMBRAER" in joined and ("170" in joined or "175" in joined))
    ):
        extra.update({"E170", "E175", "ERJ170", "ERJ175", "EMB170", "EMB175"})

    # Embraer 145/135 ↔ Legacy (Legacy 600/650)
    if (
        any(x in extra for x in ("E145", "E135", "ERJ145", "ERJ135", "LEGACY", "LEGACY600", "LEGACY650"))
        or "E145" in joined
        or "E135" in joined
        or "ERJ145" in joined
        or "ERJ135" in joined
        or "LEGACY" in joined
        or ("EMBRAER" in joined and ("145" in joined or "135" in joined))
    ):
        extra.update(
            {
                "E145",
                "E135",
                "ERJ145",
                "ERJ135",
                "EMB145",
                "EMB135",
                "LEGACY",
                "LEGACY600",
                "LEGACY650",
            }
        )

    # Bombardier CL-600-2B19 ↔ CRJ200 / CH850 / Challenger 850
    if (
        "CL6002B19" in joined
        or ("CL600" in joined and "2B19" in joined)
        or "2B19" in joined
        or "CRJ200" in joined
        or "CRJ200" in extra
        or ("CRJ" in joined and "200" in joined)
        or "CH850" in joined
        or "CH850" in extra
        or "CHALLENGER850" in joined
        or ("CHALLENGER" in joined and "850" in joined)
    ):
        extra.update(
            {
                "CL600",
                "CL6002B19",
                "2B19",
                "CRJ200",
                "CH850",
                "CHALLENGER850",
            }
        )

    # Bombardier BD-700-1A10/1A11 ↔ Global Express / Global 5000–6500 family
    global_tokens = {
        "BD700",
        "BD7001A10",
        "BD7001A11",
        "1A10",
        "1A11",
        "GLOBAL",
        "GLOBALEXPRESS",
        "GLOBALEXPRESSXRS",
        "XRS",
        "GLOBAL6000",
        "GLOBAL6500",
        "GLOBAL5000",
        "GLOBAL5000GVFD",
        "GVFD",
        "GLOBAL5500",
        "G6000",
        "G6500",
        "G5000",
        "G5500",
    }
    if (
        "BD700" in joined
        or "BD-700" in norm_text(value).upper()
        or "1A10" in joined
        or "1A11" in joined
        or any(t in joined for t in ("GLOBALEXPRESS", "GLOBAL6000", "GLOBAL6500", "GLOBAL5000", "GLOBAL5500"))
        or any(t in extra for t in ("G6000", "G6500", "G5000", "G5500", "GLOBAL", "XRS", "GVFD"))
        or ("GLOBAL" in joined and any(n in joined for n in ("5000", "5500", "6000", "6500", "EXPRESS")))
    ):
        extra.update(global_tokens)
        # also keep spaced-style tokens from parts
        extra.update({"BD700", "GLOBAL", "EXPRESS", "6000", "6500", "5000", "5500"})

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
    doc_date: date | None = None
    valid_until: date | None = None
    active: bool = True


DOC_DATE_RE = re.compile(r"от\s*(\d{1,2})[./](\d{1,2})[./](\d{2,4})", re.I)


def parse_mpt_doc_date(doc: str) -> date | None:
    """Extract receipt date from «№ и дата документа Минпромторга»."""
    s = norm_text(doc).replace("\xa0", " ")
    m = DOC_DATE_RE.search(s)
    if not m:
        return None
    day, month, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if year < 100:
        year += 2000
    try:
        return date(year, month, day)
    except ValueError:
        return None


def add_one_year(d: date) -> date:
    try:
        return d.replace(year=d.year + 1)
    except ValueError:
        return d.replace(year=d.year + 1, day=28)


def enrich_perm_validity(hit: PermHit, report_date: date) -> PermHit:
    """МПТ действует 1 год от даты получения (даты документа)."""
    doc_date = parse_mpt_doc_date(hit.doc)
    if doc_date is None:
        hit.doc_date = None
        hit.valid_until = None
        hit.active = True  # в актуальном перечне без даты — считаем действующим
        return hit
    valid_until = add_one_year(doc_date)
    hit.doc_date = doc_date
    hit.valid_until = valid_until
    hit.active = report_date <= valid_until
    return hit


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
    report_date: date,
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
            # prefer AC match, then still-active MPT
            best = None
            best_rank = 99
            for _, row in hits.iterrows():
                ac_m = ac_types_compatible(taz_ac, row["_ac"])
                cand = enrich_perm_validity(
                    PermHit(
                        sheet=sh,
                        pn=row["_pn"],
                        description=row["_desc"],
                        ac_type=row["_ac"],
                        doc=row["_doc"],
                        matched_via=via,
                        ac_match=ac_m,
                    ),
                    report_date,
                )
                rank = {"match": 0, "unknown": 1, "mismatch": 2}[ac_m]
                # prefer active over expired at same AC rank
                rank = rank * 2 + (0 if cand.active else 1)
                if rank < best_rank:
                    best_rank = rank
                    best = cand
            if best is not None:
                return best
        return None

    hit = lookup(pn, "pn")
    if hit:
        return hit
    return lookup(alt, "alt")


def _closed_client_pns(taz: pd.DataFrame) -> set[tuple[str, str]]:
    """(client, pn) pairs that have CANCELLED / REFUND rows in TAZ."""
    out: set[tuple[str, str]] = set()
    for _, r in taz.iterrows():
        st = norm_text(r.get(COL_STATUS))
        if st not in CLOSED_STATUSES and "CANCEL" not in st.upper() and "REFUND" not in st.upper():
            continue
        client = norm_client(r.get(COL_CUSTOMER))
        pn = norm_pn(r.get(COL_PN))
        if client and pn:
            out.add((client, pn))
    return out


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
    closed_pairs = _closed_client_pns(taz)

    rows: list[MptRow] = []
    for _, r in df.iterrows():
        order_date = parse_date(r.get(COL_ORDER_DATE))
        if order_date is None:
            continue
        age = (report_date - order_date).days
        if age <= 7:
            continue

        client = norm_client(r.get(COL_CUSTOMER))
        if no_mpt_needed(client, r.get(COL_COMMENT)):
            continue

        pn = norm_pn(r.get(COL_PN))
        invoice = normalize_invoice(r.get(COL_INVOICE))
        # Хвост после отмены/refund: открытый статус, но № счета пустой/«-»
        if invoice.casefold() in EMPTY_INVOICES and (client, pn) in closed_pairs:
            continue

        alt = norm_pn(r.get(COL_ALT_PN))
        action = norm_text(r.get(COL_MPT_ACTION))
        invoicer = norm_text(r.get(COL_INVOICER))
        ac_type = norm_text(r.get(COL_AC_TYPE))
        mpt_col = norm_text(r.get(COL_MPT))

        list_hit = find_list_hit(lists, client, pn, alt)
        perm_hit = find_perm_hit(perms, invoicer, pn, alt, ac_type, report_date)

        has_active_perm = perm_hit is not None and perm_hit.active
        taz_has = action.casefold() == "есть мпт"
        if has_active_perm and taz_has:
            continue  # OK — skip

        ipc = action.casefold() in {"нет ipc от клиента", "ipc не будет получен"}
        alert = pick_alert(ipc=ipc, has_perm=has_active_perm, list_hit=list_hit, age=age)
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
<colgroup>
  <col/><col/><col class="pn"/><col class="desc"/><col/><col/><col/><col/><col/>
</colgroup>
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

        # Если есть действующее МПТ (≤1 год от даты получения) — заявка в списках не важна
        if row.perm_hit and row.perm_hit.active:
            until = (
                f" · до {row.perm_hit.valid_until.strftime('%d.%m.%Y')}"
                if row.perm_hit.valid_until
                else ""
            )
            list_cell = f"<div><b>ЕСТЬ МПТ</b>{esc(until)}</div>"
        elif row.list_hit:
            status_up = (row.list_hit.status_raw or stage_label(row.list_hit.stage)).upper()
            batch = f" · {esc(row.list_hit.batch)}" if row.list_hit.batch else ""
            list_cell = f'<div><b>{esc(status_up)}</b>{batch}</div>'
        else:
            list_cell = '<span class="bad">НЕТ ЗАЯВКИ</span>'

        if row.perm_hit:
            ac_type_perm = row.perm_hit.ac_type or "—"
            ac_tag = {
                "match": '<span class="tag ok">совпадает</span>',
                "mismatch": '<span class="tag bad">не совпадает</span>',
                "unknown": '<span class="tag">примерно</span>',
            }[row.perm_hit.ac_match]
            active_note = ""
            if not row.perm_hit.active and row.perm_hit.valid_until:
                active_note = (
                    f' <span class="tag bad">истекло '
                    f'{row.perm_hit.valid_until.strftime("%d.%m.%Y")}</span>'
                )
            elif row.perm_hit.valid_until:
                active_note = (
                    f' <span class="tag ok">до '
                    f'{row.perm_hit.valid_until.strftime("%d.%m.%Y")}</span>'
                )
            perm_cell = (
                f'<div><b>ЕСТЬ МПТ на</b> «{esc(ac_type_perm)}» {ac_tag}{active_note}</div>'
                f'<div class="muted">{esc(row.perm_hit.sheet)}'
                f'{(" · " + esc(row.perm_hit.doc)) if row.perm_hit.doc else ""}</div>'
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

    img_banner = data_uri("tamagotchi-banner-wide.png")
    img_device = data_uri("tamagotchi-device-small.png")
    img_pattern = data_uri("tamagotchi-pattern.png")
    img_hungry = data_uri("tamagotchi-pet-hungry.png")
    img_sick = data_uri("tamagotchi-pet-sick.png")
    img_happy = data_uri("tamagotchi-pet-happy.png")
    img_ok = data_uri("tamagotchi-pet-ok.png")

    return f"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>Задержки МПТ — {report_date.strftime('%d.%m.%Y')}</title>
<style>
:root {{
  --shell:#ff9ebb;
  --shell-deep:#e8628c;
  --shell-light:#ffe3ee;
  --lcd:#aac678;
  --lcd-deep:#284020;
  --lcd-lit:#c8dca0;
  --btn:#fff078;
  --teal:#5cc8c4;
  --ink:#284020;
  --muted:#5a6b4a;
  --card:#fff7fb;
  --line:#e8628c;
  --red:#c0392b;
  --orange:#d35400;
  --green:#1e7a46;
  --pixel: 4px;
}}
* {{ box-sizing:border-box; }}
body {{
  margin:0;
  font-family: "Courier New", Courier, ui-monospace, monospace;
  color:var(--ink);
  background-color: var(--shell-light);
  background-image:
    linear-gradient(180deg, rgba(255,158,187,.55) 0%, rgba(255,247,251,.92) 220px),
    url('{img_pattern}');
  background-size: auto, 128px 128px;
  background-attachment: fixed;
  line-height:1.45;
  image-rendering: pixelated;
}}
.wrap {{ max-width:1400px; margin:0 auto; padding:24px 16px 64px; }}
.hero {{
  display:grid; grid-template-columns: 1.4fr .7fr; gap:16px; align-items:stretch;
  margin-bottom:18px;
}}
.hero-banner {{
  border:4px solid var(--shell-deep);
  box-shadow: 6px 6px 0 #28402033;
  background:#fff;
  overflow:hidden;
  position:relative;
}}
.hero-banner img {{
  display:block; width:100%; height:180px; object-fit:cover; object-position:center;
  image-rendering: pixelated;
}}
.hero-device {{
  border:4px solid var(--shell-deep);
  background: linear-gradient(160deg, var(--shell) 0%, #ffd6e4 55%, var(--teal) 160%);
  box-shadow: 6px 6px 0 #28402033;
  display:flex; flex-direction:column; align-items:center; justify-content:center;
  padding:12px; text-align:center; gap:8px;
}}
.hero-device img {{ width:110px; height:auto; image-rendering:pixelated; }}
.brand-mark {{
  font-size:12px; letter-spacing:.18em; text-transform:uppercase;
  background:var(--btn); color:var(--lcd-deep); padding:6px 10px; font-weight:800;
  border:3px solid var(--lcd-deep);
}}
.lcd-title {{
  background:var(--lcd); color:var(--lcd-deep); border:4px solid var(--lcd-deep);
  padding:14px 16px; margin:0 0 10px;
  box-shadow: inset 0 0 0 2px var(--lcd-lit);
}}
.lcd-title h1 {{
  font-size:26px; margin:0 0 6px; font-weight:800; letter-spacing:.04em;
  text-transform:uppercase;
}}
.sub {{ color:var(--muted); font-size:13px; }}
.blink {{ animation: blink 1.4s steps(2, jump-none) infinite; }}
@keyframes blink {{ 50% {{ opacity:.35; }} }}
h2 {{
  font-size:18px; margin:24px 0 12px; color:var(--lcd-deep);
  letter-spacing:.08em; text-transform:uppercase;
}}
.card {{
  background:var(--card); border:4px solid var(--shell-deep);
  padding:16px 16px 8px; box-shadow: 5px 5px 0 #28402022;
}}
.kpi-grid {{
  display:grid; grid-template-columns:repeat(6,minmax(0,1fr)); gap:10px; margin:12px 0 14px;
}}
.kpi {{
  background:var(--lcd); border:3px solid var(--lcd-deep); padding:10px 10px 12px;
  box-shadow: inset 0 0 0 2px var(--lcd-lit);
  display:flex; flex-direction:column; gap:4px; min-height:110px;
}}
.kpi img {{ width:42px; height:42px; image-rendering:pixelated; align-self:flex-start; }}
.label {{ font-size:10px; text-transform:uppercase; letter-spacing:.05em; color:var(--lcd-deep); font-weight:800; }}
.value {{ font-size:24px; font-weight:800; font-variant-numeric:tabular-nums; color:var(--lcd-deep); }}
.muted {{ color:var(--muted); font-size:12px; margin-top:4px; }}
.tabs {{ display:flex; flex-wrap:wrap; gap:8px; margin:8px 0 14px; }}
.tabs a {{
  text-decoration:none; color:var(--lcd-deep); background:var(--btn);
  border:3px solid var(--lcd-deep); padding:7px 10px; font-size:12px; font-weight:800;
}}
.tabs a:hover {{ background:var(--teal); }}
details.group {{
  background:var(--card); border:4px solid var(--shell-deep);
  margin:12px 0; overflow:visible; box-shadow: 4px 4px 0 #28402022;
}}
details.group summary {{
  cursor:pointer; list-style:none; padding:12px 14px; display:flex; flex-wrap:wrap; gap:8px; align-items:center;
  background: linear-gradient(90deg, var(--shell) 0 14px, var(--lcd) 14px 18px, #fff7fb 18px);
}}
details.group summary::-webkit-details-marker {{ display:none; }}
.group-name {{ font-weight:800; font-size:15px; margin-right:auto; color:var(--lcd-deep); padding-left:6px; letter-spacing:.03em; }}
.pill {{
  font-size:11px; background:var(--btn); border:2px solid var(--lcd-deep); color:var(--lcd-deep);
  padding:3px 8px; font-weight:800;
}}
.group-body {{ padding:0 12px 14px; border-top:3px solid var(--shell-deep); overflow:visible; background:#f7fbe8; }}
table {{ width:100%; border-collapse:separate; border-spacing:0; font-size:12px; border:3px solid var(--lcd-deep); table-layout:fixed; background:var(--lcd); }}
col.pn {{ width:9%; }}
col.desc {{ width:14%; }}
th, td {{
  border-bottom:1px solid #7a9a58; border-right:1px solid #7a9a58;
  padding:8px 7px; vertical-align:top; text-align:left;
  word-wrap:break-word; overflow-wrap:anywhere; color:var(--lcd-deep);
}}
th:last-child, td:last-child {{ border-right:none; }}
th {{
  position:sticky; top:0; z-index:5;
  font-size:12px; font-weight:800; letter-spacing:.03em; text-transform:uppercase;
  color:var(--lcd-deep); background:#9cb86a;
  border-bottom:3px solid var(--lcd-deep);
}}
td.num {{ text-align:right; font-variant-numeric:tabular-nums; white-space:nowrap; }}
tr.wh {{ background:#c5dc98; }}
tbody tr:hover {{ background:#b9d486; }}
.tag {{
  display:inline-block; font-size:10px; font-weight:800; border-radius:0;
  padding:2px 6px; margin:1px 2px; border:2px solid var(--lcd-deep); background:var(--lcd-lit); color:var(--lcd-deep);
}}
.tag.alt {{ background:#fff3cd; border-color:#a08020; }}
.tag.wh {{ background:var(--teal); border-color:var(--lcd-deep); }}
.tag.ok {{ background:#d6f5de; border-color:var(--green); color:var(--green); }}
.tag.bad, .bad {{ background:#ffd0d0; border-color:var(--red); color:var(--red); }}
.tag.a1 {{ background:#ffd0d0; border-color:var(--red); color:var(--red); }}
.tag.a2 {{ background:#fff3cd; border-color:#a08020; }}
.tag.a3 {{ background:#ffd8b8; border-color:var(--orange); color:var(--orange); }}
.tag.a4 {{ background:#d7ecff; border-color:#3a6ea5; }}
.tag.aipc {{ background:#e8d8ff; border-color:#6b4ea0; }}
.footer {{
  margin-top:28px; color:var(--muted); font-size:12px;
  border-top:3px dashed var(--shell-deep); padding-top:12px;
}}
.pixel-row {{ display:flex; gap:10px; flex-wrap:wrap; align-items:center; margin-top:6px; }}
.pixel-row img {{ width:36px; height:36px; image-rendering:pixelated; border:2px solid var(--lcd-deep); background:var(--lcd); }}
@media (max-width:900px) {{
  .hero {{ grid-template-columns:1fr; }}
  .hero-banner img {{ height:140px; }}
  .kpi-grid {{ grid-template-columns:repeat(2,minmax(0,1fr)); }}
}}
</style>
</head>
<body>
<div class="wrap">
  <div class="hero">
    <div class="hero-banner">
      <img src="{img_banner}" alt="Tamagotchi MPT Delay banner"/>
    </div>
    <div class="hero-device">
      <div class="brand-mark">FASTAIR</div>
      <img src="{img_device}" alt="Tamagotchi device"/>
      <div style="font-size:11px;font-weight:800;letter-spacing:.12em;text-transform:uppercase;">CARE · MPT · PET</div>
    </div>
  </div>

  <div class="lcd-title">
    <h1><span class="blink">▶</span> Задержки получения МПТ</h1>
    <div class="sub">
      Срез {report_date.strftime('%d.%m.%Y')} · Category ROTABLE · Status 1/2/6 · заказ &gt;7 дн. · без Белавиа ·
      одна тревога на заказ · разделы по клиентам · внутри — от ранних к свежим ·
      Utair DAT MOW / DAT TEHRAN и АК Сибирь / S7 Engineering DAP MOW исключены (МПТ не нужен)
    </div>
    <div class="pixel-row" aria-hidden="true">
      <img src="{img_hungry}" alt=""/><img src="{img_sick}" alt=""/><img src="{img_ok}" alt=""/><img src="{img_happy}" alt=""/>
    </div>
  </div>

  <section class="card">
    <h2 style="margin-top:0">Сводка · status screen</h2>
    <div class="kpi-grid">
      <div class="kpi"><img src="{img_ok}" alt=""/><div class="label">Всего</div><div class="value">{len(rows)}</div>
        <div class="muted">{len(by_client)} клиентов · склад {wh_n}</div></div>
      <div class="kpi"><img src="{img_hungry}" alt=""/><div class="label">1 · Нет заявки</div><div class="value">{counts.get('1', 0)}</div></div>
      <div class="kpi"><img src="{img_sick}" alt=""/><div class="label">2 · Нет вырезки &gt;7д</div><div class="value">{counts.get('2', 0)}</div></div>
      <div class="kpi"><img src="{img_hungry}" alt=""/><div class="label">3 · Нет МПТ &gt;14д</div><div class="value">{counts.get('3', 0)}</div></div>
      <div class="kpi"><img src="{img_ok}" alt=""/><div class="label">4 · ТАЗ не обновлён</div><div class="value">{counts.get('4', 0)}</div></div>
      <div class="kpi"><img src="{img_sick}" alt=""/><div class="label">IPC</div><div class="value">{counts.get('ipc', 0)}</div></div>
    </div>
    <p class="muted">
      Приоритет одной тревоги: IPC → ТАЗ не обновлён (действующее МПТ в Перечне, ≤1 год) → нет заявки → нет вырезки → нет МПТ &gt;14д.
      В «Списки по МПТ» при действующем МПТ пишем «ЕСТЬ МПТ» — наличие заявки не важно.
      Исключены позиции, где МПТ есть и в Перечне, и в ТАЗ («Есть МПТ»).
      Тип ВС сверяется мягко («тип ВС ≠» при расхождении). Зелёная строка — «на склад».
    </p>
  </section>

  <div class="tabs">
    {''.join(tabs)}
  </div>

  <h2>По клиентам · pets</h2>
  {''.join(client_blocks)}

  <div class="footer">
    Источники: ТАЗ · Брокерские заявки (Списки на МПТ) · Перечни зап. частей 0% НДС (ФЭИ/ЛР/МЕТЕОР 2025–2026).
    Self-contained HTML — Tamagotchi skin · открывается локально без интернета.
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
