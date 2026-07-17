#!/usr/bin/env python3
"""
Оценка ликвидности склада АТИ по рынку (ТАЗ / ТУЗ / EXPENDABLES).
"""
from __future__ import annotations

import re
import warnings
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from statistics import median
from typing import Any, Iterable, Optional

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")

DATA = Path("/workspace/data")
OUT = Path("/opt/cursor/artifacts/ATI_liquidity_assessment.xlsx")
OUT_COPY = Path("/workspace/output/ATI_liquidity_assessment.xlsx")

SAMPLE_PN = {
    "SAMPLE",
    "ОБРАЗЕЦ",
    "DEFAULT",
    "P/N",
    "PN",
    "PART NUMBER",
    "PARTNUMBER",
    "-",
    "—",
    "N/A",
    "NA",
    "NONE",
    "NULL",
    "TEST",
}
SAMPLE_CLIENT = {
    "DEFAULT",
    "SAMPLE",
    "ОБРАЗЕЦ",
    "TBA",
    "TEST",
    "CLIENT",
}

# Внутренние/технические контрагенты — не считаем рыночным клиентом
INTERNAL_CLIENT_MARKERS = (
    "закупка на склад",
    "на склад",
    "warehouse",
    "stock purchase",
    "internal",
)


def is_market_client(client: str) -> bool:
    if not client:
        return False
    c = client.strip().lower()
    if client.upper() in SAMPLE_CLIENT or c in {x.lower() for x in SAMPLE_CLIENT}:
        return False
    return not any(m in c for m in INTERNAL_CLIENT_MARKERS)


def norm_pn(value: Any) -> str:
    if value is None:
        return ""
    s = str(value).strip().upper()
    if not s:
        return ""
    # unicode dashes / spaces / zero-width
    s = (
        s.replace("\u2212", "-")
        .replace("\u2013", "-")
        .replace("\u2014", "-")
        .replace("\u00a0", "")
        .replace("\u200b", "")
        .replace("\t", "")
        .replace("\n", "")
        .replace("\r", "")
    )
    s = re.sub(r"\s+", "", s)
    # drop trailing punctuation noise
    s = s.strip(".,;:|/\\")
    return s


def soft_pn_key(pn: str) -> str:
    """Ключ без дефисов/точек — ловит опечатки в разделителях."""
    return re.sub(r"[^A-Z0-9]", "", pn)


def norm_client(value: Any) -> str:
    if value is None:
        return ""
    s = str(value).strip()
    s = re.sub(r"\s+", " ", s)
    return s


def is_sample_pn(pn: str) -> bool:
    if not pn or pn in SAMPLE_PN:
        return True
    if not re.search(r"[0-9]", pn):
        return True
    if len(pn) < 2:
        return True
    return False


def parse_money(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        if value != value:  # NaN
            return None
        return float(value)
    s = str(value).strip()
    if not s or s in {"-", "—", "n/a", "N/A"}:
        return None
    s = s.replace("\u00a0", " ").replace("\u202f", " ")
    s = re.sub(r"[^\d,.\-]", "", s)
    if not s or s in {".", ",", "-", "-.", ".-"}:
        return None
    # European: 1.234,56 or 1 234,56 already stripped spaces
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        # 100,00 or 2939,77
        parts = s.split(",")
        if len(parts[-1]) <= 2:
            s = s.replace(",", ".")
        else:
            s = s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


def parse_qty(value: Any) -> float:
    if value is None or value == "":
        return 0.0
    if isinstance(value, (int, float)):
        return float(value) if value == value else 0.0
    m = re.search(r"-?\d+(?:[.,]\d+)?", str(value).replace("\u00a0", ""))
    if not m:
        return 0.0
    return parse_money(m.group(0)) or 0.0


@dataclass
class Event:
    source: str  # TAZ / TUZ / EXP
    kind: str  # order / request
    pn: str
    alt: str = ""
    client: str = ""
    qty: float = 0.0
    price: Optional[float] = None
    condition: str = ""
    date: Any = None
    request_no: str = ""
    sheet: str = ""
    description: str = ""


@dataclass
class MarketAgg:
    order_events: int = 0
    order_qty: float = 0.0
    order_clients: set = field(default_factory=set)
    request_events: int = 0
    request_qty: float = 0.0
    request_clients: set = field(default_factory=set)
    exp_events: int = 0
    exp_qty: float = 0.0
    exp_clients: set = field(default_factory=set)
    order_prices: list = field(default_factory=list)
    request_prices: list = field(default_factory=list)
    exp_prices: list = field(default_factory=list)
    descriptions: set = field(default_factory=set)
    conditions_seen: set = field(default_factory=set)
    matched_via: set = field(default_factory=set)  # exact / soft / alt
    sample_clients_order: list = field(default_factory=list)
    sample_clients_request: list = field(default_factory=list)

    def merge_event(self, e: Event):
        if e.description:
            self.descriptions.add(str(e.description).strip()[:120])
        if e.condition:
            self.conditions_seen.add(str(e.condition).strip().upper())
        if e.kind == "order":
            self.order_events += 1
            self.order_qty += e.qty or 0
            if is_market_client(e.client):
                self.order_clients.add(e.client)
                if len(self.sample_clients_order) < 8 and e.client not in self.sample_clients_order:
                    self.sample_clients_order.append(e.client)
            if e.price is not None and e.price > 0:
                self.order_prices.append(e.price)
        elif e.source == "EXP":
            self.exp_events += 1
            self.exp_qty += e.qty or 0
            if is_market_client(e.client):
                self.exp_clients.add(e.client)
                if len(self.sample_clients_request) < 8 and e.client not in self.sample_clients_request:
                    self.sample_clients_request.append(e.client)
            if e.price is not None and e.price > 0:
                self.exp_prices.append(e.price)
        else:
            self.request_events += 1
            self.request_qty += e.qty or 0
            if is_market_client(e.client):
                self.request_clients.add(e.client)
                if len(self.sample_clients_request) < 8 and e.client not in self.sample_clients_request:
                    self.sample_clients_request.append(e.client)
            if e.price is not None and e.price > 0:
                self.request_prices.append(e.price)


def price_summary(prices: list[float]) -> tuple[Optional[float], Optional[float], Optional[float], int]:
    if not prices:
        return None, None, None, 0
    return min(prices), median(prices), max(prices), len(prices)


def liquidity_score(m: MarketAgg) -> tuple[float, str, str]:
    """Возвращает (score, grade A-D, rationale)."""
    n_ord_cli = len(m.order_clients)
    n_req_cli = len(m.request_clients | m.exp_clients)
    n_all_cli = len(m.order_clients | m.request_clients | m.exp_clients)

    score = (
        m.order_events * 12.0
        + n_ord_cli * 10.0
        + min(m.order_qty, 40) * 0.4
        + m.request_events * 3.0
        + len(m.request_clients) * 5.0
        + m.exp_events * 2.0
        + len(m.exp_clients) * 3.0
        + min(m.request_qty + m.exp_qty, 80) * 0.15
    )

    # лёгкий буст за повторный спрос у нескольких клиентов
    if n_all_cli >= 3:
        score += 8
    if n_all_cli >= 5:
        score += 10

    if m.order_events >= 3 and n_ord_cli >= 2:
        grade = "A"
    elif m.order_events >= 1 and (m.request_events + m.exp_events >= 1 or n_ord_cli >= 1):
        grade = "A" if (m.order_events >= 2 or n_all_cli >= 3) else "B"
    elif m.request_events + m.exp_events >= 5 and n_req_cli >= 2:
        grade = "B"
    elif m.request_events + m.exp_events >= 2 or n_req_cli >= 2:
        grade = "C"
    elif m.request_events + m.exp_events >= 1 or m.order_events >= 1:
        grade = "C"
    else:
        grade = "D"
        score = 0.0

    # refine borderline by score
    if grade != "D":
        if score >= 55 and grade in {"B", "C"}:
            grade = "A"
        elif score >= 25 and grade == "C" and (m.order_events >= 1 or n_req_cli >= 2):
            grade = "B"

    parts = []
    if m.order_events:
        parts.append(
            f"заказов ТАЗ: {m.order_events} (клиентов: {n_ord_cli}"
            + (f" — {', '.join(m.sample_clients_order[:5])}" if m.sample_clients_order else "")
            + f", qty≈{m.order_qty:g})"
        )
    if m.request_events:
        parts.append(
            f"запросов ТУЗ: {m.request_events} (клиентов: {len(m.request_clients)}"
            + (f" — {', '.join(m.sample_clients_request[:5])}" if m.sample_clients_request else "")
            + f", qty≈{m.request_qty:g})"
        )
    if m.exp_events:
        parts.append(
            f"запросов EXP: {m.exp_events} (клиентов: {len(m.exp_clients)}, qty≈{m.exp_qty:g})"
        )
    if not parts:
        parts.append("совпадений по рынку не найдено")

    via = ", ".join(sorted(m.matched_via)) if m.matched_via else "—"
    grade_ru = {"A": "высокая", "B": "средняя", "C": "низкая", "D": "нет спроса"}.get(grade, grade)
    rationale = f"Ликвидность {grade} ({grade_ru}). " + "; ".join(parts) + f". Сопоставление: {via}."
    return round(score, 1), grade, rationale


def find_header_map(rows: list[tuple], aliases: dict[str, list[str]], scan: int = 5) -> tuple[int, dict[str, int]]:
    """Ищет строку заголовков и маппинг колонок."""
    def clean(h):
        if h is None:
            return ""
        return re.sub(r"\s+", " ", str(h).replace("\n", " ")).strip().lower()

    for idx in range(min(scan, len(rows))):
        row = rows[idx]
        if not row:
            continue
        cleaned = [clean(c) for c in row]
        mapping = {}
        for key, names in aliases.items():
            for i, cell in enumerate(cleaned):
                if not cell:
                    continue
                for name in names:
                    if cell == name or name in cell:
                        mapping[key] = i
                        break
                if key in mapping:
                    break
        # need at least pn
        if "pn" in mapping:
            return idx, mapping
    return 0, {}


def iter_sheet_rows(path: Path, sheet: str) -> Iterable[tuple]:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet]
    for row in ws.iter_rows(values_only=True):
        yield row
    wb.close()


def load_taz(path: Path) -> list[Event]:
    events: list[Event] = []
    seen = set()
    for sheet in ["ORDERS"]:
        rows = list(iter_sheet_rows(path, sheet))
        if not rows:
            continue
        # fixed known layout for ORDERS
        # skip header + sample
        for r in rows[1:]:
            if not r or len(r) < 33:
                continue
            pn = norm_pn(r[10])
            if is_sample_pn(pn):
                continue
            client = norm_client(r[6])
            if client.upper() in SAMPLE_CLIENT:
                client = ""
            invoice = str(r[0]).strip() if r[0] is not None else ""
            qty = parse_qty(r[13])
            price = parse_money(r[32])  # продажная ед.
            if price is None:
                price = parse_money(r[30])  # закупка как fallback
            cond = str(r[24]).strip() if r[24] is not None else ""
            alt = norm_pn(r[11])
            desc = str(r[12]).strip() if r[12] is not None else ""
            date = r[16]
            key = ("TAZ", sheet, invoice, pn, client, round(qty, 4), round(price or 0, 2))
            if key in seen:
                continue
            seen.add(key)
            # skip empty-looking rows
            if not client and not invoice and qty == 0:
                continue
            events.append(
                Event(
                    source="TAZ",
                    kind="order",
                    pn=pn,
                    alt=alt if alt != pn else "",
                    client=client,
                    qty=qty,
                    price=price,
                    condition=cond,
                    date=date,
                    request_no=invoice,
                    sheet=sheet,
                    description=desc,
                )
            )
    return events


TUZ_SHEETS = [
    "AFL Group",
    "Группа A",
    "Группа B",
    "Группа C",
    "Группа 3 (old)",
    "Группа 2 (old)",
    "Группа 5 (old)",
    "Questions v2 AFL",
    "Questions v2",
    "MRO",
    "TROUBLES",
    "ASSETS",
    "АФЛ проценка шоп",
]

TUZ_ALIASES = {
    "pn": ["p/n", "part number", "partnumber"],
    "alt": ["alt. p/n", "alt p/n", "alt pn", "alt.pn"],
    "client": ["client", "customer"],
    "desc": ["description"],
    "qty": ["qty"],
    "cond": ["cond", "condition"],
    "date": ["request date"],
    "req": ["request №", "request no", "request #"],
    "price": ["offered", "supplier price", "root price"],
    "invoice": ["invoice to customer"],
}


def load_tuz(path: Path) -> list[Event]:
    events: list[Event] = []
    seen = set()
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    for sheet in TUZ_SHEETS:
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        # read first rows to detect header
        preview = []
        all_rows = []
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            all_rows.append(row)
            if i < 5:
                preview.append(row)
        hdr_idx, mapping = find_header_map(preview, TUZ_ALIASES, scan=5)
        # fallback classic layout
        if not mapping:
            mapping = {
                "date": 1,
                "req": 2,
                "client": 4,
                "pn": 6,
                "alt": 7,
                "desc": 8,
                "qty": 9,
                "cond": 18,
                "price": 24,
                "invoice": 27,
            }
            hdr_idx = 0
        else:
            # ensure classics if missing
            mapping.setdefault("client", 4)
            mapping.setdefault("alt", mapping["pn"] + 1)
            mapping.setdefault("desc", mapping["pn"] + 2)
            mapping.setdefault("qty", mapping["pn"] + 3)

        for r in all_rows[hdr_idx + 1 :]:
            if not r:
                continue
            def get(key, default=None):
                idx = mapping.get(key)
                if idx is None or idx >= len(r):
                    return default
                return r[idx]

            pn_raw = get("pn")
            pn = norm_pn(pn_raw)
            if is_sample_pn(pn):
                continue
            # guard: description wrongly in PN col
            if isinstance(pn_raw, str) and " " in pn_raw.strip() and len(pn_raw) > 40:
                continue

            client = norm_client(get("client"))
            if client.upper() in SAMPLE_CLIENT:
                continue
            # if client looks like a part number and classic shift — try recover
            if client and is_sample_pn(norm_pn(client)) is False and re.fullmatch(r"[A-Z0-9\-./]+", norm_pn(client)) and not re.search(r"[А-Яа-яA-Za-z]{3,}", client.replace(" ", "")):
                # likely shifted; skip unsafe rows rather than corrupt stats
                pass

            alt = norm_pn(get("alt"))
            desc = str(get("desc") or "").strip()
            qty = parse_qty(get("qty"))
            # price: try offered then supplier
            price = parse_money(get("price"))
            if price is None:
                # try known offsets near end
                for idx in (24, 26, 27, 15, 13, 18):
                    if idx < len(r):
                        price = parse_money(r[idx])
                        if price is not None:
                            break
            cond = str(get("cond") or "").strip()
            if cond.lower() in {"cond", "condition"}:
                cond = ""
            date = get("date")
            req = str(get("req") or "").strip()
            invoice = str(get("invoice") or "").strip()

            # dedupe: same request/pn/client/status-noise
            dedupe_key = (sheet, pn, client, req or str(date), round(qty, 4))
            if dedupe_key in seen:
                continue
            seen.add(dedupe_key)

            events.append(
                Event(
                    source="TUZ",
                    kind="request",
                    pn=pn,
                    alt=alt if alt and alt != pn else "",
                    client=client,
                    qty=qty,
                    price=price,
                    condition=cond,
                    date=date,
                    request_no=req or invoice,
                    sheet=sheet,
                    description=desc,
                )
            )
    wb.close()
    return events


EXP_SHEETS = ["EXP NEW", "EXP UTAIR", "EXP GR#1", "EXP GR#2", "EXP GR#3", "EXP GR#4"]


def load_exp(path: Path) -> list[Event]:
    events: list[Event] = []
    seen = set()
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    for sheet in EXP_SHEETS:
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        for i, r in enumerate(ws.iter_rows(values_only=True)):
            if i < 2:
                continue
            if not r or len(r) < 10:
                continue
            pn = norm_pn(r[6])
            if is_sample_pn(pn):
                continue
            client = norm_client(r[3])
            if client.upper() in SAMPLE_CLIENT:
                continue
            alt = norm_pn(r[5]) if len(r) > 5 else ""
            desc = str(r[7]).strip() if r[7] is not None else ""
            qty = parse_qty(r[8])
            market = parse_money(r[9]) if len(r) > 9 else None
            sell = parse_money(r[24]) if len(r) > 24 else None
            price = sell if sell is not None else market
            cond = str(r[14]).strip() if len(r) > 14 and r[14] is not None else ""
            date = r[1] if len(r) > 1 else None
            key = (sheet, pn, client, str(date), round(qty, 4), round(price or 0, 2))
            if key in seen:
                continue
            seen.add(key)
            events.append(
                Event(
                    source="EXP",
                    kind="request",
                    pn=pn,
                    alt=alt if alt and alt != pn else "",
                    client=client,
                    qty=qty,
                    price=price,
                    condition=cond,
                    date=date,
                    sheet=sheet,
                    description=desc,
                )
            )
    wb.close()
    return events


def load_ati(path: Path) -> list[dict]:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = []
    for i, r in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue
        if not r or r[0] is None:
            continue
        rows.append(
            {
                "partno": str(r[0]).strip(),
                "pn": norm_pn(r[0]),
                "serialno": str(r[1]).strip() if r[1] is not None else "",
                "description": str(r[2]).strip() if r[2] is not None else "",
                "ata": str(r[3]).strip() if len(r) > 3 and r[3] is not None else "",
                "ac_typ": str(r[4]).strip() if len(r) > 4 and r[4] is not None else "",
                "condition": str(r[5]).strip() if len(r) > 5 and r[5] is not None else "",
                "qty": parse_qty(r[6]) if len(r) > 6 else 1,
            }
        )
    wb.close()
    return rows


def build_market_index(events: list[Event]) -> tuple[dict[str, MarketAgg], dict[str, set[str]], dict[str, set[str]]]:
    by_pn: dict[str, MarketAgg] = defaultdict(MarketAgg)
    soft_to_pns: dict[str, set[str]] = defaultdict(set)
    alt_to_pns: dict[str, set[str]] = defaultdict(set)

    for e in events:
        by_pn[e.pn].merge_event(e)
        soft_to_pns[soft_pn_key(e.pn)].add(e.pn)
        if e.alt and not is_sample_pn(e.alt):
            alt_to_pns[e.alt].add(e.pn)
            soft_to_pns[soft_pn_key(e.alt)].add(e.pn)
            # also record alt as its own node lightly? No — map only
    return by_pn, soft_to_pns, alt_to_pns


def resolve_market(
    pn: str,
    by_pn: dict[str, MarketAgg],
    soft_to_pns: dict[str, set[str]],
    alt_to_pns: dict[str, set[str]],
) -> MarketAgg:
    """Собирает агрегат рынка для P/N склада с учётом soft/alt."""
    result = MarketAgg()
    matched_pns = set()

    if pn in by_pn:
        matched_pns.add(pn)
        result.matched_via.add("точное P/N")

    # alt references: market events where our PN was listed as ALT of another, or vice versa
    if pn in alt_to_pns:
        matched_pns |= alt_to_pns[pn]
        result.matched_via.add("ALT P/N")

    # soft match only if exact missing or to enlarge
    soft = soft_pn_key(pn)
    candidates = soft_to_pns.get(soft, set())
    if candidates:
        # only accept soft if single cluster or exact already present
        if pn not in by_pn:
            # avoid aggressive overmatch: only if exactly one distinct soft candidate
            # OR candidates include obvious hyphen variants of same
            close = {c for c in candidates if soft_pn_key(c) == soft}
            if len(close) == 1 or all(soft_pn_key(c) == soft for c in close):
                # merge close variants
                for c in close:
                    if c != pn:
                        result.matched_via.add("нормализация (дефисы/пробелы)")
                    matched_pns.add(c)
        else:
            for c in candidates:
                if soft_pn_key(c) == soft and c != pn:
                    # hyphen variant of same
                    if abs(len(c) - len(pn)) <= 2 or soft_pn_key(c) == soft:
                        matched_pns.add(c)
                        result.matched_via.add("нормализация (дефисы/пробелы)")

    for mp in matched_pns:
        src = by_pn.get(mp)
        if not src:
            continue
        # merge fields
        result.order_events += src.order_events
        result.order_qty += src.order_qty
        result.order_clients |= src.order_clients
        result.request_events += src.request_events
        result.request_qty += src.request_qty
        result.request_clients |= src.request_clients
        result.exp_events += src.exp_events
        result.exp_qty += src.exp_qty
        result.exp_clients |= src.exp_clients
        result.order_prices.extend(src.order_prices)
        result.request_prices.extend(src.request_prices)
        result.exp_prices.extend(src.exp_prices)
        result.descriptions |= src.descriptions
        result.conditions_seen |= src.conditions_seen
        for c in src.sample_clients_order:
            if c not in result.sample_clients_order and len(result.sample_clients_order) < 8:
                result.sample_clients_order.append(c)
        for c in src.sample_clients_request:
            if c not in result.sample_clients_request and len(result.sample_clients_request) < 8:
                result.sample_clients_request.append(c)

    if not result.matched_via and matched_pns:
        result.matched_via.add("точное P/N")
    return result


def condition_section(cond: str) -> int:
    c = (cond or "").strip().upper()
    if c == "":
        return 1
    if c in {"US", "NA"}:
        return 2
    return 3


def condition_note(cond: str) -> str:
    c = (cond or "").strip().upper()
    mapping = {
        "": "состояние не указано — риск для продажи, нужна верификация",
        "US": "UNSERVICEABLE — спрос на P/N есть/нет отдельно; для продажи обычно нужен ремонт/обмен",
        "NA": "N/A / не определено — требует уточнения пригодности",
        "N": "NEW",
        "NE": "NEW / NE",
        "SV": "SERVICEABLE",
        "OH": "OVERHAULED",
        "R": "REPAIRED / BER?",
        "S": "SERVICEABLE?",
        "IT": "INSPECTED / TESTED?",
        "I": "INSPECTED?",
        "T": "TESTED?",
        "M": "MODIFIED / MIXED?",
    }
    return mapping.get(c, c)


GRADE_ORDER = {"A": 0, "B": 1, "C": 2, "D": 3}


def fmt_price(v: Optional[float]) -> str:
    if v is None:
        return "—"
    return f"${v:,.2f}"


def build_row(ati: dict, market: MarketAgg) -> dict:
    score, grade, rationale = liquidity_score(market)
    omin, omed, omax, ocnt = price_summary(market.order_prices)
    rmin, rmed, rmax, rcnt = price_summary(market.request_prices + market.exp_prices)
    # prefer order median as market sell ref
    ref_price = omed if omed is not None else rmed
    cond = ati["condition"]
    extra = condition_note(cond)
    if grade != "D":
        rationale = rationale + f" Состояние склада: {extra}."
    else:
        rationale = rationale + f" Состояние склада: {extra}."

    return {
        "liquidity_grade": grade,
        "liquidity_score": score,
        "partno": ati["partno"],
        "serialno": ati["serialno"],
        "description": ati["description"] or (next(iter(market.descriptions), "") if market.descriptions else ""),
        "ac_typ": ati["ac_typ"],
        "ata": ati["ata"],
        "condition": cond if cond else "(пусто)",
        "qty": ati["qty"],
        "taz_orders": market.order_events,
        "taz_clients_n": len(market.order_clients),
        "taz_clients": ", ".join(market.sample_clients_order[:6]),
        "taz_qty": market.order_qty,
        "tuz_requests": market.request_events,
        "tuz_clients_n": len(market.request_clients),
        "tuz_clients": ", ".join([c for c in market.sample_clients_request if c in market.request_clients][:6]),
        "exp_requests": market.exp_events,
        "exp_clients_n": len(market.exp_clients),
        "req_clients_all_n": len(market.request_clients | market.exp_clients),
        "price_ref_usd": ref_price,
        "price_taz_median": omed,
        "price_taz_min": omin,
        "price_taz_max": omax,
        "price_taz_n": ocnt,
        "price_req_median": rmed,
        "price_req_min": rmin,
        "price_req_max": rmax,
        "price_req_n": rcnt,
        "match_via": ", ".join(sorted(market.matched_via)) if market.matched_via else "нет",
        "rationale": rationale,
        "section": condition_section(cond),
    }


HEADERS = [
    ("Ранг", 6),
    ("Ликвидность", 12),
    ("Балл", 8),
    ("P/N", 18),
    ("S/N", 16),
    ("Description", 28),
    ("A/C", 10),
    ("ATA", 8),
    ("Condition", 10),
    ("Qty склад", 9),
    ("Заказов ТАЗ", 11),
    ("Клиентов ТАЗ", 11),
    ("Клиенты (заказы)", 28),
    ("Qty заказов", 10),
    ("Запросов ТУЗ", 11),
    ("Клиентов ТУЗ", 11),
    ("Запросов EXP", 11),
    ("Клиентов EXP", 11),
    ("Цена ориентир USD", 14),
    ("Цена ТАЗ медиана", 13),
    ("Цена ТАЗ min-max", 16),
    ("Цена запросов медиана", 14),
    ("Цена запросов min-max", 16),
    ("Сопоставление", 18),
    ("Обоснование оценки", 70),
]


def style_header(ws, fill_color: str):
    fill = PatternFill("solid", fgColor=fill_color)
    font = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
    thin = Border(
        left=Side(style="thin", color="D0D0D0"),
        right=Side(style="thin", color="D0D0D0"),
        top=Side(style="thin", color="D0D0D0"),
        bottom=Side(style="thin", color="D0D0D0"),
    )
    for col, (name, width) in enumerate(HEADERS, 1):
        cell = ws.cell(1, col, name)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = thin
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.row_dimensions[1].height = 36
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}1"


GRADE_FILL = {
    "A": PatternFill("solid", fgColor="1F7A4D"),
    "B": PatternFill("solid", fgColor="2F6FED"),
    "C": PatternFill("solid", fgColor="C47F00"),
    "D": PatternFill("solid", fgColor="8A8A8A"),
}
GRADE_FONT = Font(bold=True, color="FFFFFF", name="Calibri")
ZEBRA = PatternFill("solid", fgColor="F7F9FC")
THIN = Border(
    left=Side(style="thin", color="E5E5E5"),
    right=Side(style="thin", color="E5E5E5"),
    top=Side(style="thin", color="E5E5E5"),
    bottom=Side(style="thin", color="E5E5E5"),
)


def write_rows(ws, rows: list[dict], header_color: str):
    style_header(ws, header_color)
    for i, r in enumerate(rows, 1):
        price_mm_taz = (
            f"{fmt_price(r['price_taz_min'])} – {fmt_price(r['price_taz_max'])}"
            if r["price_taz_n"]
            else "—"
        )
        price_mm_req = (
            f"{fmt_price(r['price_req_min'])} – {fmt_price(r['price_req_max'])}"
            if r["price_req_n"]
            else "—"
        )
        values = [
            i,
            r["liquidity_grade"],
            r["liquidity_score"],
            r["partno"],
            r["serialno"],
            r["description"],
            r["ac_typ"],
            r["ata"],
            r["condition"],
            r["qty"],
            r["taz_orders"],
            r["taz_clients_n"],
            r["taz_clients"],
            r["taz_qty"],
            r["tuz_requests"],
            r["tuz_clients_n"],
            r["exp_requests"],
            r["exp_clients_n"],
            r["price_ref_usd"],
            r["price_taz_median"],
            price_mm_taz,
            r["price_req_median"],
            price_mm_req,
            r["match_via"],
            r["rationale"],
        ]
        for col, val in enumerate(values, 1):
            cell = ws.cell(i + 1, col, val)
            cell.border = THIN
            cell.alignment = Alignment(vertical="center", wrap_text=(col in {6, 13, 25}))
            if i % 2 == 0:
                cell.fill = ZEBRA
            if col == 2:
                cell.fill = GRADE_FILL.get(r["liquidity_grade"], GRADE_FILL["D"])
                cell.font = GRADE_FONT
                cell.alignment = Alignment(horizontal="center", vertical="center")
            if col in {19, 20, 22} and isinstance(val, (int, float)):
                cell.number_format = '"$"#,##0.00'
        ws.row_dimensions[i + 1].height = 48
    if rows:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}{len(rows)+1}"


def write_summary(ws, ati_rows, scored, taz_n, tuz_n, exp_n, notes: list[str]):
    ws["A1"] = "Оценка ликвидности склада АТИ"
    ws["A1"].font = Font(bold=True, size=16, color="1A1A1A", name="Calibri")
    ws["A2"] = f"Дата отчёта: {datetime.now():%Y-%m-%d %H:%M} | Источники: ТАЗ ORDERS, ТУЗ (группы запросов), EXPENDABLES, АТИ"
    ws.merge_cells("A2:F2")

    summary = [
        ("Позиций на складе АТИ (строк)", len(ati_rows)),
        ("Уникальных P/N на складе", len({r['pn'] for r in ati_rows})),
        ("Раздел 1 — Condition пусто", sum(1 for r in scored if r["section"] == 1)),
        ("Раздел 2 — Condition US/NA", sum(1 for r in scored if r["section"] == 2)),
        ("Раздел 3 — прочие Condition", sum(1 for r in scored if r["section"] == 3)),
        ("Ликвидность A (высокая)", sum(1 for r in scored if r["liquidity_grade"] == "A")),
        ("Ликвидность B (средняя)", sum(1 for r in scored if r["liquidity_grade"] == "B")),
        ("Ликвидность C (низкая)", sum(1 for r in scored if r["liquidity_grade"] == "C")),
        ("Ликвидность D (нет спроса)", sum(1 for r in scored if r["liquidity_grade"] == "D")),
        ("Событий ТАЗ (заказы, после дедупа)", taz_n),
        ("Событий ТУЗ (запросы, после дедупа)", tuz_n),
        ("Событий EXPENDABLES (после дедупа)", exp_n),
    ]
    ws["A4"] = "Сводка"
    ws["A4"].font = Font(bold=True, size=13)
    for i, (k, v) in enumerate(summary, 5):
        ws.cell(i, 1, k).font = Font(name="Calibri", size=11)
        ws.cell(i, 2, v).font = Font(bold=True, name="Calibri", size=11)

    ws["A18"] = "Методика оценки"
    ws["A18"].font = Font(bold=True, size=13)
    method = [
        "A — высокая: подтверждённые заказы и/или устойчивый спрос у нескольких клиентов.",
        "B — средняя: есть заказы или заметные повторные запросы.",
        "C — низкая: единичные/редкие запросы без сильной истории продаж.",
        "D — нет спроса: P/N не найден в ТАЗ/ТУЗ/EXPENDABLES (с учётом нормализации и ALT).",
        "Балл учитывает число заказов/запросов, число уникальных клиентов, объёмы; заказы ТАЗ весят больше запросов.",
        "Цена ориентир: медиана продажной цены ТАЗ; если заказов нет — медиана цен из запросов/EXP.",
        "Склад без цены в АТИ — рыночная цена берётся только из истории спроса/продаж.",
        "US/NA выделены отдельно: спрос на P/N ≠ лёгкая продажа в текущем состоянии.",
    ]
    for i, t in enumerate(method, 19):
        ws.cell(i, 1, t)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=6)

    ws["A28"] = "Замечания по качеству данных (авто)"
    ws["A28"].font = Font(bold=True, size=13)
    for i, t in enumerate(notes, 29):
        ws.cell(i, 1, f"• {t}")
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=6)

    ws.column_dimensions["A"].width = 55
    ws.column_dimensions["B"].width = 18


def main():
    print("Loading TAZ...")
    taz = load_taz(DATA / "TAZ_17.07.2026.xlsx")
    print(f"  TAZ events: {len(taz)}")
    print("Loading TUZ...")
    tuz = load_tuz(DATA / "TUZ_17.07.2026.xlsx")
    print(f"  TUZ events: {len(tuz)}")
    print("Loading EXPENDABLES...")
    exp = load_exp(DATA / "EXPENDABLES.xlsx")
    print(f"  EXP events: {len(exp)}")
    print("Loading ATI...")
    ati = load_ati(DATA / "ATI.xlsx")
    print(f"  ATI rows: {len(ati)}")

    all_events = taz + tuz + exp
    by_pn, soft_to_pns, alt_to_pns = build_market_index(all_events)
    print(f"  Market unique PNs: {len(by_pn)}")

    # data quality notes
    ati_pn_counts = defaultdict(int)
    for r in ati:
        ati_pn_counts[r["pn"]] += 1
    dup_pn = sum(1 for v in ati_pn_counts.values() if v > 1)
    empty_cond = sum(1 for r in ati if not r["condition"])
    notes = [
        f"В АТИ {dup_pn} P/N встречаются более одного раза (разные S/N) — оценка рынка дана на уровне P/N, строки склада сохранены отдельно.",
        f"Пустой Condition в АТИ: {empty_cond} строк.",
        "В ТУЗ обнаружены сдвиги/дубли заголовков и повторные строки одного запроса с разными статусами — выполнен дедуп по (лист, P/N, клиент, №запроса/дата, qty).",
        "В EXPENDABLES несколько вкладок и sample-строки — sample/DEFAULT отфильтрованы, дубли цен/предложений сжаты.",
        "Сопоставление P/N: верхний регистр, удаление пробелов, унификация тире; дополнительно soft-ключ без разделителей и ALT P/N.",
        "Цена в файле АТИ отсутствует — в отчёте рыночные ориентиры из ТАЗ/запросов.",
        "Контрагенты вида «Закупка на склад» / internal не считаются рыночными клиентами (заказы при этом учитываются).",
    ]

    scored = []
    for row in ati:
        market = resolve_market(row["pn"], by_pn, soft_to_pns, alt_to_pns)
        scored.append(build_row(row, market))

    def sort_key(r):
        return (GRADE_ORDER[r["liquidity_grade"]], -r["liquidity_score"], r["partno"], r["serialno"])

    sec1 = sorted([r for r in scored if r["section"] == 1], key=sort_key)
    sec2 = sorted([r for r in scored if r["section"] == 2], key=sort_key)
    sec3 = sorted([r for r in scored if r["section"] == 3], key=sort_key)
    top = sorted([r for r in scored if r["liquidity_grade"] in {"A", "B"}], key=sort_key)[:80]

    print("Writing Excel...")
    wb = openpyxl.Workbook()
    ws0 = wb.active
    ws0.title = "0. Сводка"
    write_summary(ws0, ati, scored, len(taz), len(tuz), len(exp), notes)

    ws_top = wb.create_sheet("ТОП ликвидных")
    write_rows(ws_top, top, "1F7A4D")

    ws1 = wb.create_sheet("1. Condition пусто")
    write_rows(ws1, sec1, "6B4C9A")

    ws2 = wb.create_sheet("2. Condition US_NA")
    write_rows(ws2, sec2, "B33B3B")

    ws3 = wb.create_sheet("3. Прочие Condition")
    write_rows(ws3, sec3, "2F5D9F")

    # compact legend sheet
    wsl = wb.create_sheet("Легенда")
    wsl["A1"] = "Цвета ликвидности"
    wsl["A1"].font = Font(bold=True, size=13)
    for i, (g, name) in enumerate([("A", "Высокая"), ("B", "Средняя"), ("C", "Низкая"), ("D", "Нет спроса")], 3):
        c = wsl.cell(i, 1, g)
        c.fill = GRADE_FILL[g]
        c.font = GRADE_FONT
        wsl.cell(i, 2, name)
    wsl["A8"] = "Разделы сформированы по Condition склада АТИ: 1=пусто, 2=US и NA, 3=все остальные (N/NE/SV/OH/R/...)."
    wsl.column_dimensions["A"].width = 12
    wsl.column_dimensions["B"].width = 20

    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT_COPY.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    wb.save(OUT_COPY)
    print(f"Saved: {OUT}")
    print(f"Saved: {OUT_COPY}")
    print(
        "Counts:",
        f"sec1={len(sec1)} sec2={len(sec2)} sec3={len(sec3)} top={len(top)}",
        f"A={sum(1 for r in scored if r['liquidity_grade']=='A')}",
        f"B={sum(1 for r in scored if r['liquidity_grade']=='B')}",
        f"C={sum(1 for r in scored if r['liquidity_grade']=='C')}",
        f"D={sum(1 for r in scored if r['liquidity_grade']=='D')}",
    )


if __name__ == "__main__":
    main()
