#!/usr/bin/env python3
"""
Оценка ликвидности склада АТИ по рынку (ТАЗ / ТУЗ / EXPENDABLES).
"""
from __future__ import annotations

import re
import warnings
from datetime import datetime, date
from collections import defaultdict
from dataclasses import dataclass, field
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


def day_key(value: Any) -> str:
    """Календарный день для уникальности запроса."""
    if value is None or value == "":
        return "NO_DATE"
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    s = str(value).strip()
    if not s:
        return "NO_DATE"
    try:
        return datetime.fromisoformat(s.replace("Z", "").split(".")[0]).date().isoformat()
    except Exception:
        pass
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(s[:10], fmt).date().isoformat()
        except Exception:
            continue
    return s[:10] if len(s) >= 8 else "NO_DATE"


def client_key(client: str) -> str:
    return (client or "").strip().lower()


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
    # Уникальные заказы: номер заказа (счёта). 1 заказ = P/N + номер заказа
    order_keys: set = field(default_factory=set)
    order_qty: float = 0.0
    order_clients: set = field(default_factory=set)
    # Уникальные запросы ТУЗ+EXP: (клиент, день). 1 запрос = P/N + клиент + день
    request_keys: set = field(default_factory=set)
    request_qty: float = 0.0
    request_clients: set = field(default_factory=set)
    order_prices: list = field(default_factory=list)
    request_prices: list = field(default_factory=list)
    descriptions: set = field(default_factory=set)
    conditions_seen: set = field(default_factory=set)
    matched_via: set = field(default_factory=set)
    sample_clients_order: list = field(default_factory=list)
    sample_clients_request: list = field(default_factory=list)
    # для прозрачности источников
    request_keys_tuz: set = field(default_factory=set)
    request_keys_exp: set = field(default_factory=set)

    @property
    def order_events(self) -> int:
        return len(self.order_keys)

    @property
    def request_events(self) -> int:
        return len(self.request_keys)

    def merge_event(self, e: Event):
        if e.description:
            self.descriptions.add(str(e.description).strip()[:120])
        if e.condition:
            self.conditions_seen.add(str(e.condition).strip().upper())

        if e.kind == "order":
            ono = (e.request_no or "").strip()
            if not ono:
                # без номера заказа — не склеиваем всё в одну кучу
                ono = f"NO_INV|{client_key(e.client)}|{day_key(e.date)}|{round(e.qty or 0, 4)}"
            if ono not in self.order_keys:
                self.order_keys.add(ono)
            self.order_qty += e.qty or 0
            if is_market_client(e.client):
                self.order_clients.add(e.client)
                if len(self.sample_clients_order) < 8 and e.client not in self.sample_clients_order:
                    self.sample_clients_order.append(e.client)
            if e.price is not None and e.price > 0:
                self.order_prices.append(e.price)
            return

        # request from TUZ or EXP — общее правило уникальности
        ck = client_key(e.client)
        dk = day_key(e.date)
        rkey = (ck, dk)
        if rkey not in self.request_keys:
            self.request_keys.add(rkey)
        self.request_qty += e.qty or 0
        if e.source == "EXP":
            self.request_keys_exp.add(rkey)
        else:
            self.request_keys_tuz.add(rkey)
        if is_market_client(e.client):
            self.request_clients.add(e.client)
            if len(self.sample_clients_request) < 8 and e.client not in self.sample_clients_request:
                self.sample_clients_request.append(e.client)
        if e.price is not None and e.price > 0:
            self.request_prices.append(e.price)

    def merge_agg(self, src: "MarketAgg"):
        new_orders = src.order_keys - self.order_keys
        # qty only for newly added order keys — approximate: add proportional avg
        if new_orders and src.order_keys:
            avg_q = src.order_qty / max(len(src.order_keys), 1)
            self.order_qty += avg_q * len(new_orders)
        self.order_keys |= src.order_keys
        self.order_clients |= src.order_clients

        new_reqs = src.request_keys - self.request_keys
        if new_reqs and src.request_keys:
            avg_q = src.request_qty / max(len(src.request_keys), 1)
            self.request_qty += avg_q * len(new_reqs)
        self.request_keys |= src.request_keys
        self.request_clients |= src.request_clients
        self.request_keys_tuz |= src.request_keys_tuz
        self.request_keys_exp |= src.request_keys_exp

        self.order_prices.extend(src.order_prices)
        self.request_prices.extend(src.request_prices)
        self.descriptions |= src.descriptions
        self.conditions_seen |= src.conditions_seen
        for c in src.sample_clients_order:
            if c not in self.sample_clients_order and len(self.sample_clients_order) < 8:
                self.sample_clients_order.append(c)
        for c in src.sample_clients_request:
            if c not in self.sample_clients_request and len(self.sample_clients_request) < 8:
                self.sample_clients_request.append(c)


def price_summary(prices: list[float]) -> tuple[Optional[float], Optional[float], Optional[float], int]:
    if not prices:
        return None, None, None, 0
    return min(prices), median(prices), max(prices), len(prices)


def liquidity_score(m: MarketAgg) -> tuple[float, str, str]:
    """Возвращает (score, grade A-D, rationale)."""
    n_ord_cli = len(m.order_clients)
    n_req_cli = len(m.request_clients)
    n_all_cli = len(m.order_clients | m.request_clients)
    n_req = m.request_events
    n_ord = m.order_events

    score = (
        n_ord * 12.0
        + n_ord_cli * 10.0
        + min(m.order_qty, 40) * 0.4
        + n_req * 3.0
        + n_req_cli * 5.0
        + min(m.request_qty, 80) * 0.15
    )

    if n_all_cli >= 3:
        score += 8
    if n_all_cli >= 5:
        score += 10

    if n_ord >= 3 and n_ord_cli >= 2:
        grade = "A"
    elif n_ord >= 1 and (n_req >= 1 or n_ord_cli >= 1):
        grade = "A" if (n_ord >= 2 or n_all_cli >= 3) else "B"
    elif n_req >= 5 and n_req_cli >= 2:
        grade = "B"
    elif n_req >= 2 or n_req_cli >= 2:
        grade = "C"
    elif n_req >= 1 or n_ord >= 1:
        grade = "C"
    else:
        grade = "D"
        score = 0.0

    if grade != "D":
        if score >= 55 and grade in {"B", "C"}:
            grade = "A"
        elif score >= 25 and grade == "C" and (n_ord >= 1 or n_req_cli >= 2):
            grade = "B"

    parts = []
    if n_ord:
        parts.append(
            f"заказов ТАЗ: {n_ord} (уник. №заказа+P/N; клиентов: {n_ord_cli}"
            + (f" — {', '.join(m.sample_clients_order[:5])}" if m.sample_clients_order else "")
            + f", qty≈{m.order_qty:g})"
        )
    if n_req:
        parts.append(
            f"запросов ТУЗ+EXP: {n_req} (уник. P/N+клиент+день; клиентов: {n_req_cli}"
            + (f" — {', '.join(m.sample_clients_request[:5])}" if m.sample_clients_request else "")
            + f", qty≈{m.request_qty:g})"
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
            # Только продажная цена клиенту — без fallback на закупку
            price = parse_money(r[32])  # Продажная, ед.
            cond = str(r[24]).strip() if r[24] is not None else ""
            alt = norm_pn(r[11])
            desc = str(r[12]).strip() if r[12] is not None else ""
            date = r[16]
            key = ("TAZ", pn, invoice)  # 1 заказ = P/N + номер заказа
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
    "price": ["offered\nper unit", "offered per unit", "offered"],
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
            # Только Offered per unit $ — без supplier/root
            price = None
            price_idx = mapping.get("price")
            if price_idx is not None and price_idx < len(r):
                price = parse_money(r[price_idx])
            else:
                # classic ≈24; с Sales/Purch ID ≈27
                for idx in (24, 27):
                    if idx < len(r):
                        cand = parse_money(r[idx])
                        if cand is not None:
                            price = cand
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
            # Только Sell Price EA — без fallback на Market Price EA / Purchase
            price = parse_money(r[24]) if len(r) > 24 else None
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
        result.merge_agg(src)

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


def build_row_from_stock(stock: dict, market: MarketAgg) -> dict:
    score, grade, rationale = liquidity_score(market)
    omin, omed, omax, ocnt = price_summary(market.order_prices)
    rmin, rmed, rmax, rcnt = price_summary(market.request_prices)
    ref_price = omed if omed is not None else rmed
    has_demand = (market.order_events + market.request_events) > 0
    has_sell_offered = ref_price is not None
    if not has_demand:
        price_flag = "н/п (нет спроса)"
    elif has_sell_offered:
        price_flag = "есть sell/offered"
    else:
        price_flag = "НЕТ sell/offered — цена не оценена"

    conds = sorted(stock["conditions"])
    cond_display = ", ".join(c if c else "(пусто)" for c in conds) if conds else "(пусто)"
    notes_cond = [condition_note(c) for c in conds]
    extra = "; ".join(dict.fromkeys(notes_cond))

    rationale = rationale + f" На складе: {stock['qty']:g} шт. ({stock['lines']} строк АТИ)."
    rationale = rationale + f" Состояние склада: {extra}."
    if has_demand and not has_sell_offered:
        rationale += (
            " ВНИМАНИЕ: по рынку есть спрос/заказы, но нет продажной цены ТАЗ "
            "(Продажная, ед.) и нет Offered per unit $ / Sell Price EA — "
            "денежную оценку позиции выполнить нельзя."
        )

    desc = stock["description"] or (
        next(iter(market.descriptions), "") if market.descriptions else ""
    )
    ac = ", ".join(sorted(x for x in stock["ac_typs"] if x)) or ""

    return {
        "liquidity_grade": grade,
        "liquidity_score": score,
        "partno": stock["partno"],
        "pn": stock["pn"],
        "description": desc,
        "ac_typ": ac,
        "condition": cond_display,
        "qty": stock["qty"],
        "lines": stock["lines"],
        "taz_orders": market.order_events,
        "taz_clients_n": len(market.order_clients),
        "taz_clients": ", ".join(market.sample_clients_order[:6]),
        "requests": market.request_events,
        "req_clients_n": len(market.request_clients),
        "req_clients": ", ".join(market.sample_clients_request[:6]),
        "req_tuz_part": len(market.request_keys_tuz),
        "req_exp_part": len(market.request_keys_exp),
        "price_flag": price_flag,
        "has_sell_offered": has_sell_offered,
        "has_demand": has_demand,
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
        "section": stock["section"],
    }


def aggregate_ati_stock(ati_rows: list[dict]) -> list[dict]:
    """Одна строка на P/N внутри раздела Condition; qty = сумма штук."""
    groups: dict[tuple, dict] = {}
    for r in ati_rows:
        sec = condition_section(r["condition"])
        key = (r["pn"], sec)
        if key not in groups:
            groups[key] = {
                "pn": r["pn"],
                "partno": r["partno"],
                "qty": 0.0,
                "lines": 0,
                "conditions": set(),
                "description": r.get("description") or "",
                "ac_typs": set(),
                "section": sec,
            }
        g = groups[key]
        q = r["qty"] if r.get("qty") and r["qty"] > 0 else 1.0
        g["qty"] += q
        g["lines"] += 1
        g["conditions"].add((r.get("condition") or "").strip().upper())
        if r.get("description") and not g["description"]:
            g["description"] = r["description"]
        if r.get("ac_typ"):
            g["ac_typs"].add(r["ac_typ"].strip())
        # keep a readable partno (first seen)
    return list(groups.values())


HEADERS = [
    ("Ранг", 6),
    ("Ликвидность", 12),
    ("Балл", 8),
    ("P/N", 18),
    ("Description", 32),
    ("A/C", 12),
    ("Condition", 12),
    ("Кол-во на складе", 12),
    ("Заказов ТАЗ", 11),
    ("Клиентов ТАЗ", 11),
    ("Клиенты (заказы)", 26),
    ("Запросов ТУЗ+EXP", 14),
    ("Клиентов запросов", 12),
    ("Клиенты (запросы)", 26),
    ("Флаг цены sell/offered", 18),
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


PRICE_MISSING_FILL = PatternFill("solid", fgColor="F8D7DA")
PRICE_OK_FILL = PatternFill("solid", fgColor="D4EDDA")


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
            r["description"],
            r["ac_typ"],
            r["condition"],
            r["qty"],
            r["taz_orders"],
            r["taz_clients_n"],
            r["taz_clients"],
            r["requests"],
            r["req_clients_n"],
            r["req_clients"],
            r["price_flag"],
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
            cell.alignment = Alignment(vertical="center", wrap_text=(col in {5, 11, 14, 22}))
            if i % 2 == 0:
                cell.fill = ZEBRA
            if col == 2:
                cell.fill = GRADE_FILL.get(r["liquidity_grade"], GRADE_FILL["D"])
                cell.font = GRADE_FONT
                cell.alignment = Alignment(horizontal="center", vertical="center")
            if col == 8:
                cell.font = Font(bold=True, name="Calibri", size=11)
                cell.alignment = Alignment(horizontal="center", vertical="center")
            if col == 15:
                if r["has_demand"] and not r["has_sell_offered"]:
                    cell.fill = PRICE_MISSING_FILL
                    cell.font = Font(bold=True, color="721C24", name="Calibri")
                elif r["has_sell_offered"]:
                    cell.fill = PRICE_OK_FILL
            if col in {16, 17, 19} and isinstance(val, (int, float)):
                cell.number_format = '"$"#,##0.00'
        ws.row_dimensions[i + 1].height = 48
    if rows:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}{len(rows)+1}"


def write_summary(ws, ati_rows, scored, taz_n, tuz_n, exp_n, notes: list[str]):
    ws["A1"] = "Оценка ликвидности склада АТИ"
    ws["A1"].font = Font(bold=True, size=16, color="1A1A1A", name="Calibri")
    ws["A2"] = f"Дата отчёта: {datetime.now():%Y-%m-%d %H:%M} | Источники: ТАЗ ORDERS, ТУЗ (группы запросов), EXPENDABLES, АТИ"
    ws.merge_cells("A2:F2")

    no_price = sum(1 for r in scored if r["has_demand"] and not r["has_sell_offered"])
    with_price = sum(1 for r in scored if r["has_sell_offered"])

    summary = [
        ("Строк в исходном АТИ", len(ati_rows)),
        ("Позиций в отчёте (P/N после агрегации)", len(scored)),
        ("Уникальных P/N на складе", len({r["pn"] for r in scored})),
        ("Суммарное кол-во на складе, шт.", sum(r["qty"] for r in scored)),
        ("Раздел 1 — Condition пусто", sum(1 for r in scored if r["section"] == 1)),
        ("Раздел 2 — Condition US/NA", sum(1 for r in scored if r["section"] == 2)),
        ("Раздел 3 — прочие Condition", sum(1 for r in scored if r["section"] == 3)),
        ("Ликвидность A (высокая)", sum(1 for r in scored if r["liquidity_grade"] == "A")),
        ("Ликвидность B (средняя)", sum(1 for r in scored if r["liquidity_grade"] == "B")),
        ("Ликвидность C (низкая)", sum(1 for r in scored if r["liquidity_grade"] == "C")),
        ("Ликвидность D (нет спроса)", sum(1 for r in scored if r["liquidity_grade"] == "D")),
        ("Есть sell/offered цена", with_price),
        ("Спрос есть, но НЕТ sell/offered (цена не оценена)", no_price),
        ("Событий ТАЗ (заказы, после дедупа)", taz_n),
        ("Событий ТУЗ (запросы, после дедупа)", tuz_n),
        ("Событий EXPENDABLES (после дедупа)", exp_n),
    ]
    ws["A4"] = "Сводка"
    ws["A4"].font = Font(bold=True, size=13)
    for i, (k, v) in enumerate(summary, 5):
        ws.cell(i, 1, k).font = Font(name="Calibri", size=11)
        cell = ws.cell(i, 2, v)
        cell.font = Font(bold=True, name="Calibri", size=11)
        if "НЕТ sell/offered" in k:
            cell.fill = PRICE_MISSING_FILL

    ws["A23"] = "Методика оценки"
    ws["A23"].font = Font(bold=True, size=13)
    method = [
        "A — высокая: подтверждённые заказы и/или устойчивый спрос у нескольких клиентов.",
        "B — средняя: есть заказы или заметные повторные запросы.",
        "C — низкая: единичные/редкие запросы без сильной истории продаж.",
        "D — нет спроса: P/N не найден в ТАЗ/ТУЗ/EXPENDABLES (с учётом нормализации и ALT).",
        "Балл учитывает число заказов/запросов, число уникальных клиентов, объёмы; заказы ТАЗ весят больше запросов.",
        "Одинаковые P/N на складе сведены в одну строку внутри раздела Condition; «Кол-во на складе» = сумма штук.",
        "Заказы ТАЗ: 1 заказ = P/N + номер заказа (счёта).",
        "Запросы ТУЗ+EXP: 1 запрос = P/N + клиент + календарный день (источники объединены).",
        "Цена ориентир ТОЛЬКО из: ТАЗ «Продажная, ед.» / ТУЗ «Offered per unit $» / EXP «Sell Price EA».",
        "Закупка, Supplier Price, Root Price, Market Price EA в ориентир НЕ входят.",
        "Колонка «Флаг цены sell/offered»: если спрос есть, а этих цен нет — позиция помечена, денежная оценка невозможна.",
        "US/NA выделены отдельно: спрос на P/N ≠ лёгкая продажа в текущем состоянии.",
    ]
    for i, t in enumerate(method, 24):
        ws.cell(i, 1, t)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=6)

    ws["A36"] = "Замечания по качеству данных (авто)"
    ws["A36"].font = Font(bold=True, size=13)
    for i, t in enumerate(notes, 37):
        ws.cell(i, 1, f"• {t}")
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=6)

    ws.column_dimensions["A"].width = 62
    ws.column_dimensions["B"].width = 18


def main():
    print("Loading TAZ...")
    taz = load_taz(DATA / "TAZ_17.07.2026.xlsx")
    print(f"  TAZ unique orders (P/N+№заказа preload): {len(taz)}")
    print("Loading TUZ...")
    tuz = load_tuz(DATA / "TUZ_17.07.2026.xlsx")
    print(f"  TUZ raw rows kept: {len(tuz)}")
    print("Loading EXPENDABLES...")
    exp_path = DATA / "EXPENDABLES.xlsx"
    if exp_path.exists():
        exp = load_exp(exp_path)
        print(f"  EXP raw rows kept: {len(exp)}")
    else:
        exp = []
        print("  WARNING: EXPENDABLES.xlsx not found — requests counted from TUZ only")
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
    stock = aggregate_ati_stock(ati)
    notes = [
        f"В АТИ {dup_pn} P/N встречались более одного раза — в отчёте одна строка на P/N внутри раздела Condition, кол-во = сумма штук.",
        f"Исходных строк АТИ: {len(ati)}; после агрегации позиций: {len(stock)}.",
        "Правило запросов: 1 запрос = P/N + клиент + календарный день (ТУЗ и EXP вместе).",
        "Правило заказов ТАЗ: 1 заказ = P/N + номер заказа (счёта).",
        "Колонки ATA и S/N убраны; запросы ТУЗ и EXP объединены в один столбец.",
        "Сопоставление P/N: регистр, пробелы, тире; soft-ключ и ALT P/N.",
        "Цена ориентир только из «Продажная, ед.» / «Offered per unit $» / «Sell Price EA».",
        "«Закупка на склад» не считается рыночным клиентом.",
    ]
    if not exp:
        notes.insert(0, "ВНИМАНИЕ: файл EXPENDABLES отсутствует в среде — столбец запросов пока без EXP.")

    scored = []
    for item in stock:
        market = resolve_market(item["pn"], by_pn, soft_to_pns, alt_to_pns)
        scored.append(build_row_from_stock(item, market))

    def sort_key(r):
        return (GRADE_ORDER[r["liquidity_grade"]], -r["liquidity_score"], -r["qty"], r["partno"])

    sec1 = sorted([r for r in scored if r["section"] == 1], key=sort_key)
    sec2 = sorted([r for r in scored if r["section"] == 2], key=sort_key)
    sec3 = sorted([r for r in scored if r["section"] == 3], key=sort_key)

    top_by_pn: dict[str, dict] = {}
    for r in scored:
        if r["liquidity_grade"] not in {"A", "B"}:
            continue
        pn = r["pn"]
        if pn not in top_by_pn:
            top_by_pn[pn] = dict(r)
        else:
            cur = top_by_pn[pn]
            cur["qty"] += r["qty"]
            cur["lines"] = cur.get("lines", 0) + r.get("lines", 0)
            conds = {x.strip() for x in cur["condition"].split(",") if x.strip()}
            conds |= {x.strip() for x in r["condition"].split(",") if x.strip()}
            cur["condition"] = ", ".join(sorted(conds))
            if GRADE_ORDER[r["liquidity_grade"]] < GRADE_ORDER[cur["liquidity_grade"]] or (
                r["liquidity_grade"] == cur["liquidity_grade"] and r["liquidity_score"] > cur["liquidity_score"]
            ):
                for k in (
                    "liquidity_grade",
                    "liquidity_score",
                    "taz_orders",
                    "taz_clients_n",
                    "taz_clients",
                    "requests",
                    "req_clients_n",
                    "req_clients",
                    "price_flag",
                    "has_sell_offered",
                    "has_demand",
                    "price_ref_usd",
                    "price_taz_median",
                    "price_taz_min",
                    "price_taz_max",
                    "price_taz_n",
                    "price_req_median",
                    "price_req_min",
                    "price_req_max",
                    "price_req_n",
                    "match_via",
                    "rationale",
                    "description",
                    "ac_typ",
                ):
                    cur[k] = r[k]
                cur["rationale"] = (
                    f"{r['rationale']} (в ТОП qty суммирован по всем состояниям склада: {cur['qty']:g} шт.)"
                )
    top = sorted(top_by_pn.values(), key=sort_key)[:80]

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
    wsl["A8"] = "Флаг цены sell/offered"
    wsl["A8"].font = Font(bold=True, size=13)
    c = wsl.cell(9, 1, "есть")
    c.fill = PRICE_OK_FILL
    wsl.cell(9, 2, "есть Продажная ед. / Offered / Sell Price EA")
    c = wsl.cell(10, 1, "НЕТ")
    c.fill = PRICE_MISSING_FILL
    wsl.cell(10, 2, "спрос есть, но sell/offered цены нет — денежная оценка невозможна")
    wsl["A12"] = "Разделы по Condition склада АТИ: 1=пусто, 2=US и NA, 3=прочие (N/NE/SV/OH/R/...)."
    wsl["A13"] = "В цену ориентир НЕ входят: Закупка, Supplier/Root Price, Market Price EA."
    wsl.column_dimensions["A"].width = 12
    wsl.column_dimensions["B"].width = 70

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
        f"no_sell_offered={sum(1 for r in scored if r['has_demand'] and not r['has_sell_offered'])}",
    )


if __name__ == "__main__":
    main()
