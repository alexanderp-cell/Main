#!/usr/bin/env python3
"""Top-10 stock positions by market demand (TAZ orders + TUZ requests + Expendables RFQ)
with indicative sell price from our orders/offers.
"""
from __future__ import annotations

import csv
import re
import statistics
import warnings
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")

DATA = Path("/workspace/data")
STOCK = DATA / "STOCK.xlsx"
TAZ = DATA / "TAZ.xlsx"
TUZ = DATA / "TUZ_30.10.2026.xlsx"
EXP = DATA / "EXPENDABLES.csv"
OUT = Path("/workspace/output/TOP10_stock_demand_prices.xlsx")
OUT_ART = Path("/opt/cursor/artifacts/TOP10_stock_demand_prices.xlsx")

SAMPLE_PN = {
    "SAMPLE", "ОБРАЗЕЦ", "DEFAULT", "P/N", "PN", "PART NUMBER", "PARTNUMBER",
    "-", "—", "N/A", "NA", "NONE", "NULL", "TEST", "NSN", "TBA",
}
SAMPLE_CLIENT = {"DEFAULT", "SAMPLE", "ОБРАЗЕЦ", "TBA", "TEST", "CLIENT", "N/A"}
INTERNAL_CLIENT_MARKERS = (
    "закупка на склад", "на склад", "warehouse", "stock purchase", "internal",
)
TAZ_EXCLUDED = {"5 CANCELLED", "7 REFUND"}
TUZ_SKIP = {
    "suppliers list", "dropdown", "customer list",
    "transport rates and terms", "assets", "troubles",
}


def norm_pn(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float):
        if value != value:
            return ""
        if value == int(value) and abs(value) < 1e15:
            s = str(int(value))
        else:
            s = str(value).strip().upper()
    elif isinstance(value, int):
        s = str(value)
    else:
        s = str(value).strip().upper()
    if not s:
        return ""
    s = (
        s.replace("\u2212", "-").replace("\u2013", "-").replace("\u2014", "-")
        .replace("\u00a0", "").replace("\u200b", "")
        .replace("\t", "").replace("\n", "").replace("\r", "")
    )
    s = re.sub(r"\s+", "", s)
    if re.fullmatch(r"\d+\.0+", s):
        s = s.split(".", 1)[0]
    return s.strip(".,;:|/\\")


def soft_pn_key(pn: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", pn)


def is_sample_pn(pn: str) -> bool:
    if not pn or len(soft_pn_key(pn)) < 3:
        return True
    return pn.upper() in SAMPLE_PN


def norm_client(value: Any) -> str:
    if value is None:
        return ""
    return str(value).replace("\xa0", " ").strip()


def is_market_client(client: str) -> bool:
    if not client:
        return False
    c = client.strip().lower()
    if client.upper() in SAMPLE_CLIENT or c in {x.lower() for x in SAMPLE_CLIENT}:
        return False
    return not any(m in c for m in INTERNAL_CLIENT_MARKERS)


def is_plausible_client(client: str, pn: str = "") -> bool:
    if not client or not is_market_client(client):
        return False
    c = client.strip()
    if re.fullmatch(r"\d+(\.0+)?", c.replace(" ", "").replace("\u00a0", "")):
        return False
    if pn and soft_pn_key(c.upper()) == soft_pn_key(pn):
        return False
    return len(c) >= 2


def parse_money(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        if value != value:
            return None
        return float(value) if value != 0 else None
    s = str(value).replace("\xa0", "").replace(" ", "").strip()
    if not s or s.upper() in {"TBA", "N/A", "NA", "-", "—", "#VALUE!", "#REF!"}:
        return None
    for sep in ("/", "–"):
        if sep in s and re.fullmatch(r"[\d.,]+" + re.escape(sep) + r"[\d.,]+", s):
            parts = []
            for p in s.split(sep):
                try:
                    parts.append(float(p.replace(",", ".")))
                except ValueError:
                    pass
            return min(parts) if parts else None
    s = s.replace("$", "").replace("€", "").replace("USD", "")
    # European format 1.234,56 or 1234,56
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s and s.count(",") == 1:
        left, right = s.split(",")
        if len(right) <= 2:
            s = left.replace(".", "") + "." + right
        else:
            s = s.replace(",", "")
    s = re.sub(r"[^\d.\-]", "", s)
    try:
        v = float(s)
        return v if v != 0 else None
    except ValueError:
        return None


def parse_qty(value: Any) -> float:
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value) if value == value else 0.0
    s = str(value).replace("\xa0", "").replace(" ", "").replace(",", ".").strip()
    try:
        return float(re.sub(r"[^\d.\-]", "", s) or 0)
    except ValueError:
        return 0.0


def clean_header_cell(h: Any) -> str:
    if h is None:
        return ""
    return re.sub(r"\s+", " ", str(h).replace("\n", " ").replace("\xa0", " ")).strip().lower()


def _header_map(row: tuple, aliases: dict[str, list[str]]) -> dict[str, int]:
    cleaned = [clean_header_cell(c) for c in row]
    mapping: dict[str, int] = {}
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
    return mapping


def find_header_map(rows: list[tuple], aliases: dict[str, list[str]], scan: int = 5):
    for idx in range(min(scan, len(rows))):
        mapping = _header_map(rows[idx], aliases)
        if "pn" in mapping:
            return idx, mapping
    return None, {}


def median(vals: list[float]) -> Optional[float]:
    vals = [v for v in vals if v is not None and v > 0]
    if not vals:
        return None
    return float(statistics.median(vals))


def client_key(c: str) -> str:
    return re.sub(r"\s+", " ", c.strip().lower())


def day_key(dt: Any) -> str:
    if dt is None:
        return ""
    if hasattr(dt, "date"):
        try:
            return dt.date().isoformat()
        except Exception:
            pass
    s = str(dt)
    m = re.search(r"(\d{4}-\d{2}-\d{2})", s)
    return m.group(1) if m else s[:10]


# ---------------------------------------------------------------------------
# Loaders
# ---------------------------------------------------------------------------

@dataclass
class StockLine:
    pn: str
    alt: str
    desc: str
    qty: float
    cond: str
    ac: str


@dataclass
class Event:
    kind: str  # order | request | rfq
    pn: str
    alt: str
    client: str
    qty: float
    price: Optional[float]
    desc: str = ""
    cond: str = ""
    purchase: Optional[float] = None
    market_price: Optional[float] = None


@dataclass
class MarketAgg:
    n_orders: int = 0
    n_requests: int = 0
    n_rfq: int = 0
    clients_ord: set = field(default_factory=set)
    clients_req: set = field(default_factory=set)
    clients_rfq: set = field(default_factory=set)
    order_prices: list = field(default_factory=list)
    offer_prices: list = field(default_factory=list)
    exp_sell_prices: list = field(default_factory=list)
    exp_purchase_prices: list = field(default_factory=list)
    exp_market_prices: list = field(default_factory=list)
    order_qty: float = 0.0
    request_qty: float = 0.0
    descriptions: set = field(default_factory=set)
    conditions: set = field(default_factory=set)

    @property
    def demand_score(self) -> float:
        # Prior methodology + RFQ weight similar to TUZ requests
        return (
            self.n_orders * 12
            + len(self.clients_ord) * 10
            + self.n_requests * 3
            + len(self.clients_req) * 5
            + self.n_rfq * 2
            + len(self.clients_rfq) * 3
        )

    def add_event(self, e: Event):
        if e.desc:
            self.descriptions.add(e.desc)
        if e.cond:
            self.conditions.add(e.cond)
        if e.kind == "order":
            self.n_orders += 1
            if e.client:
                self.clients_ord.add(e.client)
            self.order_qty += e.qty
            if e.price:
                self.order_prices.append(e.price)
        elif e.kind == "request":
            self.n_requests += 1
            if e.client:
                self.clients_req.add(e.client)
            self.request_qty += e.qty
            if e.price:
                self.offer_prices.append(e.price)
        elif e.kind == "rfq":
            self.n_rfq += 1
            if e.client:
                self.clients_rfq.add(e.client)
            if e.price:
                self.exp_sell_prices.append(e.price)
            if e.purchase:
                self.exp_purchase_prices.append(e.purchase)
            if e.market_price:
                self.exp_market_prices.append(e.market_price)


def load_stock(path: Path) -> list[StockLine]:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    out: list[StockLine] = []
    for r in rows[1:]:
        if not r or not r[0]:
            continue
        pn = norm_pn(r[0])
        if is_sample_pn(pn):
            continue
        alt_raw = str(r[1] or "")
        alts = [norm_pn(x) for x in re.split(r"[,;/]", alt_raw) if norm_pn(x)]
        out.append(StockLine(
            pn=pn,
            alt=",".join(a for a in alts if a and a != pn),
            desc=str(r[2] or "").strip(),
            qty=parse_qty(r[3]),
            cond=str(r[4] or "").strip(),
            ac=str(r[5] or "").strip(),
        ))
    return out


TAZ_ALIASES = {
    "invoice": ["номер счета", "номер счёта", "счет", "счёт"],
    "customer": ["client", "customer"],
    "pn": ["p/n", "part number", "partnumber"],
    "alt": ["alt p/n", "alt. p/n"],
    "desc": ["description"],
    "qty": ["qty in po", "qtyinpo"],
    "date": ["заказ взят в работу", "заказ клиента взят"],
    "cond": ["condition"],
    "sell": ["продажная, ед"],
    "status": ["status"],
    "category": ["category"],
}


def load_taz_events(path: Path) -> list[Event]:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["ORDERS"]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    mapping = _header_map(rows[0], TAZ_ALIASES)
    if "pn" not in mapping:
        for idx in range(min(3, len(rows))):
            mapping = _header_map(rows[idx], TAZ_ALIASES)
            if "pn" in mapping:
                break
    start = 1
    for i, row in enumerate(rows[:5]):
        if row and any(clean_header_cell(c) in {"p/n", "part number", "pn"} for c in row[:15] if c):
            start = i + 1

    def cell(r, key, default=None):
        idx = mapping.get(key)
        if idx is None or idx >= len(r):
            return default
        return r[idx]

    seen = set()
    events: list[Event] = []
    for r in rows[start:]:
        if not r:
            continue
        status = str(cell(r, "status") or "").strip().upper()
        if status in TAZ_EXCLUDED or "CANCEL" in status or "REFUND" in status:
            continue
        pn = norm_pn(cell(r, "pn"))
        if is_sample_pn(pn):
            continue
        client = norm_client(cell(r, "customer"))
        if not is_market_client(client):
            continue
        invoice = str(cell(r, "invoice") or "").strip()
        key = (pn, invoice or f"NO_INV|{client_key(client)}|{day_key(cell(r, 'date'))}")
        if key in seen:
            continue
        seen.add(key)
        events.append(Event(
            kind="order",
            pn=pn,
            alt=norm_pn(cell(r, "alt")),
            client=client_key(client),
            qty=parse_qty(cell(r, "qty")),
            price=parse_money(cell(r, "sell")),
            desc=str(cell(r, "desc") or "").strip(),
            cond=str(cell(r, "cond") or "").strip(),
        ))
    return events


TUZ_ALIASES = {
    "pn": ["p/n", "part number", "partnumber"],
    "alt": ["alt. p/n", "alt p/n", "alt pn"],
    "client": ["client", "customer"],
    "desc": ["description"],
    "qty": ["qty"],
    "cond": ["cond", "condition"],
    "date": ["request date"],
    "req": ["request №", "request no", "request #"],
    "price": ["offered per unit", "offered"],
}


def _tuz_layouts(primary: dict[str, int]) -> list[dict[str, int]]:
    classic = {
        "date": 1, "req": 2, "client": 4, "ac": 5, "pn": 6, "alt": 7, "desc": 8, "qty": 9,
        "cond": 21, "price": 28,
    }
    classic_q = {
        "date": 3, "req": 4, "client": 6, "ac": 7, "pn": 8, "alt": 9, "desc": 10, "qty": 11,
        "cond": 20, "price": 26,
    }
    shifted = {
        "date": 3, "req": 4, "client": 6, "ac": 7, "pn": 8, "alt": 9, "desc": 10, "qty": 11,
        "cond": 21, "price": 28,
    }
    layouts = []
    if primary:
        layouts.append(dict(primary))
    for lay in (classic, shifted, classic_q):
        if lay not in layouts:
            layouts.append(lay)
    return layouts


def _extract_tuz_row(r: tuple, layouts: list[dict[str, int]]) -> Optional[dict]:
    for mapping in layouts:
        def get(key, default=None, _m=mapping):
            idx = _m.get(key)
            if idx is None or idx >= len(r):
                return default
            return r[idx]

        pn_raw = get("pn")
        pn = norm_pn(pn_raw)
        if is_sample_pn(pn):
            continue
        if isinstance(pn_raw, str) and " " in pn_raw.strip() and len(pn_raw) > 40:
            continue
        client = norm_client(get("client"))
        if not is_plausible_client(client, pn):
            continue
        if pn.lower() in {"first", "repeat", "crit", "aog", "routine"}:
            continue
        alt = norm_pn(get("alt"))
        return {
            "pn": pn,
            "alt": alt if alt and alt != pn else "",
            "client": client,
            "desc": str(get("desc") or "").strip(),
            "qty": parse_qty(get("qty")),
            "price": parse_money(get("price")),
            "cond": str(get("cond") or "").strip(),
            "req": str(get("req") or "").strip(),
            "date": get("date"),
        }
    return None


def load_tuz_events(path: Path) -> list[Event]:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    seen = set()
    events: list[Event] = []
    for sheet in wb.sheetnames:
        if sheet.strip().lower() in TUZ_SKIP:
            continue
        ws = wb[sheet]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        preview = rows[:5]
        hdr_idx, mapping = find_header_map(preview, TUZ_ALIASES, scan=5)
        layouts = _tuz_layouts(mapping)
        start = (hdr_idx + 1) if hdr_idx is not None else 1
        for r in rows[start:]:
            if not r:
                continue
            parsed = _extract_tuz_row(r, layouts)
            if not parsed:
                continue
            pn = parsed["pn"]
            dedup = (
                pn, client_key(parsed["client"]), day_key(parsed["date"]),
                parsed["req"], round(parsed["qty"], 4),
            )
            if dedup in seen:
                continue
            seen.add(dedup)
            events.append(Event(
                kind="request",
                pn=pn,
                alt=parsed["alt"],
                client=client_key(parsed["client"]),
                qty=parsed["qty"],
                price=parsed["price"],
                desc=parsed["desc"],
                cond=parsed["cond"],
            ))
    wb.close()
    return events


def load_exp_events(path: Path) -> list[Event]:
    with open(path, newline="", encoding="utf-8") as f:
        rows = list(csv.reader(f))
    seen = set()
    events: list[Event] = []
    for r in rows[2:]:
        if len(r) < 10:
            continue
        pn = norm_pn(r[6] if len(r) > 6 else "")
        if is_sample_pn(pn):
            continue
        client = norm_client(r[3] if len(r) > 3 else "")
        alt = norm_pn(r[5] if len(r) > 5 else "")
        qty = parse_qty(r[8] if len(r) > 8 else None)
        market = parse_money(r[9] if len(r) > 9 else None)
        purchase = parse_money(r[19] if len(r) > 19 else None)
        sell = parse_money(r[24] if len(r) > 24 else None)
        cond = str(r[14] if len(r) > 14 else "").strip()
        desc = str(r[7] if len(r) > 7 else "").strip()
        rfq_date = str(r[1] if len(r) > 1 else "").strip()
        # one RFQ event per pn+client+date
        event_key = (pn, client_key(client), rfq_date)
        if event_key in seen:
            # still collect extra price quotes on same RFQ
            pass
        else:
            seen.add(event_key)
            events.append(Event(
                kind="rfq",
                pn=pn,
                alt=alt if alt and alt != pn else "",
                client=client_key(client) if client and is_market_client(client) else "",
                qty=qty,
                price=sell,
                desc=desc,
                cond=cond,
                purchase=purchase,
                market_price=market,
            ))
            continue
        # duplicate RFQ line (another supplier quote): attach prices to a synthetic price-only bump
        # by emitting extra rfq with empty client so it won't inflate client/RFQ counts if we filter —
        # simpler: append prices onto a shadow event ignored for counts in match via kind=rfq_price
        if sell or purchase or market:
            events.append(Event(
                kind="rfq_price",
                pn=pn,
                alt=alt if alt and alt != pn else "",
                client="",
                qty=qty,
                price=sell,
                desc=desc,
                cond=cond,
                purchase=purchase,
                market_price=market,
            ))
    return events


def stock_keys(pn: str, alt: str) -> tuple[set[str], set[str]]:
    exact = {pn}
    for a in alt.split(","):
        a = a.strip()
        if a:
            exact.add(a)
    softs = {soft_pn_key(x) for x in exact if len(soft_pn_key(x)) >= 5}
    # short numeric lamps (387, 6839…) — exact only, no soft
    return exact, softs


def event_matches(e: Event, exact: set[str], softs: set[str]) -> str:
    """Return match mode or empty."""
    if e.pn in exact:
        return "exact"
    if e.alt and e.alt in exact:
        return "alt"
    ep, ea = soft_pn_key(e.pn), soft_pn_key(e.alt) if e.alt else ""
    if softs and ep in softs:
        return "soft"
    if softs and ea and ea in softs:
        return "soft-alt"
    return ""


def aggregate_for_stock(pn: str, alt: str, events: list[Event]) -> tuple[MarketAgg, str]:
    exact, softs = stock_keys(pn, alt)
    m = MarketAgg()
    modes: set[str] = set()
    rfq_seen: set[tuple] = set()
    for e in events:
        mode = event_matches(e, exact, softs)
        if not mode:
            continue
        modes.add(mode)
        if e.kind == "rfq":
            key = (e.pn, e.client)
            if key in rfq_seen:
                # still take prices
                if e.price:
                    m.exp_sell_prices.append(e.price)
                if e.purchase:
                    m.exp_purchase_prices.append(e.purchase)
                if e.market_price:
                    m.exp_market_prices.append(e.market_price)
                continue
            rfq_seen.add(key)
            m.add_event(e)
        elif e.kind == "rfq_price":
            if e.price:
                m.exp_sell_prices.append(e.price)
            if e.purchase:
                m.exp_purchase_prices.append(e.purchase)
            if e.market_price:
                m.exp_market_prices.append(e.market_price)
            if e.desc:
                m.descriptions.add(e.desc)
        else:
            m.add_event(e)
    if not modes:
        return m, "none"
    if "exact" in modes:
        return m, "exact"
    if "alt" in modes:
        return m, "alt"
    return m, "soft"


def recommend_price(m: MarketAgg) -> tuple[Optional[float], str, dict]:
    taz_med = median(m.order_prices)
    tuz_med = median(m.offer_prices)
    exp_sell = median(m.exp_sell_prices)
    exp_purch = median(m.exp_purchase_prices)
    exp_mkt = median(m.exp_market_prices)
    detail = {
        "taz_median": taz_med,
        "taz_n": len(m.order_prices),
        "tuz_median": tuz_med,
        "tuz_n": len(m.offer_prices),
        "exp_sell_median": exp_sell,
        "exp_sell_n": len(m.exp_sell_prices),
        "exp_purchase_median": exp_purch,
        "exp_market_median": exp_mkt,
    }
    # Priority: real order sell > our TUZ offer > Expendables sell quote
    if taz_med is not None and len(m.order_prices) >= 2:
        return round(taz_med, 2), "ТАЗ sell (median, n≥2)", detail
    if taz_med is not None and tuz_med is not None:
        # blend 70/30 when single order
        blended = 0.7 * taz_med + 0.3 * tuz_med
        return round(blended, 2), "blend 70% ТАЗ + 30% ТУЗ", detail
    if taz_med is not None:
        return round(taz_med, 2), "ТАЗ sell (median)", detail
    if tuz_med is not None and exp_sell is not None:
        blended = 0.6 * tuz_med + 0.4 * exp_sell
        return round(blended, 2), "blend 60% ТУЗ + 40% Exp sell", detail
    if tuz_med is not None:
        return round(tuz_med, 2), "ТУЗ offered (median)", detail
    if exp_sell is not None:
        return round(exp_sell, 2), "Expendables Sell EA (median)", detail
    if exp_mkt is not None:
        return round(exp_mkt, 2), "Expendables Market EA (median)", detail
    return None, "нет цены", detail


def write_report(rows_out: list[dict], meta: dict, path: Path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "TOP-10"

    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(bold=True, color="FFFFFF")
    thin = Border(
        left=Side(style="thin", color="B0B0B0"),
        right=Side(style="thin", color="B0B0B0"),
        top=Side(style="thin", color="B0B0B0"),
        bottom=Side(style="thin", color="B0B0B0"),
    )
    money_font = Font(bold=True, color="0B6E4F")

    headers = [
        "#", "P/N", "Alt. PN", "Description", "Qty stock", "Cond", "AC Type",
        "Demand score", "Заказов ТАЗ", "Клиентов (заказы)", "Запросов ТУЗ",
        "Клиентов (запросы)", "RFQ Exp", "Клиентов RFQ",
        "Цена рекоменд. USD", "Источник цены",
        "ТАЗ sell median", "n ТАЗ", "ТУЗ offered median", "n ТУЗ",
        "Exp Sell median", "Exp Purchase median", "Матч",
    ]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(1, i, h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(wrap_text=True, vertical="center")

    for ri, row in enumerate(rows_out, 2):
        vals = [
            row["rank"], row["pn"], row["alt"], row["desc"], row["qty"], row["cond"], row["ac"],
            row["score"], row["n_orders"], row["clients_ord"], row["n_requests"],
            row["clients_req"], row["n_rfq"], row["clients_rfq"],
            row["price"], row["price_src"],
            row["taz_median"], row["taz_n"], row["tuz_median"], row["tuz_n"],
            row["exp_sell"], row["exp_purch"], row["match"],
        ]
        for ci, v in enumerate(vals, 1):
            cell = ws.cell(ri, ci, v)
            cell.border = thin
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if ci == 15 and isinstance(v, (int, float)):
                cell.font = money_font
                cell.number_format = '#,##0.00'

    widths = [4, 18, 18, 28, 10, 10, 18, 12, 10, 12, 10, 12, 8, 10, 14, 28, 12, 8, 12, 8, 12, 12, 8]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 36
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(rows_out)+1}"
    ws.freeze_panes = "B2"

    # Methodology
    ws2 = wb.create_sheet("Методика")
    notes = [
        "Топ-10 складских позиций по рыночному спросу + индикативная цена",
        "",
        f"Склад: {meta['stock_file']} (без Сириуса), позиций: {meta['stock_n']}",
        f"ТАЗ ORDERS: событий заказов: {meta['taz_n']}",
        f"ТУЗ запросы: событий: {meta['tuz_n']}",
        f"Expendables RFQ строк: {meta['exp_n']}",
        f"Позиций склада с ненулевым спросом: {meta['hits']}",
        "",
        "Спрос (score) = заказы×12 + клиенты_заказов×10 + запросы_ТУЗ×3 + клиенты_запросов×5 + RFQ×2 + клиенты_RFQ×3",
        "Исключены: CANCEL/REFUND в ТАЗ; внутренние «закупка на склад»; SAMPLE/TBA.",
        "Матч P/N: exact → soft (без разделителей); Alt. P/N индексируется.",
        "",
        "Цена (приоритет):",
        "1) медиана Продажная ед. из ТАЗ (если ≥2 наблюдений)",
        "2) иначе blend 70% ТАЗ + 30% ТУЗ Offered (если есть оба)",
        "3) иначе медиана ТАЗ / медиана ТУЗ Offered",
        "4) иначе blend 60% ТУЗ + 40% Expendables Sell EA (Routine)",
        "5) иначе Expendables Sell / Market EA",
        "",
        "Это индикативная рыночная цена продажи исходя из наших сделок и предложений, не себестоимость.",
    ]
    for i, line in enumerate(notes, 1):
        ws2.cell(i, 1, line)
        if i == 1:
            ws2.cell(i, 1).font = Font(bold=True, size=13)
    ws2.column_dimensions["A"].width = 110

    # All hits ranked (for transparency)
    ws3 = wb.create_sheet("Все хиты склада")
    hit_headers = ["#", "P/N", "Description", "Qty", "Score", "ТАЗ", "ТУЗ", "RFQ", "Цена", "Источник"]
    for i, h in enumerate(hit_headers, 1):
        c = ws3.cell(1, i, h)
        c.fill = header_fill
        c.font = header_font
    for ri, row in enumerate(meta["all_hits"], 2):
        for ci, v in enumerate([
            row["rank"], row["pn"], row["desc"], row["qty"], row["score"],
            row["n_orders"], row["n_requests"], row["n_rfq"], row["price"], row["price_src"],
        ], 1):
            ws3.cell(ri, ci, v).border = thin
    for i, w in enumerate([4, 18, 28, 8, 10, 8, 8, 8, 12, 28], 1):
        ws3.column_dimensions[get_column_letter(i)].width = w

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    OUT_ART.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT_ART)


def main():
    print("Loading stock...")
    stock = load_stock(STOCK)
    print(f"  stock lines: {len(stock)}")

    print("Loading TAZ ORDERS...")
    taz_events = load_taz_events(TAZ)
    print(f"  TAZ events: {len(taz_events)}")

    print("Loading TUZ...")
    tuz_events = load_tuz_events(TUZ)
    print(f"  TUZ events: {len(tuz_events)}")

    print("Loading Expendables...")
    exp_events = load_exp_events(EXP)
    print(f"  Exp events: {len(exp_events)}")

    events = taz_events + tuz_events + exp_events

    # Aggregate stock by PN (sum qty, keep first desc/cond/ac)
    by_pn: dict[str, StockLine] = {}
    for s in stock:
        if s.pn not in by_pn:
            by_pn[s.pn] = StockLine(s.pn, s.alt, s.desc, s.qty, s.cond, s.ac)
        else:
            by_pn[s.pn].qty += s.qty
            if s.alt and s.alt not in by_pn[s.pn].alt:
                by_pn[s.pn].alt = ",".join(x for x in [by_pn[s.pn].alt, s.alt] if x)

    ranked = []
    for pn, s in by_pn.items():
        m, mode = aggregate_for_stock(pn, s.alt, events)
        if m.demand_score <= 0:
            continue
        price, src, detail = recommend_price(m)
        ranked.append({
            "pn": pn,
            "alt": s.alt,
            "desc": s.desc or (next(iter(m.descriptions), "")),
            "qty": s.qty,
            "cond": s.cond,
            "ac": s.ac,
            "score": m.demand_score,
            "n_orders": m.n_orders,
            "clients_ord": len(m.clients_ord),
            "n_requests": m.n_requests,
            "clients_req": len(m.clients_req),
            "n_rfq": m.n_rfq,
            "clients_rfq": len(m.clients_rfq),
            "price": price,
            "price_src": src,
            "taz_median": detail["taz_median"],
            "taz_n": detail["taz_n"],
            "tuz_median": detail["tuz_median"],
            "tuz_n": detail["tuz_n"],
            "exp_sell": detail["exp_sell_median"],
            "exp_purch": detail["exp_purchase_median"],
            "match": mode,
        })

    ranked.sort(key=lambda r: (-r["score"], -(r["price"] or 0), r["pn"]))

    # Dedupe near-identical stock lines that share soft PN / mutual alts:
    # keep highest-score representative.
    def identity(r):
        keys = {soft_pn_key(r["pn"])}
        for a in (r["alt"] or "").split(","):
            a = a.strip()
            if a:
                keys.add(soft_pn_key(a))
        return frozenset(k for k in keys if k)

    deduped = []
    used: set[str] = set()
    for r in ranked:
        ids = identity(r)
        if ids & used:
            continue
        used |= ids
        deduped.append(r)
    ranked = deduped

    for i, r in enumerate(ranked, 1):
        r["rank"] = i

    top10 = ranked[:10]
    print("\nTOP-10:")
    for r in top10:
        print(
            f"  #{r['rank']:2d} {r['pn']:<22} score={r['score']:.0f} "
            f"TAZ={r['n_orders']}({r['clients_ord']}cl) TUZ={r['n_requests']} "
            f"RFQ={r['n_rfq']} price={r['price']} ({r['price_src']}) [{r['match']}]"
        )

    meta = {
        "stock_file": STOCK.name,
        "stock_n": len(by_pn),
        "taz_n": len(taz_events),
        "tuz_n": len(tuz_events),
        "exp_n": len(exp_events),
        "hits": len(ranked),
        "all_hits": ranked[:200],
    }
    write_report(top10, meta, OUT)
    print(f"\nWrote {OUT}")
    print(f"Wrote {OUT_ART}")


if __name__ == "__main__":
    main()
