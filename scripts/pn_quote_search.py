#!/usr/bin/env python3
"""Search Gmail for supplier quotes by part number(s) and print TUZ O–AA rows.

Usage:
  python scripts/pn_quote_search.py 7036410-1904
  python scripts/pn_quote_search.py 7036410-1904 1590M21P01
  python scripts/pn_quote_search.py --file pns.txt
  echo -e "7036410-1904\\n1590M21P01" | python scripts/pn_quote_search.py --stdin

Credentials: .secrets/gmail.env  (GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
"""

from __future__ import annotations

import argparse
import email
import imaplib
import re
import ssl
import sys
from dataclasses import dataclass, field
from datetime import datetime
from email.header import decode_header
from email.utils import parseaddr, parsedate_to_datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_ENV = ROOT / ".secrets" / "gmail.env"

# TUZ columns O–AA (paste order)
TUZ_O_AA_HEADERS = [
    "Quote obtained date&time",  # O
    "Initials",  # P
    "Root Price per unit, $",  # Q
    "Root",  # R
    "Supplier Price per unit, $",  # S
    "Supplier",  # T
    "Transit per unit, $",  # U
    "Cond",  # V
    "Lead Time (days)",  # W
    "Qty Avail",  # X
    "Cert Info",  # Y
    "PPWK",  # Z
    "Remarks",  # AA
]

INTERNAL_DOMAINS = {
    "kt-mnt.com",
    "partsbase.com",
    "app.rotabull.com",
    "accounts.google.com",
}

INTERNAL_NAME_HINTS = (
    "purchasing@",
    "rfqs@",
    "no-reply@",
    "noreply@",
)

# Longer tokens first so "SV EX" wins over "SV"
COND_RE = re.compile(
    r"\b(SV\s*EX|IN\s*REP|AS\s*REMOVED|OVERHAULED|SERVICEABLE|REPAIRABLE|REPAIRED|REPAIR|"
    r"NEW|NE|OH|SV|RP|AR|FN)\b",
    re.I,
)
PRICE_PATTERNS = [
    re.compile(r"unit\s*price\s*[:\-]?\s*\$?\s*([\d,]+(?:\.\d+)?)\s*([kmb])?\s*(?:usd|\$)?", re.I),
    re.compile(r"(?:price|offer(?:ed)?|cost|up)\s*[:\-]?\s*\$\s*([\d,]+(?:\.\d+)?)\s*([kmb])?", re.I),
    re.compile(r"\$\s*([\d,]+(?:\.\d+)?)\s*([kmb])?\s*(?:each|ea|\/ea|usd|nett|net|/1ea)?", re.I),
    re.compile(r"\b([\d,]+(?:\.\d+)?)\s*([kmb])\s*(?:usd|\$|/1ea|/ea)?\b", re.I),
    re.compile(r"\b([\d,]+(?:\.\d+)?)\s*(?:usd|\$)\b", re.I),
    re.compile(r"(?:for|at)\s+\$\s*([\d,]+(?:\.\d+)?)\s*([kmb])?", re.I),
]
QTY_RE = re.compile(r"(?:qty|quantity|qnty)\s*[:\-]?\s*(\d+(?:\.\d+)?)\b", re.I)
LT_RE = re.compile(
    r"(?:lead\s*time|delivery)\s*[:\-]?\s*([^\n|;]{1,40})",
    re.I,
)
LT_DAYS_BEFORE_RE = re.compile(r"\b(\d+)\s*(?:day|days|d)\s*l\.?t\.?\b", re.I)
LT_SHORT_RE = re.compile(r"\blt\s*[:\-]?\s*([^\n|;]{1,30})", re.I)
STOCK_RE = re.compile(r"\b(in\s*stock|stock|ex\s*stock|ready\s*to\s*ship)\b", re.I)
CERT_RE = re.compile(
    r"\b(8130(?:-3)?|EASA\s*Form[-\s]?1|Dual\s*Release|FAA|CoC|ARC|NIS|OEM\s*tag|trace(?:ability)?)\b",
    re.I,
)
PN_IN_TEXT_RE_TMPL = r"(?<![A-Z0-9]){pn}(?![A-Z0-9])"


@dataclass
class QuoteHit:
    pn: str
    uid: str
    date: datetime | None
    from_raw: str
    supplier: str
    subject: str
    price: float | None = None
    cond: str = ""
    qty: str = ""
    lead_time: str = ""
    cert: str = ""
    remarks: str = ""
    attachments: list[str] = field(default_factory=list)
    snippet: str = ""
    source_note: str = ""

    def tuz_row(self) -> list[str]:
        date_s = self.date.strftime("%Y-%m-%d %H:%M") if self.date else ""
        price_s = f"{self.price:g}" if self.price is not None else ""
        remarks = self.remarks
        if self.source_note:
            remarks = (remarks + " | " if remarks else "") + self.source_note
        return [
            date_s,  # O
            "",  # P Initials — fill manually
            "",  # Q Root Price
            "",  # R Root
            price_s,  # S Supplier Price
            self.supplier,  # T
            "",  # U Transit
            self.cond,  # V
            self.lead_time,  # W
            self.qty,  # X
            self.cert,  # Y
            "",  # Z PPWK
            remarks,  # AA
        ]


def load_env(path: Path) -> dict[str, str]:
    if not path.exists():
        raise SystemExit(
            f"Missing credentials file: {path}\n"
            "Create it with:\n"
            "  GMAIL_ADDRESS=you@domain.com\n"
            "  GMAIL_APP_PASSWORD='xxxx xxxx xxxx xxxx'\n"
        )
    env: dict[str, str] = {}
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, v = line.split("=", 1)
        env[k.strip()] = v.strip().strip("'\"")
    return env


def decode_mime(value: str | None) -> str:
    if not value:
        return ""
    parts = []
    for chunk, charset in decode_header(value):
        if isinstance(chunk, bytes):
            parts.append(chunk.decode(charset or "utf-8", errors="replace"))
        else:
            parts.append(chunk)
    return "".join(parts)


def normalize_pn(pn: str) -> str:
    return re.sub(r"\s+", "", pn.strip().upper())


def soft_pn(pn: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", normalize_pn(pn))


def parse_pns(args: argparse.Namespace) -> list[str]:
    raw: list[str] = []
    if args.stdin:
        raw.extend(sys.stdin.read().splitlines())
    if args.file:
        raw.extend(Path(args.file).read_text(encoding="utf-8").splitlines())
    raw.extend(args.part_numbers or [])
    pns = []
    seen = set()
    for line in raw:
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        # allow comma/semicolon separated on one line
        for token in re.split(r"[,;\s]+", line):
            token = token.strip()
            if not token:
                continue
            key = normalize_pn(token)
            if key in seen:
                continue
            seen.add(key)
            pns.append(token.strip())
    if not pns:
        raise SystemExit("No part numbers provided.")
    return pns


def extract_body(msg: email.message.Message) -> str:
    texts: list[str] = []
    htmls: list[str] = []
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            disp = str(part.get("Content-Disposition") or "")
            if "attachment" in disp.lower():
                continue
            payload = part.get_payload(decode=True) or b""
            text = payload.decode(part.get_content_charset() or "utf-8", errors="replace")
            if ctype == "text/plain":
                texts.append(text)
            elif ctype == "text/html":
                htmls.append(text)
    else:
        payload = msg.get_payload(decode=True) or b""
        texts.append(payload.decode(msg.get_content_charset() or "utf-8", errors="replace"))
    if texts:
        body = "\n".join(texts)
    elif htmls:
        html = "\n".join(htmls)
        html = re.sub(r"(?is)<(script|style).*?>.*?</\1>", " ", html)
        body = re.sub(r"<br\s*/?>", "\n", html, flags=re.I)
        body = re.sub(r"</p>|</div>|</tr>", "\n", body, flags=re.I)
        body = re.sub(r"<[^>]+>", " ", body)
    else:
        body = ""
    body = body.replace("\xa0", " ").replace("\r", "")
    body = re.sub(r"[ \t]+", " ", body)
    body = re.sub(r"\n{3,}", "\n\n", body)
    return body.strip()


def top_reply(body: str, max_chars: int = 1800) -> str:
    """Keep the newest reply; cut quoted history."""
    cut_markers = [
        r"\nOn .+ wrote:\n",
        r"\nFrom:\s+.+\nSent:\s+",
        r"\n-+Original Message-+",
        r"\n_{5,}\n",
        r"\nPartsBase Quote Request",
        r"\nRequest for Quote\n",
    ]
    cut_at = len(body)
    for pat in cut_markers:
        m = re.search(pat, body, flags=re.I)
        if m:
            cut_at = min(cut_at, m.start())
    return body[: min(cut_at, max_chars)].strip()


def domain_of(addr: str) -> str:
    return addr.split("@")[-1].lower().strip() if "@" in addr else ""


def supplier_from_from(from_raw: str) -> str:
    name, addr = parseaddr(from_raw)
    name = (name or "").strip().strip("'\"")
    addr = (addr or "").strip()
    # "X via purchasing" → prefer display before via
    if " via " in name.lower():
        name = re.split(r"\s+via\s+", name, flags=re.I)[0].strip().strip("'\"")
    name = name.strip("'\"")
    # Prefer company-looking domain local-part over bare "Sales"
    if name.lower() in {"sales", "spares", "info", "offer", "quotes", "quote"} and addr:
        dom = domain_of(addr)
        if dom and dom not in INTERNAL_DOMAINS and dom != "kt-mnt.com":
            company = dom.split(".")[0]
            return company.upper() if len(company) <= 4 else company.title()
    if name and domain_of(addr) not in {"", "kt-mnt.com"}:
        return name
    if name:
        return name
    return addr or from_raw


def is_likely_internal(from_raw: str) -> bool:
    name, addr = parseaddr(from_raw)
    dom = domain_of(addr)
    if dom == "kt-mnt.com":
        return True
    if dom in {"accounts.google.com"}:
        return True
    return False


def parse_money(token: str, suffix: str | None = None) -> float | None:
    try:
        v = float(token.replace(",", ""))
    except ValueError:
        return None
    if suffix:
        mult = {"k": 1_000, "m": 1_000_000, "b": 1_000_000_000}.get(suffix.lower())
        if mult:
            v *= mult
    if v <= 0 or v > 50_000_000:
        return None
    # skip years / RFQ ids mistaken as prices
    if 1900 <= v <= 2100 and "." not in token and not suffix:
        return None
    if v >= 10_000_000:  # RFQ-like
        return None
    return v


def extract_prices(text: str) -> list[float]:
    found: list[float] = []
    for pat in PRICE_PATTERNS:
        for m in pat.finditer(text):
            groups = m.groups()
            token = groups[0]
            suffix = groups[1] if len(groups) > 1 else None
            # patterns without suffix group still work
            if suffix and suffix.lower() not in {"k", "m", "b"}:
                suffix = None
            v = parse_money(token, suffix)
            if v is not None and v not in found:
                # ignore tiny numbers that are usually qty/days unless clearly $-tagged
                if v < 20 and "$" not in m.group(0) and "usd" not in m.group(0).lower():
                    continue
                found.append(v)
    return found


def normalize_cond(raw: str) -> str:
    s = re.sub(r"\s+", " ", raw.strip().upper())
    mapping = {
        "NEW": "NE",
        "OVERHAULED": "OH",
        "SERVICEABLE": "SV",
        "AS REMOVED": "AR",
        "REPAIR": "RP",
        "REPAIRED": "SV",
        "REPAIRABLE": "RP",
        "SV EX": "SV EX",
        "SV E": "SV EX",
        "IN REP": "IN REP",
    }
    return mapping.get(s, s)


def extract_conds(text: str) -> list[str]:
    out = []
    for m in COND_RE.finditer(text):
        c = normalize_cond(m.group(1))
        if c not in out:
            out.append(c)
    return out


def _clean_lt(lt: str) -> str:
    lt = lt.strip(" .;,:")
    lt = re.split(r"\s{2,}|\|", lt)[0].strip()
    if re.search(r"available\s*qty|unless different|comments|price\b", lt, re.I):
        # table header noise — try to pull days if present
        days = re.search(r"(\d+)\s*[-–]?\s*(\d+)?\s*(?:day|days|d|business)", lt, re.I)
        if days:
            return days.group(0).replace("business", "").strip()
        return ""
    days = re.search(r"(\d+)\s*[-–]\s*(\d+)\s*(?:day|days|weeks?|w)\b", lt, re.I)
    if days:
        unit = "weeks" if re.search(r"week", lt, re.I) else ""
        return f"{days.group(1)}-{days.group(2)}" + (f" {unit}" if unit else "")
    days = re.search(r"(\d+)\s*(?:day|days|d)\b", lt, re.I)
    if days:
        return days.group(1)
    weeks = re.search(r"(\d+\s*[-–]\s*\d+|\d+)\s*weeks?\b", lt, re.I)
    if weeks:
        return weeks.group(0)
    if STOCK_RE.search(lt):
        return "STOCK"
    if CERT_RE.search(lt) and not re.search(r"\d", lt):
        return ""
    if len(lt) > 30 and not re.search(r"\d", lt):
        return ""
    return lt[:40]


def extract_lead_time(text: str) -> str:
    m = LT_DAYS_BEFORE_RE.search(text)
    if m:
        return m.group(1)
    m = LT_RE.search(text)
    if m:
        cleaned = _clean_lt(m.group(1))
        if cleaned:
            return cleaned
    m = LT_SHORT_RE.search(text)
    if m:
        cleaned = _clean_lt(m.group(1))
        if cleaned:
            return cleaned
    # "1-2 days" near qty/price tables
    m = re.search(r"\b(\d+\s*[-–]\s*\d+)\s*(?:business\s*)?days?\b", text, re.I)
    if m:
        return m.group(1).replace(" ", "")
    if STOCK_RE.search(text):
        return "STOCK"
    return ""


def extract_certs(text: str) -> str:
    hits = []
    for m in CERT_RE.finditer(text):
        token = re.sub(r"\s+", " ", m.group(0)).strip()
        key = token.upper()
        if key not in {h.upper() for h in hits}:
            hits.append(token)
    return ", ".join(hits[:6])


def extract_qty(text: str) -> str:
    m = QTY_RE.search(text)
    if m:
        return m.group(1)
    # "/1ea" or "1 ea" near price — only accept small quantities
    m = re.search(r"(?<![\d.,])(\d{1,4})\s*(?:ea|each|pcs|units?)\b", text, re.I)
    if m:
        return m.group(1)
    m = re.search(r"/\s*(\d{1,4})\s*ea\b", text, re.I)
    if m:
        return m.group(1)
    return ""


def pn_mentioned(text: str, pn: str) -> bool:
    soft = soft_pn(pn)
    soft_text = soft_pn(text)
    if soft and soft in soft_text:
        return True
    # also allow subject-only searches where body is RFQ boilerplate without PN in top reply
    return False


def message_about_pn(subject: str, body: str, pn: str) -> bool:
    """PN in subject/top reply, or quote-reply whose thread body contains PN."""
    head = top_reply(body, max_chars=2500)
    subj_soft = soft_pn(subject)
    head_soft = soft_pn(head)
    body_soft = soft_pn(body)
    target = soft_pn(pn)
    if not target:
        return False
    if target in subj_soft or target in head_soft:
        return True
    # Supplier reply often omits P/N in the short answer; RFQ is below the quote cut.
    quote_ish = bool(re.search(r"\b(quote|rfq|quotation|offer)\b", subject, re.I))
    if quote_ish and target in body_soft and extract_prices(head):
        return True
    return False


FEE_PRICE_RE = re.compile(
    r"(minimum\s*order|processing\s*\$|hazmat\s*fee|fee\s+is\s+added|additional\s+special\s+crate|"
    r"under\s*\$\s*\d)",
    re.I,
)


def structured_line_offers(text: str, pn: str) -> list[dict]:
    """Parse multi-line offers like 'Condition: SV | QTY: 6 | ... Unit Price: 78000'."""
    offers = []
    for line in text.splitlines():
        if not re.search(r"(condition|unit price|\$\s*\d)", line, re.I):
            continue
        prices = extract_prices(line)
        conds = extract_conds(line)
        # Only keep structured rows that actually carry a price
        if not prices:
            continue
        offers.append(
            {
                "price": prices[0],
                "cond": conds[0] if conds else "",
                "qty": extract_qty(line),
                "lead_time": extract_lead_time(line),
                "cert": extract_certs(line),
                "line": line.strip()[:240],
            }
        )
    return offers


def parse_quotes_from_message(
    *,
    pn: str,
    uid: str,
    msg: email.message.Message,
) -> list[QuoteHit]:
    subject = decode_mime(msg.get("Subject"))
    from_raw = decode_mime(msg.get("From"))
    date_raw = msg.get("Date")
    try:
        dt = parsedate_to_datetime(date_raw) if date_raw else None
        if dt and dt.tzinfo:
            dt = dt.astimezone().replace(tzinfo=None)
    except (TypeError, ValueError, IndexError):
        dt = None

    atts = []
    for part in msg.walk():
        fn = part.get_filename()
        if fn:
            atts.append(decode_mime(fn))

    body = extract_body(msg)
    if not message_about_pn(subject, body, pn):
        return []
    head = top_reply(body)
    blob = f"{subject}\n{head}"

    # Skip obvious non-quotes
    if re.search(r"\bno\s*quote\b|\bnq\b|unable to quote|no stock", blob, re.I):
        if not extract_prices(head):
            return []

    # Internal follow-ups without a price in the latest reply → skip
    if is_likely_internal(from_raw) and not extract_prices(head):
        return []

    supplier = supplier_from_from(from_raw)
    # Body often has "Quotation from COMPANY"
    m_co = re.search(
        r"(?:quotation|quote|offer)\s+from\s+([A-Z][A-Za-z0-9 &.'\-/]{2,60})",
        head,
        re.I,
    )
    if m_co and supplier.lower() in {"sales", "spares", "info", "offer", "quotes", "quote"}:
        supplier = m_co.group(1).strip().rstrip(".")

    hits: list[QuoteHit] = []
    head_conds = extract_conds(head)
    head_cert = extract_certs(head)
    head_lt = extract_lead_time(head)
    head_qty = extract_qty(head)

    line_offers = structured_line_offers(head, pn)
    # Drop fee / minimum-order lines
    line_offers = [
        o
        for o in line_offers
        if not FEE_PRICE_RE.search(o["line"])
        and not (o["price"] is not None and o["price"] < 500 and FEE_PRICE_RE.search(head))
    ]
    if line_offers:
        for off in line_offers:
            hits.append(
                QuoteHit(
                    pn=normalize_pn(pn),
                    uid=uid,
                    date=dt,
                    from_raw=from_raw,
                    supplier=supplier,
                    subject=subject,
                    price=off["price"],
                    cond=off["cond"] or (head_conds[0] if head_conds else ""),
                    qty=off["qty"] or head_qty,
                    lead_time=off["lead_time"] or head_lt,
                    cert=off["cert"] or head_cert,
                    remarks=off["line"],
                    attachments=atts,
                    snippet=head[:300],
                    source_note=f"uid:{uid}",
                )
            )
        return hits

    prices = extract_prices(head)
    # Filter fee-like prices when the surrounding sentence looks like a fee
    filtered_prices: list[float] = []
    for p in prices:
        # find a window around this price in text
        keep = True
        for m in re.finditer(re.escape(f"{p:g}"), head.replace(",", "")):
            pass
        for sent in re.split(r"[\n.]", head):
            if not extract_prices(sent):
                continue
            if p in extract_prices(sent) and FEE_PRICE_RE.search(sent):
                keep = False
                break
        if keep:
            filtered_prices.append(p)
    prices = filtered_prices or prices[:1]

    # If no price in top reply, don't invent from quoted history
    if not prices:
        # attachment-only quote hint
        if any(re.search(r"\.(pdf|xlsx?|csv)$", a, re.I) for a in atts):
            hits.append(
                QuoteHit(
                    pn=normalize_pn(pn),
                    uid=uid,
                    date=dt,
                    from_raw=from_raw,
                    supplier=supplier,
                    subject=subject,
                    remarks=f"See attachment(s): {', '.join(atts)}",
                    attachments=atts,
                    snippet=head[:300],
                    source_note=f"uid:{uid}; attachment-only",
                )
            )
        return hits

    cond = head_conds[0] if head_conds else ""
    # "Repaired condition" → SV-ish hint
    if not cond and re.search(r"\brepaired\b", head, re.I):
        cond = "SV"
    price = prices[0]
    hits.append(
        QuoteHit(
            pn=normalize_pn(pn),
            uid=uid,
            date=dt,
            from_raw=from_raw,
            supplier=supplier,
            subject=subject,
            price=price,
            cond=cond,
            qty=head_qty or "1",
            lead_time=head_lt,
            cert=head_cert,
            remarks=re.sub(r"\s+", " ", head)[:220],
            attachments=atts,
            snippet=head[:300],
            source_note=f"uid:{uid}",
        )
    )
    # Extra distinct unit prices only if they look like real offers
    for extra in prices[1:3]:
        if extra == price or extra < 1000:
            continue
        # ignore secondary price that is tiny vs primary (fees / min order)
        if price and extra < price * 0.05:
            continue
        hits.append(
            QuoteHit(
                pn=normalize_pn(pn),
                uid=uid,
                date=dt,
                from_raw=from_raw,
                supplier=supplier,
                subject=subject,
                price=extra,
                cond=head_conds[1] if len(head_conds) > 1 else cond,
                qty=head_qty,
                lead_time=head_lt,
                cert=head_cert,
                remarks="additional price from same email",
                attachments=atts,
                snippet=head[:300],
                source_note=f"uid:{uid}",
            )
        )
    return hits


def gmail_connect(address: str, app_password: str) -> imaplib.IMAP4_SSL:
    client = imaplib.IMAP4_SSL("imap.gmail.com", 993, ssl_context=ssl.create_default_context())
    client.login(address, app_password.replace(" ", ""))
    return client


def search_uids(client: imaplib.IMAP4_SSL, pn: str, limit: int) -> list[bytes]:
    # Gmail full-text search; quote PN for exact token when possible
    raw = f'"{pn}"'
    typ, data = client.uid("search", None, "X-GM-RAW", raw)
    if typ != "OK":
        return []
    uids = data[0].split() if data and data[0] else []
    if not uids:
        # softer search without quotes
        typ, data = client.uid("search", None, "X-GM-RAW", pn)
        uids = data[0].split() if data and data[0] else []
    return uids[-limit:]


def dedupe_hits(hits: list[QuoteHit]) -> list[QuoteHit]:
    seen = set()
    out = []
    for h in hits:
        key = (
            h.pn,
            h.supplier.lower(),
            h.price,
            h.cond,
            h.qty,
            h.date.date().isoformat() if h.date else "",
        )
        if key in seen:
            continue
        seen.add(key)
        out.append(h)
    out.sort(key=lambda x: (x.date or datetime.min, x.price or 0), reverse=True)
    return out


def print_human(pn: str, hits: list[QuoteHit], mail_hits: int) -> None:
    print()
    print("=" * 72)
    print(f"P/N: {pn}")
    print(f"Писем в Gmail с этим номером: {mail_hits}")
    print(f"Распознанных предложений: {len(hits)}")
    if not hits:
        print("Предложений с ценой/вложением не найдено (или только внутренние письма).")
        return
    for i, h in enumerate(hits, 1):
        when = h.date.strftime("%Y-%m-%d %H:%M") if h.date else "?"
        price = f"${h.price:,.0f}" if h.price is not None else "цена не распознана"
        bits = [
            f"{i}. {h.supplier}",
            f"{price}",
            f"Cond={h.cond or '?'}",
            f"Qty={h.qty or '?'}",
            f"LT={h.lead_time or '?'}",
            f"дата={when}",
        ]
        print(" | ".join(bits))
        if h.cert:
            print(f"   Cert: {h.cert}")
        if h.attachments:
            print(f"   Вложения: {', '.join(h.attachments)}")
        print(f"   Тема: {h.subject[:120]}")
        if h.snippet:
            sn = re.sub(r"\s+", " ", h.snippet)[:180]
            print(f"   Фрагмент: {sn}")


def print_tsv(hits: list[QuoteHit]) -> None:
    print()
    print("-" * 72)
    print("TSV для вставки в ТУЗ (колонки O → AA). Скопируйте блок ниже:")
    print("-" * 72)
    print("\t".join(TUZ_O_AA_HEADERS))
    for h in hits:
        print("\t".join(h.tuz_row()))
    print("-" * 72)


def main() -> None:
    ap = argparse.ArgumentParser(description="Find supplier quotes in Gmail by part number")
    ap.add_argument("part_numbers", nargs="*", help="One or more P/Ns")
    ap.add_argument("--file", "-f", help="Text file with P/Ns (one per line)")
    ap.add_argument("--stdin", action="store_true", help="Read P/Ns from stdin (column)")
    ap.add_argument("--env", default=str(DEFAULT_ENV), help="Path to gmail.env")
    ap.add_argument("--limit", type=int, default=40, help="Max emails to inspect per P/N")
    ap.add_argument("--mailbox", default="INBOX", help="IMAP mailbox (default INBOX)")
    ap.add_argument("--tsv-only", action="store_true", help="Print only TSV rows")
    args = ap.parse_args()

    pns = parse_pns(args)
    env = load_env(Path(args.env))
    address = env.get("GMAIL_ADDRESS", "")
    password = env.get("GMAIL_APP_PASSWORD", "")
    if not address or not password:
        raise SystemExit("GMAIL_ADDRESS / GMAIL_APP_PASSWORD missing in env file")

    client = gmail_connect(address, password)
    try:
        typ, _ = client.select(args.mailbox, readonly=True)
        if typ != "OK":
            raise SystemExit(f"Cannot open mailbox {args.mailbox}")

        all_hits: list[QuoteHit] = []
        for pn in pns:
            uids = search_uids(client, pn, args.limit)
            hits: list[QuoteHit] = []
            for uid in uids:
                typ, msg_data = client.uid("fetch", uid, "(BODY.PEEK[])")
                if typ != "OK" or not msg_data or not msg_data[0]:
                    continue
                raw = msg_data[0][1]
                if not isinstance(raw, (bytes, bytearray)):
                    continue
                msg = email.message_from_bytes(raw)
                body = extract_body(msg)
                subject = decode_mime(msg.get("Subject"))
                if not (
                    pn_mentioned(f"{subject}\n{body}", pn)
                    or soft_pn(pn) in soft_pn(subject)
                    or soft_pn(pn) in soft_pn(body)
                ):
                    continue
                hits.extend(
                    parse_quotes_from_message(pn=pn, uid=uid.decode(), msg=msg)
                )
            hits = dedupe_hits(hits)
            all_hits.extend(hits)
            if not args.tsv_only:
                print_human(pn, hits, mail_hits=len(uids))
        if args.tsv_only:
            print("\t".join(TUZ_O_AA_HEADERS))
            for h in all_hits:
                print("\t".join(h.tuz_row()))
        else:
            print_tsv(all_hits)
    finally:
        try:
            client.logout()
        except Exception:
            pass


if __name__ == "__main__":
    main()
