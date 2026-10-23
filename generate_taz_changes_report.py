#!/usr/bin/env python3
"""Weekly HTML report: TAZ snapshot diff — TROUBLE, cancellations, warranty."""

from __future__ import annotations

import argparse
import html
import re
import zipfile
from collections import defaultdict
from dataclasses import dataclass, field, replace
from datetime import date, datetime, timedelta
from pathlib import Path
from statistics import mean
from typing import Any

import pandas as pd

# --- TAZ columns (shared with other generators) ---
COL_INVOICE = "Номер счета"
COL_STATUS = "Status"
COL_CUSTOMER = "Customer"
COL_PN = "p/n"
COL_DESC = "DESCRIPTION"
COL_SUPPLIER = "Поставщик"
COL_ROOT_SUPPLIER = "Root supplier"
COL_UNIQUE_UNIT = "UNIQUE UNIT CODE"
COL_CATEGORY = "Category"
COL_QTY = "QTY IN PO"
COL_SALE = "Продажная, итого"
COL_PURCHASE = "Закупка, итого"
COL_TRANSPORT_PLAN = (
    "Стоимость доставки ПЛАН, за весь счет! Если в счете несколько строк, "
    'то "размазываем" равномерно планируюмую стоиомость транспорта на все позиции из счета.'
)
COL_FEE = "Transaction fee"
COL_DEADLINE = "КРАЙНЯЯ ДАТА ПОСТАВКИ"
COL_DAYS_TO_DELIVER = "Дней на поставку (ЧИСЛО)"
COL_ORDER_DATE = "ЗАКАЗ ВЗЯТ В РАБОТУ (ДАТА) ОТ КЛИЕНТА"
COL_DELIVERY = "ФАКТИЧЕСКАЯ ДАТА ПОСТАВКИ (СОГЛАСНО УСЛОВИЯМ ПОСТАВКИ)"
COL_COMMENT = "Комментарии"

STATUS_TROUBLE = "6 TROUBLE"
STATUS_CANCEL = "5 CANCELLED"
STATUS_REFUND = "7 REFUND"
STATUS_WARRANTY = "8 WARRANTY"
STATUS_FINISHED = "4 FINISHED"
STATUS_SHIPPED = "3 SHIPPED"

# Statuses that can “become cancelled” during the week (user rule)
CANCEL_SOURCE = {"1 NOT PAID", "2 PAID", STATUS_TROUBLE, "0 NOT PAID"}
TERMINAL_BAD = {STATUS_CANCEL, STATUS_REFUND}
RESOLVED_FROM_TROUBLE = {"1 NOT PAID", "2 PAID", STATUS_SHIPPED, STATUS_FINISHED}


def parse_numeric(value: Any) -> float:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return 0.0
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    text = str(value).strip().replace("\xa0", "").replace(" ", "").replace(",", ".")
    if not text or text.startswith("#"):
        return 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0


def parse_date(value: Any) -> date | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    if not text or text.lower() in {"nan", "none", "nat"}:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(text[:19], fmt).date()
        except ValueError:
            continue
    try:
        ts = pd.to_datetime(value, dayfirst=True, errors="coerce")
        if pd.isna(ts):
            return None
        return ts.date()
    except Exception:
        return None


def parse_report_date(text: str) -> date:
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(text.strip(), fmt).date()
        except ValueError:
            continue
    raise ValueError(f"Cannot parse date: {text!r}")


def normalize_invoice(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip()
    if text.endswith(".0") and text[:-2].isdigit():
        return text[:-2]
    return text


def load_taz(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=0)
    df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]
    rename: dict[str, str] = {}
    for c in df.columns:
        if isinstance(c, str) and c.startswith("Стоимость доставки ПЛАН"):
            rename[c] = COL_TRANSPORT_PLAN
        if isinstance(c, str) and "UNIQUE" in c.upper() and "UNIT" in c.upper():
            rename[c] = COL_UNIQUE_UNIT
    if COL_INVOICE not in df.columns:
        for alt in ("счет", "№ счета"):
            if alt in df.columns:
                rename[alt] = COL_INVOICE
                break
    if rename:
        df = df.rename(columns=rename)
    if COL_INVOICE not in df.columns:
        raise KeyError(f"Invoice column not found. Columns: {list(df.columns)[:10]}")
    return df


def add_row_key(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["_key"] = (
        out[COL_INVOICE].map(normalize_invoice)
        + "|"
        + out[COL_PN].astype(str).str.strip()
        + "|"
        + out[COL_DESC].astype(str).str.strip().str.slice(0, 60)
    )
    return out


COMMERCIAL_EPS = 1.0  # USD — ниже считаем «без изменений»


def row_margin_plan(row: pd.Series) -> float:
    sale = parse_numeric(row.get(COL_SALE))
    purchase = parse_numeric(row.get(COL_PURCHASE))
    transport = parse_numeric(row.get(COL_TRANSPORT_PLAN))
    fee = parse_numeric(row.get(COL_FEE))
    return sale - purchase - transport - fee


def fmt_money(v: float | None, digits: int = 0) -> str:
    if v is None:
        return "—"
    return f"{v:,.{digits}f}".replace(",", " ")


def fmt_pct(v: float | None) -> str:
    if v is None:
        return "—"
    return f"{v:.1f}%"


def fmt_days_range(min_d: int | None, max_d: int | None) -> str:
    if min_d is None and max_d is None:
        return "—"
    if min_d is not None and max_d is not None:
        if min_d == max_d:
            return f"{min_d} дн."
        return f"{min_d}–{max_d} дн."
    if min_d is not None:
        return f"≥{min_d} дн."
    return f"≤{max_d} дн."


def fmt_hung_days(days: int | None, *, unknown_start: bool) -> str:
    """Approximate days a resolved row spent in TROUBLE."""
    if days is None:
        return "—"
    if unknown_start:
        return f"≥{days} дн."
    return f"примерно {days} дн."


def esc(text: Any) -> str:
    return html.escape(str(text or ""), quote=True)


def index_rows(df: pd.DataFrame) -> dict[str, pd.Series]:
    keyed = add_row_key(df).drop_duplicates(subset="_key", keep="first")
    keyed = keyed.set_index("_key", drop=False)
    return {str(k): keyed.loc[k] for k in keyed.index}


@dataclass
class StatusChange:
    key: str
    invoice: str
    customer: str
    pn: str
    description: str
    category: str
    supplier: str
    root_supplier: str
    prev_status: str
    curr_status: str
    change_kind: str
    sale_prev: float
    sale_curr: float
    margin_prev: float
    margin_curr: float
    qty_prev: float
    qty_curr: float
    deadline_prev: date | None
    deadline_curr: date | None
    days_to_deliver_prev: float | None
    days_to_deliver_curr: float | None
    trouble_min_days: int | None = None
    trouble_max_days: int | None = None
    unique_unit_code: str = ""
    notes: list[str] = field(default_factory=list)
    problem_notes: list[str] = field(default_factory=list)

    @property
    def margin_delta(self) -> float:
        return self.margin_curr - self.margin_prev

    @property
    def sale_delta(self) -> float:
        return self.sale_curr - self.sale_prev

    @property
    def deadline_delta_days(self) -> int | None:
        if self.deadline_prev and self.deadline_curr:
            return (self.deadline_curr - self.deadline_prev).days
        return None

    @property
    def supplier_display(self) -> str:
        return format_supplier_display(self.supplier, self.root_supplier, self.unique_unit_code)


def normalize_unique_unit(value: Any) -> str:
    text = str(value or "").strip()
    if not text or text.lower() in {"nan", "none", "nat", "-"}:
        return ""
    return text


def split_unit_code(value: Any) -> tuple[str, str]:
    """Split BOS/LH unique unit into reference (#/№) and trailing problem text."""
    text = normalize_unique_unit(value)
    if not text:
        return "", ""
    match = re.match(
        r"^(?P<ref>(?:#|№)\s*[\w./-]+)\s+(?:[-–—:]\s*)?(?P<extra>[A-Za-zА-Яа-яЁё].+)$",
        text,
    )
    if match:
        ref = re.sub(r"\s+", "", match.group("ref"))
        extra = match.group("extra").strip()
        return ref, extra
    return re.sub(r"\s+", "", text), ""


def format_supplier_display(supplier: str, root_supplier: str, unique_unit: str = "") -> str:
    """JET/IBERIA → root supplier; Lufthansa/Blue Ocean/Avitrue → supplier (+ #/№ reference)."""
    sup_u = (supplier or "").upper()
    root = (root_supplier or "").strip()
    if "JET" in sup_u or "IBERIA" in sup_u:
        return root or supplier or "—"
    if any(x in sup_u for x in ("LUFTHANSA", "BLUE OCEAN", "AVITRUE")):
        name = supplier or "—"
        unit_ref, _ = split_unit_code(unique_unit)
        return f"{name} · {unit_ref}" if unit_ref else name
    return root or supplier or "—"


def row_str(row: pd.Series, col: str) -> str:
    value = row.get(col)
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip()


_NOISE_TEXT = {
    "",
    "-",
    "—",
    "nan",
    "none",
    "nat",
    "n/a",
    "na",
    "sample",
    "#n/a",
    "пусто",
    "нет",
    "ok",
    "ок",
}

_PROBLEM_RE = re.compile(
    r"(?i)(?:"
    r"\bcancell?ed?\b|\bdelay(?:ed|s)?\b|\baog\b|\bnff\b|\breject(?:ed|ion)?\b|"
    r"\bshortage\b|\bbackorder\b|\bawait(?:ing)?\b|\bfail(?:ed|ure)?\b|"
    r"\bmissing\b|\bwrong\b|\bdamag(?:e|ed)\b|\bscrap\b|\bber\b|"
    r"\bobsolete\b|\bunservice(?:able)?\b|\bno stock\b|\bnot available\b|"
    r"\bon hold\b|\bunable\b|\bcannot\b|\bcan't\b|\bproblem\b|\bissue\b|"
    r"\bdiscontinu(?:ed|e)\b|\bhold\b|\berror\b|\btrouble\b|\bquote\b|"
    r"отмен|срыв|сорван|брак|проблем|задержк|не найден|нет в наличии|"
    r"отказ|ожидан|слом|поломк|рекламац|не постав|ждём|ждем|"
    r"не можем|ошибк|нелетн|гарант|нет поставщик"
    r")"
)

_SKIP_PROBLEM_COLS = {
    COL_INVOICE,
    COL_STATUS,
    COL_CUSTOMER,
    COL_PN,
    COL_DESC,
    COL_SUPPLIER,
    COL_ROOT_SUPPLIER,
    COL_UNIQUE_UNIT,
    COL_CATEGORY,
    COL_QTY,
    COL_SALE,
    COL_PURCHASE,
    COL_TRANSPORT_PLAN,
    COL_FEE,
    COL_DEADLINE,
    COL_DAYS_TO_DELIVER,
    COL_ORDER_DATE,
    COL_DELIVERY,
    "_key",
}

_SKIP_PROBLEM_COL_RE = re.compile(
    r"(?i)дата|usd|стоим|оплач|закуп|продаж|qty|price|fee|баланс|"
    r"invoice|счет|customer|status|p/n|descrip|supplier|category|"
    r"invoicer|sgmt|uom|lead time|qty"
)
_COMMENT_COL_RE = re.compile(r"(?i)коммент|примечан|problem|причин")


def _clean_text_cell(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, (datetime, date)):
        return ""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return ""
    text = str(value).strip()
    if text.lower() in _NOISE_TEXT:
        return ""
    return " ".join(text.split())


def looks_like_problem(text: str) -> bool:
    cleaned = _clean_text_cell(text)
    if len(cleaned) < 3:
        return False
    return bool(_PROBLEM_RE.search(cleaned))


def _is_code_like(text: str) -> bool:
    compact = re.sub(r"[\s._-]+", "", text)
    return bool(re.fullmatch(r"(?i)(?:po|p|#|№)?[A-Z]?\d{3,}", compact))


def _is_problem_or_prose(text: str) -> bool:
    cleaned = _clean_text_cell(text)
    if not cleaned:
        return False
    if looks_like_problem(cleaned):
        return True
    if _is_code_like(cleaned):
        return False
    words = re.findall(r"[A-Za-zА-Яа-яЁё]{3,}", cleaned)
    return len(words) >= 3


def _add_problem_note(notes: list[str], seen: set[str], text: str) -> None:
    cleaned = _clean_text_cell(text)
    if not cleaned:
        return
    key = cleaned.lower()
    if key in seen:
        return
    seen.add(key)
    notes.append(cleaned)


def extract_problem_notes(prev: pd.Series, curr: pd.Series) -> list[str]:
    """Pull English/Russian trouble descriptions from TAZ row text fields."""
    notes: list[str] = []
    seen: set[str] = set()

    for row in (curr, prev):
        for col, value in row.items():
            col_name = str(col)
            if col_name == COL_COMMENT or _COMMENT_COL_RE.search(col_name):
                comment = _clean_text_cell(value)
                if comment and _is_problem_or_prose(comment):
                    _add_problem_note(notes, seen, comment)

    for row in (curr, prev):
        _, extra = split_unit_code(row.get(COL_UNIQUE_UNIT) if COL_UNIQUE_UNIT in row.index else None)
        if extra:
            _add_problem_note(notes, seen, extra)

    for row in (curr, prev):
        for col, value in row.items():
            col_name = str(col)
            if col_name in _SKIP_PROBLEM_COLS or col_name == COL_COMMENT:
                continue
            if _SKIP_PROBLEM_COL_RE.search(col_name):
                continue
            text = _clean_text_cell(value)
            if text and looks_like_problem(text):
                _add_problem_note(notes, seen, text)

    return notes


def build_status_change(
    key: str,
    prev: pd.Series,
    curr: pd.Series,
    *,
    change_kind: str,
    prev_status: str,
    curr_status: str,
    notes: list[str],
    trouble_min_days: int | None = None,
    trouble_max_days: int | None = None,
) -> StatusChange:
    sale_prev = parse_numeric(prev.get(COL_SALE))
    sale_curr = parse_numeric(curr.get(COL_SALE))
    qty_prev = parse_numeric(prev.get(COL_QTY))
    qty_curr = parse_numeric(curr.get(COL_QTY))
    curr_unit_ref, _ = split_unit_code(curr.get(COL_UNIQUE_UNIT) if COL_UNIQUE_UNIT in curr.index else None)
    prev_unit_ref, _ = split_unit_code(prev.get(COL_UNIQUE_UNIT) if COL_UNIQUE_UNIT in prev.index else None)
    return StatusChange(
        key=key,
        invoice=normalize_invoice(prev.get(COL_INVOICE) or curr.get(COL_INVOICE)),
        customer=row_str(curr, COL_CUSTOMER) or row_str(prev, COL_CUSTOMER),
        pn=row_str(prev, COL_PN) or row_str(curr, COL_PN),
        description=row_str(prev, COL_DESC) or row_str(curr, COL_DESC),
        category=row_str(prev, COL_CATEGORY) or row_str(curr, COL_CATEGORY),
        supplier=row_str(curr, COL_SUPPLIER) or row_str(prev, COL_SUPPLIER),
        root_supplier=row_str(curr, COL_ROOT_SUPPLIER) or row_str(prev, COL_ROOT_SUPPLIER),
        unique_unit_code=curr_unit_ref or prev_unit_ref,
        prev_status=prev_status,
        curr_status=curr_status,
        change_kind=change_kind,
        sale_prev=sale_prev,
        sale_curr=sale_curr,
        margin_prev=row_margin_plan(prev),
        margin_curr=row_margin_plan(curr),
        qty_prev=qty_prev,
        qty_curr=qty_curr,
        deadline_prev=parse_date(prev.get(COL_DEADLINE)),
        deadline_curr=parse_date(curr.get(COL_DEADLINE)),
        days_to_deliver_prev=parse_numeric(prev.get(COL_DAYS_TO_DELIVER)) or None,
        days_to_deliver_curr=parse_numeric(curr.get(COL_DAYS_TO_DELIVER)) or None,
        trouble_min_days=trouble_min_days,
        trouble_max_days=trouble_max_days,
        notes=notes,
        problem_notes=extract_problem_notes(prev, curr),
    )


def build_notes(category: str, qty_prev: float, qty_curr: float) -> list[str]:
    notes: list[str] = []
    cat = (category or "").upper()
    if qty_prev and qty_curr and abs(qty_prev - qty_curr) > 0.009:
        notes.append(f"кол-во {qty_prev:g} → {qty_curr:g}")
        if "EXP" in cat:
            notes.append("EXP: вероятно правка количества / ед. изм.")
        elif "ROT" in cat:
            notes.append("ROTABLE: проверьте замену юнита")
    return notes


def build_commercial_notes(
    prev: pd.Series,
    curr: pd.Series,
    qty_prev: float = 0,
    qty_curr: float = 0,
) -> list[str]:
    """Детализация изменений по коммерции между двумя снимками."""
    notes: list[str] = []

    def money_delta(label: str, prev_v: float, curr_v: float, up: str, down: str) -> str | None:
        d = curr_v - prev_v
        if abs(d) < COMMERCIAL_EPS:
            return None
        return f"{label} {up if d > 0 else down} на {fmt_money(abs(d), 0)} USD"

    sale_p = parse_numeric(prev.get(COL_SALE))
    sale_c = parse_numeric(curr.get(COL_SALE))
    purch_p = parse_numeric(prev.get(COL_PURCHASE))
    purch_c = parse_numeric(curr.get(COL_PURCHASE))
    trans_p = parse_numeric(prev.get(COL_TRANSPORT_PLAN))
    trans_c = parse_numeric(curr.get(COL_TRANSPORT_PLAN))
    fee_p = parse_numeric(prev.get(COL_FEE))
    fee_c = parse_numeric(curr.get(COL_FEE))

    for part in (
        money_delta("закупка", purch_p, purch_c, "увеличилась", "уменьшилась"),
        money_delta("транспорт", trans_p, trans_c, "увеличился", "уменьшился"),
        money_delta("продажа", sale_p, sale_c, "увеличилась", "уменьшилась"),
    ):
        if part:
            notes.append(part)

    sup_p = str(prev.get(COL_SUPPLIER) or "").strip()
    sup_c = str(curr.get(COL_SUPPLIER) or "").strip()
    if sup_p.lower() != sup_c.lower() and (sup_p or sup_c):
        notes.append(f"поставщик: {sup_p or '—'} → {sup_c or '—'}")

    root_p = str(prev.get(COL_ROOT_SUPPLIER) or "").strip()
    root_c = str(curr.get(COL_ROOT_SUPPLIER) or "").strip()
    if root_p.lower() != root_c.lower() and (root_p or root_c):
        notes.append(f"root supplier: {root_p or '—'} → {root_c or '—'}")

    qty_changed = qty_prev and qty_curr and abs(qty_prev - qty_curr) > 0.009
    fee_note = money_delta("fee", fee_p, fee_c, "увеличился", "уменьшился")
    if fee_note and (notes or qty_changed):
        notes.append(fee_note)

    return notes


def has_commercial_in_notes(notes: list[str]) -> bool:
    text = " ".join(notes)
    return any(kw in text for kw in ("закупка", "транспорт", "продажа", "fee", "поставщик", "root supplier"))


def append_trouble_notes(qty_notes: list[str], commercial_notes: list[str], extra: list[str]) -> list[str]:
    out = [*qty_notes]
    if commercial_notes:
        out.extend(commercial_notes)
    elif not qty_notes:
        out.append("коммерция без изменений")
    out.extend(extra)
    return out


def meaningful_margin_delta(e: StatusChange) -> float | None:
    if e.change_kind in {"entered", "ongoing"}:
        return None
    txt, _ = fmt_margin_delta_cell(e)
    return None if txt in {"—", "0"} else e.margin_delta


def sum_meaningful_margin(items: list[StatusChange]) -> float:
    return sum(d for e in items if (d := meaningful_margin_delta(e)) is not None)


def fmt_sale_cell(e: StatusChange) -> str:
    if abs(e.sale_prev - e.sale_curr) < COMMERCIAL_EPS:
        return fmt_money(e.sale_curr, 0)
    return f"{fmt_money(e.sale_prev, 0)} → {fmt_money(e.sale_curr, 0)}"


def fmt_margin_delta_cell(e: StatusChange) -> tuple[str, str]:
    """Δ маржа план = (продажа − закупка − транспорт − fee)_curr − _prev.

    For new/unresolved TROUBLE the new supplier is not yet found, so Δ is always 0.
    """
    if e.change_kind in {"entered", "ongoing"}:
        return "0", ""
    d = e.margin_delta
    if not has_commercial_in_notes(e.notes):
        return "—", ""
    css = "pos" if d > COMMERCIAL_EPS else ("neg" if d < -COMMERCIAL_EPS else "")
    return fmt_money(d, 0), css


def compare_snapshots(
    prev_rows: dict[str, pd.Series],
    curr_rows: dict[str, pd.Series],
    period_days: int,
    history: list[tuple[date, dict[str, pd.Series]]] | None = None,
) -> tuple[list[StatusChange], list[StatusChange], list[StatusChange], list[StatusChange]]:
    """Return trouble_events, cancellations, refunds, warranty_transitions."""
    trouble: list[StatusChange] = []
    cancellations: list[StatusChange] = []
    refunds: list[StatusChange] = []
    warranty: list[StatusChange] = []

    common = set(prev_rows) & set(curr_rows)
    for key in common:
        prev = prev_rows[key]
        curr = curr_rows[key]
        ps = str(prev.get(COL_STATUS) or "").strip()
        cs = str(curr.get(COL_STATUS) or "").strip()

        if ps == cs:
            if ps == STATUS_TROUBLE:
                qty_prev = parse_numeric(prev.get(COL_QTY))
                qty_curr = parse_numeric(curr.get(COL_QTY))
                category = row_str(prev, COL_CATEGORY) or row_str(curr, COL_CATEGORY)
                qty_notes = build_notes(category, qty_prev, qty_curr)
                commercial_notes = build_commercial_notes(prev, curr, qty_prev, qty_curr)
                trouble.append(
                    build_status_change(
                        key,
                        prev,
                        curr,
                        change_kind="ongoing",
                        prev_status=ps,
                        curr_status=cs,
                        trouble_min_days=period_days,
                        trouble_max_days=None,
                        notes=append_trouble_notes(
                            qty_notes,
                            commercial_notes,
                            [f"в TROUBLE ≥{period_days} дн. без решения (оба снимка)"],
                        ),
                    )
                )
            continue

        sale_prev = parse_numeric(prev.get(COL_SALE))
        sale_curr = parse_numeric(curr.get(COL_SALE))
        qty_prev = parse_numeric(prev.get(COL_QTY))
        qty_curr = parse_numeric(curr.get(COL_QTY))
        category = row_str(prev, COL_CATEGORY) or row_str(curr, COL_CATEGORY)
        qty_notes = build_notes(category, qty_prev, qty_curr)
        commercial_notes = build_commercial_notes(prev, curr, qty_prev, qty_curr)

        base = build_status_change(
            key,
            prev,
            curr,
            change_kind="",
            prev_status=ps,
            curr_status=cs,
            notes=qty_notes,
        )

        # --- TROUBLE lifecycle ---
        if ps != STATUS_TROUBLE and cs == STATUS_TROUBLE:
            trouble.append(
                replace(
                    base,
                    change_kind="entered",
                    trouble_min_days=0,
                    trouble_max_days=period_days,
                    notes=append_trouble_notes(qty_notes, commercial_notes, ["новый TROUBLE за период"]),
                )
            )
        elif ps == STATUS_TROUBLE and cs in RESOLVED_FROM_TROUBLE:
            ev_notes = append_trouble_notes(qty_notes, commercial_notes, [])
            dd = base.deadline_delta_days
            if dd is not None and dd != 0:
                ev_notes.append(f"срок поставки сдвинут на {dd:+d} дн.")
            dtd_prev = base.days_to_deliver_prev
            dtd_curr = base.days_to_deliver_curr
            if dtd_prev is not None and dtd_curr is not None and abs(dtd_prev - dtd_curr) > 0.1:
                ev_notes.append(f"«дней на поставку» {dtd_prev:g} → {dtd_curr:g}")
            trouble.append(
                replace(
                    base,
                    change_kind="resolved",
                    trouble_min_days=0,
                    trouble_max_days=period_days,
                    notes=ev_notes,
                )
            )
        elif ps == STATUS_TROUBLE and cs in TERMINAL_BAD:
            trouble.append(
                replace(
                    base,
                    change_kind="cancelled_from_trouble",
                    trouble_min_days=0,
                    trouble_max_days=period_days,
                    notes=[*qty_notes, *commercial_notes, "TROUBLE → отмена/возврат"],
                )
            )
        elif ps == STATUS_TROUBLE and cs == STATUS_WARRANTY:
            trouble.append(
                replace(
                    base,
                    change_kind="warranty_from_trouble",
                    trouble_min_days=0,
                    trouble_max_days=period_days,
                    notes=[*qty_notes, *commercial_notes, "TROUBLE → гарантия"],
                )
            )

        # --- Cancellations / refunds (indirect date) ---
        if ps in CANCEL_SOURCE and cs == STATUS_CANCEL:
            cancellations.append(
                replace(base, change_kind="cancelled", notes=[*qty_notes, *commercial_notes])
            )
        elif ps in CANCEL_SOURCE and cs == STATUS_REFUND:
            refunds.append(
                replace(base, change_kind="refunded", notes=[*qty_notes, *commercial_notes])
            )

        # --- Warranty (active pipeline → warranty, same rule as cancellations) ---
        if ps in CANCEL_SOURCE and cs == STATUS_WARRANTY:
            warranty.append(
                replace(base, change_kind="warranty", notes=[*qty_notes, *commercial_notes])
            )

    # Строки, которых не было в прошлом снимке, но уже TROUBLE в текущем
    # (заказ взят в работу в период и почти сразу попал в TROUBLE)
    for key in set(curr_rows) - set(prev_rows):
        curr = curr_rows[key]
        cs = str(curr.get(COL_STATUS) or "").strip()
        if cs != STATUS_TROUBLE:
            continue
        order_date = parse_date(curr.get(COL_ORDER_DATE))
        notes = ["строки не было в ТАЗ на начало периода", "новый TROUBLE за период"]
        if order_date:
            notes.insert(1, f"заказ взят в работу {order_date.strftime('%d.%m.%Y')}")
        trouble.append(
            build_status_change(
                key,
                curr,
                curr,
                change_kind="entered",
                prev_status="(нет в прошлом снимке)",
                curr_status=cs,
                trouble_min_days=0,
                trouble_max_days=period_days,
                notes=notes,
            )
        )

    if history:
        _apply_history_duration(trouble, history)

    return trouble, cancellations, refunds, warranty


def _status_of(rows: dict[str, pd.Series], key: str) -> str | None:
    if key not in rows:
        return None
    return str(rows[key].get(COL_STATUS) or "").strip() or None


def duration_in_trouble(
    key: str,
    history: list[tuple[date, dict[str, pd.Series]]],
) -> tuple[int, int | None, date | None, date | None, date | None]:
    """Continuous TROUBLE streak ending at the latest TROUBLE snapshot.

    Returns (hung_days, max_days, first_trouble_date, last_trouble_date, closed_by).

    For a closed streak hung_days is how long the row was seen in TROUBLE:
    last TROUBLE snapshot minus first TROUBLE snapshot. If it was TROUBLE in
    only one file and gone in the next, hung_days is until that next file.

    max_days is None when the streak already exists in the oldest snapshot
    (start may be earlier). For an open streak with a known start, max_days is
    the upper bound using the snapshot before the first TROUBLE.
    """
    empty = (0, None, None, None, None)
    if not history:
        return empty

    last_t_idx: int | None = None
    for i in range(len(history) - 1, -1, -1):
        if _status_of(history[i][1], key) == STATUS_TROUBLE:
            last_t_idx = i
            break
    if last_t_idx is None:
        return empty

    first_t_idx = last_t_idx
    for i in range(last_t_idx - 1, -1, -1):
        if _status_of(history[i][1], key) == STATUS_TROUBLE:
            first_t_idx = i
        else:
            break

    first_date = history[first_t_idx][0]
    last_t_date = history[last_t_idx][0]
    still_open = last_t_idx == len(history) - 1
    closed_by = None if still_open else history[last_t_idx + 1][0]
    started_at_oldest = first_t_idx == 0

    if still_open:
        min_d = max(0, (last_t_date - first_date).days)
        if started_at_oldest:
            return min_d, None, first_date, last_t_date, None
        before_date = history[first_t_idx - 1][0]
        return min_d, max(min_d, (last_t_date - before_date).days), first_date, last_t_date, None

    hung = max(0, (last_t_date - first_date).days)
    if hung == 0 and closed_by is not None:
        hung = (closed_by - first_date).days
    if started_at_oldest:
        return hung, None, first_date, last_t_date, closed_by
    return hung, hung, first_date, last_t_date, closed_by


def _closed_trouble_span(
    first_date: date | None,
    last_t_date: date | None,
    *,
    unknown_start: bool,
) -> str:
    if not first_date:
        return ""
    start = f"с {first_date.strftime('%d.%m')}"
    if unknown_start:
        start += " (или раньше)"
    if last_t_date and last_t_date != first_date:
        return f"{start} по {last_t_date.strftime('%d.%m')}"
    return start


def _hung_comment(
    hung_days: int,
    max_days: int | None,
    first_date: date | None,
    last_t_date: date | None,
    closed_by: date | None,
    *,
    resolved: bool = False,
) -> str:
    unknown_start = max_days is None
    days_txt = fmt_hung_days(hung_days, unknown_start=unknown_start)
    lead = f"решили за {days_txt}" if resolved else f"висела {days_txt}"
    span = _closed_trouble_span(first_date, last_t_date, unknown_start=unknown_start)
    parts = [lead]
    if span:
        parts.append(span)
    if closed_by:
        parts.append(f"в отчёте {closed_by.strftime('%d.%m')} уже не TROUBLE")
    if len(parts) == 1:
        return lead
    return f"{parts[0]} ({', '.join(parts[1:])})"


def _apply_history_duration(
    events: list[StatusChange],
    history: list[tuple[date, dict[str, pd.Series]]],
) -> None:
    closed_kinds = {"resolved", "cancelled_from_trouble", "warranty_from_trouble"}
    for e in events:
        if e.change_kind not in {"ongoing"} | closed_kinds:
            continue
        min_d, max_d, first_date, last_t_date, closed_by = duration_in_trouble(e.key, history)
        e.trouble_min_days = min_d
        e.trouble_max_days = max_d
        e.notes = [
            n
            for n in e.notes
            if not n.startswith("в TROUBLE")
            and not n.startswith("висела")
            and not n.startswith("решили за")
        ]
        if e.change_kind in closed_kinds:
            e.notes.append(
                _hung_comment(
                    min_d,
                    max_d,
                    first_date,
                    last_t_date,
                    closed_by,
                    resolved=e.change_kind == "resolved",
                )
            )
            continue
        stamp = f"с {first_date.strftime('%d.%m')}" if first_date else ""
        if stamp:
            e.notes.append(f"в TROUBLE {fmt_days_range(min_d, max_d)} ({stamp})")
        else:
            e.notes.append(f"в TROUBLE {fmt_days_range(min_d, max_d)}")


def _series_for(
    snapshots: list[tuple[date, dict[str, pd.Series]]],
    key: str,
    *,
    last: bool,
) -> pd.Series | None:
    seq = snapshots if not last else list(reversed(snapshots))
    for _, rows in seq:
        if key in rows:
            return rows[key]
    return None


def compare_full_period(
    snapshots: list[tuple[date, dict[str, pd.Series]]],
) -> tuple[list[StatusChange], list[StatusChange], list[StatusChange], list[StatusChange]]:
    """Classify TROUBLE/cancel across the whole snapshot chain (oldest → newest)."""
    if len(snapshots) < 2:
        raise ValueError("Need at least two snapshots")
    oldest = snapshots[0][1]
    newest = snapshots[-1][1]

    keys: set[str] = set()
    for _, rows in snapshots:
        keys |= set(rows)

    trouble: list[StatusChange] = []
    cancellations: list[StatusChange] = []
    refunds: list[StatusChange] = []
    warranty: list[StatusChange] = []

    def was_trouble_mid(key: str) -> bool:
        return any(_status_of(rows, key) == STATUS_TROUBLE for _, rows in snapshots)

    for key in keys:
        s0 = _status_of(oldest, key)
        sn = _status_of(newest, key)
        prev = oldest.get(key)
        curr = newest.get(key)
        if prev is None:
            prev = _series_for(snapshots, key, last=False)
        if curr is None:
            curr = _series_for(snapshots, key, last=True)
        if prev is None or curr is None:
            continue

        qty_prev = parse_numeric(prev.get(COL_QTY))
        qty_curr = parse_numeric(curr.get(COL_QTY))
        category = row_str(prev, COL_CATEGORY) or row_str(curr, COL_CATEGORY)
        qty_notes = build_notes(category, qty_prev, qty_curr)
        commercial_notes = build_commercial_notes(prev, curr, qty_prev, qty_curr)
        base = build_status_change(
            key,
            prev,
            curr,
            change_kind="",
            prev_status=s0 or "(нет в первом снимке)",
            curr_status=sn or "(нет в последнем снимке)",
            notes=qty_notes,
        )

        if sn == STATUS_TROUBLE:
            kind = "ongoing" if s0 == STATUS_TROUBLE else "entered"
            extra = (
                ["в TROUBLE оба конца периода"]
                if kind == "ongoing"
                else ["новый TROUBLE за период (ещё открыт)"]
            )
            trouble.append(
                replace(
                    base,
                    change_kind=kind,
                    notes=append_trouble_notes(qty_notes, commercial_notes, extra),
                )
            )
        elif s0 == STATUS_TROUBLE and sn in RESOLVED_FROM_TROUBLE:
            trouble.append(
                replace(
                    base,
                    change_kind="resolved",
                    notes=append_trouble_notes(qty_notes, commercial_notes, []),
                )
            )
        elif s0 == STATUS_TROUBLE and sn in TERMINAL_BAD:
            trouble.append(
                replace(
                    base,
                    change_kind="cancelled_from_trouble",
                    notes=[*qty_notes, *commercial_notes, "TROUBLE → отмена/возврат"],
                )
            )
        elif s0 == STATUS_TROUBLE and sn == STATUS_WARRANTY:
            trouble.append(
                replace(
                    base,
                    change_kind="warranty_from_trouble",
                    notes=[*qty_notes, *commercial_notes, "TROUBLE → гарантия"],
                )
            )
        elif s0 != STATUS_TROUBLE and was_trouble_mid(key):
            if sn in RESOLVED_FROM_TROUBLE:
                trouble.append(
                    replace(
                        base,
                        change_kind="resolved",
                        notes=append_trouble_notes(
                            qty_notes, commercial_notes, ["вошёл и закрыт внутри периода"]
                        ),
                    )
                )
            elif sn in TERMINAL_BAD:
                trouble.append(
                    replace(
                        base,
                        change_kind="cancelled_from_trouble",
                        notes=[
                            *qty_notes,
                            *commercial_notes,
                            "TROUBLE → отмена/возврат внутри периода",
                        ],
                    )
                )
            elif sn == STATUS_WARRANTY:
                trouble.append(
                    replace(
                        base,
                        change_kind="warranty_from_trouble",
                        notes=[*qty_notes, *commercial_notes, "TROUBLE → гарантия внутри периода"],
                    )
                )

        if (s0 in CANCEL_SOURCE or s0 is None) and sn == STATUS_CANCEL:
            cancellations.append(
                replace(base, change_kind="cancelled", notes=[*qty_notes, *commercial_notes])
            )
        elif (s0 in CANCEL_SOURCE or s0 is None) and sn == STATUS_REFUND:
            refunds.append(
                replace(base, change_kind="refunded", notes=[*qty_notes, *commercial_notes])
            )
        if (s0 in CANCEL_SOURCE or s0 is None) and sn == STATUS_WARRANTY:
            warranty.append(
                replace(base, change_kind="warranty", notes=[*qty_notes, *commercial_notes])
            )

    _apply_history_duration(trouble, snapshots)
    return trouble, cancellations, refunds, warranty


def avg(vals: list[float | int | None]) -> float | None:
    clean = [float(v) for v in vals if v is not None]
    return mean(clean) if clean else None


def trouble_kpis(events: list[StatusChange], period_days: int) -> dict[str, Any]:
    entered = [e for e in events if e.change_kind == "entered"]
    ongoing = [e for e in events if e.change_kind == "ongoing"]
    resolved = [e for e in events if e.change_kind == "resolved"]
    cancelled = [e for e in events if e.change_kind == "cancelled_from_trouble"]
    warranty = [e for e in events if e.change_kind == "warranty_from_trouble"]

    at_start = len(ongoing) + len(resolved) + len(cancelled) + len(warranty)
    closed = len(resolved) + len(cancelled) + len(warranty)
    resolve_rate = (len(resolved) / at_start * 100) if at_start else None
    cancel_rate = (len(cancelled) / at_start * 100) if at_start else None

    ongoing_days = [e.trouble_min_days for e in ongoing if e.trouble_min_days is not None]
    resolved_days = [
        e.trouble_min_days
        for e in resolved
        if e.trouble_min_days is not None
        and (e.trouble_max_days is None or e.trouble_max_days == e.trouble_min_days)
    ]
    resolved_margin = [d for e in resolved if (d := meaningful_margin_delta(e)) is not None]
    positive = sum(1 for d in resolved_margin if d > COMMERCIAL_EPS)
    negative = sum(1 for d in resolved_margin if d < -COMMERCIAL_EPS)

    return {
        "entered": len(entered),
        "ongoing": len(ongoing),
        "resolved": len(resolved),
        "cancelled": len(cancelled),
        "warranty": len(warranty),
        "at_start": at_start,
        "resolve_rate": resolve_rate,
        "cancel_rate": cancel_rate,
        "avg_ongoing_days": avg(ongoing_days),
        "avg_resolved_days": avg(resolved_days),
        "avg_margin_delta": avg(resolved_margin),
        "sum_margin_delta": sum(resolved_margin) if resolved_margin else 0.0,
        "resolved_positive": positive,
        "resolved_negative": negative,
        "period_days": period_days,
    }


FASTAIR_CSS = """
:root {
  --bg:#e7ecef; --card:#ffffff; --ink:#202020; --muted:#4c4c4c;
  --navy:#022f40; --cyan:#d5fbff; --orange:#fe621d; --line:#c4c4c4;
}
* { box-sizing:border-box; }
body {
  margin:0; font-family: Arial, Calibri, Helvetica, sans-serif;
  color:var(--ink); background: linear-gradient(180deg, #022f40 0 160px, var(--bg) 160px);
  line-height:1.45;
}
.wrap { max-width:1200px; margin:0 auto; padding:28px 20px 64px; }
.brand { display:flex; align-items:center; justify-content:space-between; gap:12px; color:#fff; margin-bottom:18px; }
.brand-mark { font-size:13px; letter-spacing:.12em; text-transform:uppercase; background:var(--cyan); color:var(--navy); padding:6px 10px; font-weight:700; }
h1 { font-size:30px; margin:0 0 4px; color:#fff; font-weight:700; }
.period { color:#fff; font-size:30px; font-weight:700; margin:0 0 8px; }
.sub { color:#d5fbff; margin-bottom:18px; font-size:14px; }
h2 { font-size:20px; margin:28px 0 12px; color:var(--navy); }
h3 { font-size:16px; margin:18px 0 10px; color:var(--navy); }
.card { background:var(--card); border:1px solid var(--line); border-radius:8px; padding:18px; }
.kpi-grid { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:12px; margin:12px 0 18px; }
.kpi-grid > div { background:#f7fbfc; border:1px solid #d7e2e6; border-radius:8px; padding:12px 14px; }
.label { font-size:11px; text-transform:uppercase; letter-spacing:.04em; color:#757575; }
.value { font-size:20px; font-weight:700; margin-top:4px; font-variant-numeric:tabular-nums; color:var(--navy); }
.muted { color:#757575; font-size:12px; margin-top:4px; }
.trouble-block { border:2px solid var(--orange); border-radius:10px; padding:16px; background:#fff9f5; margin:20px 0; }
.trouble-block h2 { margin-top:0; color:#b3470f; }
details.group { background:var(--card); border:1px solid var(--line); border-radius:8px; margin:10px 0; overflow:hidden; }
details.group summary { cursor:pointer; list-style:none; padding:14px 16px; display:flex; flex-wrap:wrap; gap:8px; align-items:center; background:linear-gradient(90deg,#022f40 0%,#022f40 12px,#fff 12px); }
details.group summary::-webkit-details-marker { display:none; }
.group-name { font-weight:700; font-size:15px; margin-right:auto; color:var(--navy); padding-left:8px; }
.pill { font-size:12px; background:var(--cyan); border:1px solid #b6e8ef; color:var(--navy); border-radius:4px; padding:4px 10px; font-weight:700; }
.pill.warn { background:#ffe8d9; border-color:#ffc4a3; color:#b3470f; }
.pill.ok { background:#e6f7ea; border-color:#b8e6c1; color:#1f6b32; }
.group-body { padding:0 14px 16px; border-top:1px solid var(--line); }
details.client { margin:8px 0; border-color:#d7e2e6; }
details.client summary { padding:10px 14px; background:linear-gradient(90deg,#022f40 0%,#022f40 8px,#f7fbfc 8px); }
details.client .group-name { font-size:14px; font-weight:600; }
details.client .group-body { padding:0 8px 12px; }
table { width:100%; border-collapse:collapse; font-size:13px; }
th, td { border-bottom:1px solid #e8e8e8; padding:8px 6px; vertical-align:top; text-align:left; }
th { font-size:11px; text-transform:uppercase; letter-spacing:.03em; color:#4c4c4c; background:#f3f7f8; }
td.num { text-align:right; font-variant-numeric:tabular-nums; white-space:nowrap; }
td.desc { max-width:280px; }
.problem { color:#b3470f; font-weight:600; margin-bottom:4px; }
.pos { color:#1f6b32; } .neg { color:#c0392b; }
.footer { margin-top:28px; color:#4c4c4c; font-size:12px; }
@media (max-width:900px) { .kpi-grid { grid-template-columns:repeat(2,minmax(0,1fr)); } }
"""


def group_by_customer(items: list[StatusChange]) -> list[tuple[str, list[StatusChange]]]:
    by_client: dict[str, list[StatusChange]] = defaultdict(list)
    for e in items:
        name = e.customer.strip() if e.customer and e.customer.strip() else "(без клиента)"
        by_client[name].append(e)
    return sorted(by_client.items(), key=lambda x: (-sum(i.sale_curr for i in x[1]), x[0]))


def fmt_comment_cell(e: StatusChange) -> str:
    parts: list[str] = []
    if e.problem_notes:
        parts.append(f"<div class='problem'>{esc('; '.join(e.problem_notes))}</div>")
    notes = "; ".join(e.notes)
    if notes:
        parts.append(esc(notes))
    elif not e.problem_notes:
        parts.append("—")
    return "".join(parts)


def sum_plan_sale(items: list[StatusChange]) -> float:
    return sum(e.sale_curr if e.sale_curr else e.sale_prev for e in items)


def sum_plan_margin(items: list[StatusChange]) -> float:
    return sum(e.margin_curr for e in items)


def sum_margin_prev(items: list[StatusChange]) -> float:
    return sum(e.margin_prev for e in items)


def header_metric_pills(
    items: list[StatusChange],
    *,
    margin_mode: str,
    pill_class: str = "",
    include_clients: bool = True,
    extra_pills: list[str] | None = None,
) -> str:
    """Dropdown header: rows, clients, line-total sale, plan margin or was/became."""
    count_class = f"pill {pill_class}".strip()
    pills = [f'<span class="{count_class}">{len(items)} поз.</span>']
    if include_clients:
        pills.append(f'<span class="pill">{len(group_by_customer(items))} кли.</span>')
    pills.append(f'<span class="pill">продажная итого {fmt_money(sum_plan_sale(items), 0)} USD</span>')
    if margin_mode == "before_after":
        pills.append(f'<span class="pill">маржа было {fmt_money(sum_margin_prev(items), 0)} USD</span>')
        pills.append(f'<span class="pill">маржа стало {fmt_money(sum_plan_margin(items), 0)} USD</span>')
        hung = [
            e.trouble_min_days
            for e in items
            if e.trouble_min_days
            and (e.trouble_max_days is None or e.trouble_max_days == e.trouble_min_days)
        ]
        if hung:
            pills.append(
                f'<span class="pill">решили за примерно {round(mean(hung))} дн. ср.</span>'
            )
    else:
        pills.append(f'<span class="pill">маржа {fmt_money(sum_plan_margin(items), 0)} USD</span>')
    if extra_pills:
        pills.extend(extra_pills)
    return "".join(pills)


def render_trouble_table(items: list[StatusChange], *, section: str | None = None) -> str:
    if not items:
        return "<p class='muted'>Нет изменений за период.</p>"
    kind = section or items[0].change_kind
    show_days = kind in {"ongoing", "resolved"}
    show_margin = kind == "resolved"
    sort_key = (
        (lambda x: (-(x.trouble_min_days or 0), x.pn, x.description))
        if kind in {"ongoing", "resolved"}
        else (lambda x: (x.pn, x.description))
    )
    headers = ["P/N", "Описание", "Cat.", "Поставщик"]
    if show_days:
        headers.append("Решили за" if kind == "resolved" else "В TROUBLE")
    headers.append("Продажа USD")
    if show_margin:
        headers.extend(["маржа было", "маржа стало"])
    headers.append("Комментарий")

    rows = []
    for e in sorted(items, key=sort_key):
        cells = [
            f"<td>{esc(e.pn)}</td>",
            f"<td class='desc'>{esc(e.description) or '—'}</td>",
            f"<td>{esc(e.category)}</td>",
            f"<td>{esc(e.supplier_display)}</td>",
        ]
        if show_days:
            if kind == "resolved":
                if e.trouble_max_days is None:
                    days_txt = fmt_hung_days(e.trouble_min_days, unknown_start=True)
                elif e.trouble_max_days == e.trouble_min_days:
                    days_txt = fmt_hung_days(e.trouble_min_days, unknown_start=False)
                else:
                    days_txt = fmt_days_range(e.trouble_min_days, e.trouble_max_days)
            else:
                days_txt = fmt_days_range(e.trouble_min_days, e.trouble_max_days)
            cells.append(f"<td class='num'>{days_txt}</td>")
        cells.append(f"<td class='num'>{fmt_sale_cell(e)}</td>")
        if show_margin:
            d = e.margin_curr - e.margin_prev
            css = "pos" if d > COMMERCIAL_EPS else ("neg" if d < -COMMERCIAL_EPS else "")
            cells.append(f"<td class='num'>{fmt_money(e.margin_prev, 0)}</td>")
            cells.append(f"<td class='num {css}'>{fmt_money(e.margin_curr, 0)}</td>")
        cells.append(f"<td>{fmt_comment_cell(e)}</td>")
        rows.append("<tr>" + "".join(cells) + "</tr>")
    header_html = "".join(f"<th>{h}</th>" for h in headers)
    return f"""<table>
  <thead><tr>{header_html}</tr></thead>
  <tbody>{''.join(rows)}</tbody>
</table>"""


def render_cancel_table(items: list[StatusChange]) -> str:
    if not items:
        return "<p class='muted'>Нет изменений за период.</p>"
    rows = []
    for e in sorted(items, key=lambda x: (x.pn, x.description)):
        margin_txt, margin_cls = fmt_margin_delta_cell(e)
        kind = "отмена" if e.curr_status == STATUS_CANCEL else "возврат"
        if e.change_kind == "cancelled_from_trouble":
            kind = "из TROUBLE → " + kind
        rows.append(
            f"""<tr>
  <td>{esc(e.pn)}</td>
  <td class='desc'>{esc(e.description) or '—'}</td>
  <td>{esc(e.category)}</td>
  <td>{esc(kind)}</td>
  <td class='num'>{fmt_money(e.sale_prev, 0)}</td>
  <td class='num {margin_cls}'>{margin_txt}</td>
  <td>{fmt_comment_cell(e)}</td>
</tr>"""
        )
    return f"""<table>
  <thead><tr>
    <th>P/N</th><th>Описание</th><th>Cat.</th><th>Тип</th>
    <th>Продажа USD</th><th>Δ маржа</th><th>Комментарий</th>
  </tr></thead>
  <tbody>{''.join(rows)}</tbody>
</table>"""


def render_client_groups(
    items: list[StatusChange],
    table_fn,
    *,
    margin_mode: str = "plan",
) -> str:
    if not items:
        return "<p class='muted'>Нет изменений за период.</p>"
    blocks = []
    for client, client_items in group_by_customer(items):
        blocks.append(
            f"""
<details class="group client">
  <summary>
    <span class="group-name">{esc(client)}</span>
    {header_metric_pills(client_items, margin_mode=margin_mode, include_clients=False)}
  </summary>
  <div class="group-body">{table_fn(client_items)}</div>
</details>"""
        )
    return "".join(blocks)


def render_section(
    title: str,
    items: list[StatusChange],
    table_fn,
    pill_class: str = "",
    *,
    margin_mode: str = "plan",
) -> str:
    if not items:
        return f"""
<details class="group">
  <summary>
    <span class="group-name">{esc(title)}</span>
    {header_metric_pills(items, margin_mode=margin_mode, pill_class=pill_class)}
  </summary>
  <div class="group-body"><p class='muted'>Нет изменений за период.</p></div>
</details>"""
    inner = render_client_groups(items, table_fn, margin_mode=margin_mode)
    return f"""
<details class="group">
  <summary>
    <span class="group-name">{esc(title)}</span>
    {header_metric_pills(items, margin_mode=margin_mode, pill_class=pill_class)}
  </summary>
  <div class="group-body">{inner}</div>
</details>"""


def merge_cancel_refund(
    cancellations: list[StatusChange],
    refunds: list[StatusChange],
    trouble: list[StatusChange],
) -> list[StatusChange]:
    """Combine cancel/refund events, dedupe by row key."""
    by_key: dict[str, StatusChange] = {}
    for items in (cancellations, refunds):
        for e in items:
            by_key[e.key] = e
    for e in trouble:
        if e.change_kind == "cancelled_from_trouble":
            by_key[e.key] = e
    return list(by_key.values())


def render_html_report(
    prev_date: date,
    curr_date: date,
    trouble: list[StatusChange],
    cancellations: list[StatusChange],
    refunds: list[StatusChange],
    warranty: list[StatusChange],
    *,
    history_start: date | None = None,
) -> str:
    period = f"{prev_date.strftime('%d.%m.%Y')} → {curr_date.strftime('%d.%m.%Y')}"
    period_days = max(1, (curr_date - prev_date).days)
    kpis = trouble_kpis(trouble, period_days)
    history_note = ""
    if history_start and history_start < prev_date:
        history_note = f" · история нерешённых с {history_start.strftime('%d.%m.%Y')}"

    new_trouble = [e for e in trouble if e.change_kind == "entered"]
    unresolved_trouble = [e for e in trouble if e.change_kind == "ongoing"]
    resolved_trouble = [e for e in trouble if e.change_kind == "resolved"]
    cancel_refund = merge_cancel_refund(cancellations, refunds, trouble)

    sections = "\n".join([
        render_section(
            "Новые TROUBLE",
            new_trouble,
            lambda xs: render_trouble_table(xs, section="entered"),
            "warn",
            margin_mode="plan",
        ),
        render_section(
            "Нерешённые TROUBLE",
            unresolved_trouble,
            lambda xs: render_trouble_table(xs, section="ongoing"),
            "warn",
            margin_mode="plan",
        ),
        render_section(
            "Решённые TROUBLE",
            resolved_trouble,
            lambda xs: render_trouble_table(xs, section="resolved"),
            "ok",
            margin_mode="before_after",
        ),
        render_section(
            "Отмены и возвраты",
            cancel_refund,
            render_cancel_table,
            "warn",
            margin_mode="plan",
        ),
    ])

    return f"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8"/>
<title>TROUBLE / CANCEL — {curr_date.strftime('%d.%m.%Y')}</title>
<style>{FASTAIR_CSS}</style>
</head>
<body>
<div class="wrap">
  <div class="brand"><div class="brand-mark">not so fastair</div><div style="font-size:13px;opacity:.9">TROUBLE / CANCEL</div></div>
  <h1>TROUBLE / CANCEL</h1>
  <div class="period">{period}</div>
  <div class="sub">{period_days} дн. · сравнение выгрузок ТАЗ{history_note}</div>

  <section class="card">
    <h2 style="margin-top:0">Сводка</h2>
    <div class="kpi-grid">
      <div><div class="label">Новые TROUBLE</div><div class="value">{len(new_trouble)}</div></div>
      <div><div class="label">Нерешённые TROUBLE</div><div class="value">{len(unresolved_trouble)}<div class="muted">{'по истории с ' + history_start.strftime('%d.%m') if history_start else f'≥{period_days} дн. в обоих снимках'}</div></div></div>
      <div><div class="label">Решённые TROUBLE</div><div class="value">{len(resolved_trouble)}<div class="muted">{fmt_pct(kpis['resolve_rate'])} от стартовых</div></div></div>
      <div><div class="label">Решили за</div><div class="value">{fmt_hung_days(round(kpis['avg_resolved_days']) if kpis['avg_resolved_days'] is not None else None, unknown_start=False)}<div class="muted">в среднем по решённым · погрешность ~неделя</div></div></div>
      <div><div class="label">Отмены + возвраты</div><div class="value">{len(cancel_refund)}</div></div>
      <div><div class="label">маржа было (решённые)</div><div class="value">{fmt_money(sum_margin_prev(resolved_trouble), 0)}</div></div>
      <div><div class="label">маржа стало (решённые)</div><div class="value">{fmt_money(sum_plan_margin(resolved_trouble), 0)}</div></div>
      <div><div class="label">Гарантии (новые)</div><div class="value">{len(warranty)}</div></div>
    </div>
    <p class="muted">TROUBLE: EXP — кол-во/ед. изм.; ROTABLE — замена юнита. «Новые» — смена статуса на TROUBLE за период и заказы, взятые в работу в период (появились в ТАЗ уже в TROUBLE). Решённые: «решили за примерно N дн.» — от первой TROUBLE до последней в снимках (если уже TROUBLE в самом старом файле — «≥»). Погрешность около недели. Отмена/возврат: было NOT PAID / PAID / TROUBLE → CANCELLED / REFUND.</p>
  </section>

  {sections}

  <div class="footer">
    not so fastair · источник ТАЗ · {esc(prev_date.isoformat())} → {esc(curr_date.isoformat())} ·
    ключ строки: счёт + P/N + описание
  </div>
</div>
</body>
</html>
"""


def save_html_report(path: Path, html_out: str) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(html_out, encoding="utf-8")
    zip_path = path.with_suffix(".zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.write(path, path.name)
    print(f"Saved {path}")
    print(f"Saved {zip_path}")
    return zip_path


def report_filename(prev: date, curr: date, *, full: bool = False) -> str:
    if full:
        return f"trouble_cancel_{prev.strftime('%d_%m')}-{curr.strftime('%d_%m_%Y')}_full.html"
    if prev.month == curr.month and prev.year == curr.year:
        return f"trouble_cancel_{prev.strftime('%d')}-{curr.strftime('%d_%m_%Y')}.html"
    return f"trouble_cancel_{prev.strftime('%d_%m')}-{curr.strftime('%d_%m_%Y')}.html"


def parse_snapshot_arg(text: str) -> tuple[date, Path]:
    if "=" not in text:
        raise ValueError(f"Expected DATE=PATH, got {text!r}")
    date_s, path_s = text.split("=", 1)
    return parse_report_date(date_s), Path(path_s)


def generate_from_chain(
    snapshots: list[tuple[date, dict[str, pd.Series]]],
    output_dir: Path,
    *,
    weekly: bool = True,
    full: bool = True,
) -> list[Path]:
    """Write weekly consecutive reports (+ optional full-period) from a dated chain."""
    snapshots = sorted(snapshots, key=lambda x: x[0])
    written: list[Path] = []
    history_start = snapshots[0][0]

    if weekly:
        for i in range(1, len(snapshots)):
            prev_date, prev_rows = snapshots[i - 1]
            curr_date, curr_rows = snapshots[i]
            period_days = max(1, (curr_date - prev_date).days)
            history = snapshots[: i + 1]
            trouble, cancellations, refunds, warranty = compare_snapshots(
                prev_rows, curr_rows, period_days, history=history
            )
            html_out = render_html_report(
                prev_date,
                curr_date,
                trouble,
                cancellations,
                refunds,
                warranty,
                history_start=history_start if history_start < prev_date else None,
            )
            out = output_dir / report_filename(prev_date, curr_date)
            save_html_report(out, html_out)
            written.append(out)
            print(
                f"  week {prev_date:%d.%m}→{curr_date:%d.%m}: "
                f"TROUBLE {len(trouble)} | cancel {len(cancellations)} | refund {len(refunds)}"
            )

    if full and len(snapshots) >= 2:
        prev_date, curr_date = snapshots[0][0], snapshots[-1][0]
        trouble, cancellations, refunds, warranty = compare_full_period(snapshots)
        html_out = render_html_report(
            prev_date,
            curr_date,
            trouble,
            cancellations,
            refunds,
            warranty,
            history_start=prev_date,
        )
        out = output_dir / report_filename(prev_date, curr_date, full=True)
        save_html_report(out, html_out)
        written.append(out)
        print(
            f"  full {prev_date:%d.%m}→{curr_date:%d.%m}: "
            f"TROUBLE {len(trouble)} | cancel {len(cancellations)} | refund {len(refunds)}"
        )

    if written:
        bundle = output_dir / "trouble_cancel_all.zip"
        with zipfile.ZipFile(bundle, "w", zipfile.ZIP_DEFLATED) as zf:
            for path in written:
                zf.write(path, path.name)
        print(f"Saved {bundle}")
    return written


def main() -> None:
    parser = argparse.ArgumentParser(description="HTML report: TAZ diff (TROUBLE, cancel, warranty)")
    parser.add_argument("--previous", "-p", type=Path, help="Earlier TAZ snapshot")
    parser.add_argument("--current", "-c", type=Path, help="Later TAZ snapshot")
    parser.add_argument("--previous-date", type=str, help="Date of previous snapshot DD.MM.YYYY")
    parser.add_argument("--current-date", type=str, help="Date of current snapshot DD.MM.YYYY")
    parser.add_argument("--output", "-o", type=Path, default=Path("output/changes/trouble_cancel.html"))
    parser.add_argument(
        "--snapshots",
        nargs="+",
        help="Chain of DATE=PATH snapshots (oldest first). Writes weekly reports + full period.",
    )
    parser.add_argument("--output-dir", type=Path, default=Path("output/changes"))
    args = parser.parse_args()

    if args.snapshots:
        dated_paths = [parse_snapshot_arg(item) for item in args.snapshots]
        dated_paths.sort(key=lambda x: x[0])
        loaded: list[tuple[date, dict[str, pd.Series]]] = []
        for snap_date, path in dated_paths:
            print(f"Loading TAZ {snap_date:%d.%m.%Y}: {path}")
            loaded.append((snap_date, index_rows(load_taz(path))))
        generate_from_chain(loaded, args.output_dir)
        return

    if not args.previous or not args.current:
        parser.error("Provide --previous and --current, or --snapshots DATE=PATH ...")

    prev_date = parse_report_date(args.previous_date) if args.previous_date else None
    curr_date = parse_report_date(args.current_date) if args.current_date else None
    if prev_date is None:
        prev_date = date.fromtimestamp(args.previous.stat().st_mtime)
    if curr_date is None:
        curr_date = date.fromtimestamp(args.current.stat().st_mtime)

    print(f"Loading previous TAZ: {args.previous} ({prev_date})")
    print(f"Loading current TAZ:  {args.current} ({curr_date})")
    prev_rows = index_rows(load_taz(args.previous))
    curr_rows = index_rows(load_taz(args.current))
    period_days = max(1, (curr_date - prev_date).days)

    trouble, cancellations, refunds, warranty = compare_snapshots(prev_rows, curr_rows, period_days)
    html_out = render_html_report(prev_date, curr_date, trouble, cancellations, refunds, warranty)
    save_html_report(args.output, html_out)
    print(
        f"TROUBLE events: {len(trouble)} | cancellations: {len(cancellations)} | "
        f"refunds: {len(refunds)} | warranty: {len(warranty)}"
    )


if __name__ == "__main__":
    main()
