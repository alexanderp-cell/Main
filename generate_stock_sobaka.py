#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Формирование файла STOCK_SOBAKA.xlsx.

Те же правила, что и у ежемесячного обновления склада
(generate_ready_stock.py в ветке обновления склада):

  1. Оставляет только «зелёные» строки (заливка 00FF00) с чёрным текстом.
  2. Чистит партийные номера: служебные слова, комментарии и кириллица.
     Ячейка Alt. PN с кириллицей считается комментарием и очищается.
     Несколько альтернативных номеров пишутся через запятую.
  3. Удаляет кириллицу из всех значений.
  4. Объединяет строки с одинаковыми Part Number и состоянием
     (NE и NEW — одно состояние), количества суммируются.
  5. Оформление совпадает с образцом готовых остатков: шрифты, заливки,
     границы, ширина столбцов, высота строк, выравнивание, форматы.

Отличия STOCK SOBAKA:
  - в готовой таблице нет столбца AC Type (тип ВС);
  - в готовой таблице нет столбца размещения на складе;
  - по столбцу размещения из исходника полностью удаляются строки,
    в которых есть «Сириус» / SIRIUS.

Запуск:
    python3 generate_stock_sobaka.py

Исходник SOURCE_STOCK_SOBAKA.xlsx и шаблон оформления не изменяются.
"""

import datetime
import re
import sys
import warnings
from collections import Counter, OrderedDict
from copy import copy

import openpyxl

warnings.filterwarnings("ignore", message="Data Validation extension is not supported")

TEMPLATE_FILE = "template_stock_sobaka.xlsx"
SOURCE_FILE = "SOURCE_STOCK_SOBAKA.xlsx"
OUTPUT_FILE = "STOCK_SOBAKA.xlsx"
SHEET_TITLE = "STOCK SOBAKA"

# Итоговые столбцы (как в образце склада, без AC Type):
# 1 Part Number | 2 Alt. PN | 3 DESCRIPTION | 4 Quantity | 5 Condition
COL_PN, COL_ALT, COL_DESC, COL_QTY, COL_COND = 1, 2, 3, 4, 5
NCOLS = 5
HEADERS = ("Part Number", "Alt. PN", "DESCRIPTION", "Quantity", "Condition")

# Столбцы исходника ищем по заголовку, а не по номеру.
SOURCE_HEADER_PATTERNS = {
    "pn":   [r"P/?N", r"PART\s*NUMBER", r"ПАРТ"],
    "alt":  [r"ALT"],
    "desc": [r"DESCR", r"НАИМЕН", r"ОПИСАН"],
    "qty":  [r"QTY", r"QUANT", r"КОЛ[-\s]?ВО", r"КОЛИЧ"],
    "cond": [r"СОСТОЯН", r"CONDITION", r"COND\b"],
    "loc":  [r"РАЗМЕЩ", r"LOCATION", r"СКЛАД"],
}

GREEN = "FF00FF00"

CYR_TO_LAT = {
    "А": "A", "В": "B", "С": "C", "Е": "E", "Н": "H", "К": "K", "М": "M",
    "О": "O", "Р": "P", "Т": "T", "Х": "X", "У": "Y", "І": "I",
    "а": "a", "в": "b", "с": "c", "е": "e", "н": "h", "к": "k", "м": "m",
    "о": "o", "р": "p", "т": "t", "х": "x", "у": "y",
}
CYRILLIC_RE = re.compile(r"[А-Яа-яЁё]")
SIRIUS_RE = re.compile(r"сириус|sirius", re.IGNORECASE)

NOISE_WORDS = [
    "PART NUMBER", "P/N", "PN", "ALT PN", "ALT", "QTY", "PCS", "EA", "EACH",
    "STOCK", "STK", "MOSCOW", "MOW", "MSK", "SIRIUS", "SIR", "VKO", "FAI", "FEI",
    "OFFERED", "QUOTED", "MOQ", "REL", "AMDT",
]
NOISE_RE = re.compile(
    r"(?<![A-Za-z0-9])(?:" + "|".join(re.escape(w) for w in NOISE_WORDS) + r")(?![A-Za-z0-9])",
    re.IGNORECASE,
)


def transliterate(text):
    return "".join(CYR_TO_LAT.get(ch, ch) for ch in text)


def strip_cyrillic(text):
    return CYRILLIC_RE.sub("", text)


def cell_str(value):
    """Приводит значение ячейки к строке.

    Excel часто превращает партийные номера вида «8296-01-01» в даты.
    Их нужно вернуть в текстовый вид, а не «8296-01-01 00:00:00».
    """
    if value is None:
        return ""
    if isinstance(value, datetime.date):
        text = "%04d-%02d-%02d" % (value.year, value.month, value.day)
        if isinstance(value, datetime.datetime) and (value.hour or value.minute or value.second):
            text += " %02d:%02d:%02d" % (value.hour, value.minute, value.second)
        return text
    return str(value).strip()


def is_green(cell):
    fill = cell.fill
    if fill is None or fill.patternType is None:
        return False
    fg = fill.fgColor
    return fg is not None and fg.type == "rgb" and str(fg.rgb) == GREEN


def is_black_font(cell):
    color = cell.font.color
    if color is None:
        return True
    if color.type == "theme" and color.theme == 1:
        return True
    if color.type == "rgb" and str(color.rgb) in ("FF000000", "00000000"):
        return True
    return False


def is_sirius_location(raw):
    """Строка склада Сириус — такие позиции в STOCK SOBAKA не попадают."""
    return bool(SIRIUS_RE.search(cell_str(raw)))


def clean_token(token):
    token = NOISE_RE.sub(" ", token)
    token = re.sub(r"\s+", " ", token).strip()
    token = token.strip(" .-/;,")
    return token.strip()


def is_valid_pn(token):
    if not token:
        return False
    if not re.search(r"\d", token):
        return False
    if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9./-]*", token):
        return False
    if token.isdigit() and len(token) < 4:
        return False
    return True


def _complete_pn(piece):
    return bool(re.search(r"\d", piece) and (re.search(r"[./-]", piece) or len(piece) >= 5))


def _short_affix(piece):
    return bool(re.fullmatch(r"[A-Za-z]{1,3}", piece))


def _alt_candidates(entry):
    pieces = entry.split()
    if not pieces:
        return []
    if len(pieces) == 1:
        return pieces
    if all(_complete_pn(p) for p in pieces):
        return pieces
    if len(pieces) == 2:
        a, b = pieces
        glue = (
            (_complete_pn(a) and _short_affix(b))
            or (_short_affix(a) and _complete_pn(b))
            or (_short_affix(a) and re.fullmatch(r"\d{3,}", b))
            or (re.fullmatch(r"\d{3,}", a) and _short_affix(b))
        )
        if glue:
            return ["".join(pieces)]
    return pieces


def clean_main_pn(raw):
    text = cell_str(raw)
    if not text:
        return ""
    text = transliterate(text)
    text = strip_cyrillic(text)
    text = text.replace("\n", " ")
    text = clean_token(text)
    return text


def clean_alt_pn(raw, main_pn=""):
    text = cell_str(raw)
    if not text:
        return ""
    # Ячейки с кириллицей в Alt PN — это комментарии, а не номера.
    if CYRILLIC_RE.search(text):
        return ""
    main_key = main_pn.strip().upper()
    result = []
    for part in re.split(r"[\n,;]+", text):
        cleaned = NOISE_RE.sub(" ", part)
        for candidate in _alt_candidates(cleaned):
            token = candidate.strip(" .-/;,")
            if not is_valid_pn(token):
                continue
            if token.upper() == main_key:
                continue
            if token not in result:
                result.append(token)
    return ", ".join(result)


def clean_text_value(raw):
    text = cell_str(raw)
    if not text:
        return ""
    text = strip_cyrillic(text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def clean_quantity(raw):
    text = cell_str(raw)
    if not text:
        return ""
    text = strip_cyrillic(text).replace("\xa0", " ").strip()
    compact = text.replace(" ", "")
    if re.fullmatch(r"-?\d+", compact):
        return int(compact)
    if re.fullmatch(r"-?\d+[.,]\d+", compact):
        return float(compact.replace(",", "."))
    return text


def canon_condition(cond):
    u = (cond or "").strip().upper()
    if u in ("NE", "NEW"):
        return "NEW"
    return u


def _uniq_join(values):
    out = []
    for v in values:
        v = (v or "").strip()
        if v and v not in out:
            out.append(v)
    return ", ".join(out)


def consolidate_records(records):
    """Одинаковые Part Number и состояние схлопываются в одну строку."""
    groups = OrderedDict()
    no_key = []
    for rec in records:
        pn_key = rec["pn"].strip().upper()
        if not pn_key:
            rec = dict(rec, cond=canon_condition(rec["cond"]))
            no_key.append(rec)
            continue
        groups.setdefault((pn_key, canon_condition(rec["cond"])), []).append(rec)

    result = []
    for (pn_key, cond_key), recs in groups.items():
        numeric = [r["qty"] for r in recs if isinstance(r["qty"], (int, float))]
        total = sum(numeric)
        if numeric and all(isinstance(q, int) for q in numeric):
            total = int(total)
        non_numeric = [str(r["qty"]).strip() for r in recs
                       if not isinstance(r["qty"], (int, float)) and str(r["qty"]).strip()]
        if numeric and non_numeric:
            qty = _uniq_join([str(total)] + non_numeric)
        elif numeric:
            qty = total
        else:
            qty = _uniq_join(non_numeric)

        alt_tokens = []
        for r in recs:
            for tok in re.split(r"[\n,;]+", str(r["alt"])):
                tok = tok.strip()
                if tok and tok.upper() != pn_key and tok not in alt_tokens:
                    alt_tokens.append(tok)
        alt = ", ".join(alt_tokens)

        descs = [r["desc"].strip() for r in recs if r["desc"].strip()]
        desc = Counter(descs).most_common(1)[0][0] if descs else ""

        result.append({
            "pn": recs[0]["pn"],
            "alt": alt,
            "desc": desc,
            "qty": qty,
            "cond": cond_key,
        })

    result.extend(no_key)
    return result


def detect_source_columns(ws):
    headers = {}
    for c in range(1, ws.max_column + 1):
        val = ws.cell(1, c).value
        if val not in (None, ""):
            headers[c] = str(val).strip().upper()

    mapping = {key: None for key in SOURCE_HEADER_PATTERNS}
    used = set()
    # Сначала специфичные поля, чтобы «ALT PN» не был принят за «P/N».
    for key in ("loc", "alt", "cond", "qty", "desc", "pn"):
        for pattern in SOURCE_HEADER_PATTERNS[key]:
            rx = re.compile(pattern)
            for col, text in headers.items():
                if col in used:
                    continue
                if rx.search(text):
                    mapping[key] = col
                    used.add(col)
                    break
            if mapping[key] is not None:
                break
    return mapping


def main():
    out_wb = openpyxl.load_workbook(TEMPLATE_FILE)
    out_ws = out_wb.active
    out_ws.title = SHEET_TITLE

    template_style = {}
    for c in range(1, NCOLS + 1):
        src = out_ws.cell(2, c)
        template_style[c] = {
            "font": copy(src.font),
            "fill": copy(src.fill),
            "border": copy(src.border),
            "alignment": copy(src.alignment),
            "number_format": src.number_format,
        }
        header = out_ws.cell(1, c, HEADERS[c - 1])
        # Заголовок уже оформлен в шаблоне; значение фиксируем явно.
        header.value = HEADERS[c - 1]
    data_row_height = out_ws.row_dimensions[2].height or out_ws.sheet_format.defaultRowHeight
    header_row_height = out_ws.row_dimensions[1].height

    src_wb = openpyxl.load_workbook(SOURCE_FILE)
    src_ws = src_wb.active
    cols = detect_source_columns(src_ws)
    missing = [k for k in ("pn", "alt", "desc", "qty", "cond", "loc") if cols[k] is None]
    if missing:
        raise SystemExit("В исходнике не найдены столбцы: %s" % ", ".join(missing))

    def src_val(row, key):
        return src_ws.cell(row, cols[key]).value

    max_src_col = max(cols.values())
    records = []
    stats = {
        "total": 0,
        "kept": 0,
        "dropped_sirius": 0,
        "dropped_color": 0,
        "dropped_font": 0,
    }
    for row in range(2, src_ws.max_row + 1):
        if all(src_ws.cell(row, c).value in (None, "") for c in range(1, max_src_col + 1)):
            continue
        stats["total"] += 1
        if is_sirius_location(src_val(row, "loc")):
            stats["dropped_sirius"] += 1
            continue
        pn_cell = src_ws.cell(row, cols["pn"])
        if not is_green(pn_cell):
            stats["dropped_color"] += 1
            continue
        if not is_black_font(pn_cell):
            stats["dropped_font"] += 1
            continue

        pn = clean_main_pn(pn_cell.value)
        records.append({
            "pn": pn,
            "alt": clean_alt_pn(src_val(row, "alt"), pn),
            "desc": clean_text_value(src_val(row, "desc")),
            "qty": clean_quantity(src_val(row, "qty")),
            "cond": clean_text_value(src_val(row, "cond")),
        })
        stats["kept"] += 1

    rows_before_merge = len(records)
    records = consolidate_records(records)
    rows_after_merge = len(records)

    def sort_key(rec):
        # Без типа ВС сортировка по Part Number, затем по состоянию.
        # Строка без номера (масло без P/N) уходит в конец.
        blank = 0 if rec["pn"].strip() else 1
        return (blank, rec["pn"].upper(), rec["cond"].upper())

    records.sort(key=sort_key)

    if out_ws.max_row >= 2:
        out_ws.delete_rows(2, out_ws.max_row - 1)
    if header_row_height:
        out_ws.row_dimensions[1].height = header_row_height

    def write_cell(row, col, value):
        cell = out_ws.cell(row, col, value)
        st = template_style[col]
        cell.font = copy(st["font"])
        cell.fill = copy(st["fill"])
        cell.border = copy(st["border"])
        cell.alignment = copy(st["alignment"])
        cell.number_format = "General" if col == COL_QTY else "@"
        return cell

    for i, rec in enumerate(records):
        r = i + 2
        write_cell(r, COL_PN, rec["pn"])
        write_cell(r, COL_ALT, rec["alt"])
        write_cell(r, COL_DESC, rec["desc"])
        write_cell(r, COL_QTY, rec["qty"] if rec["qty"] != "" else None)
        write_cell(r, COL_COND, rec["cond"])
        out_ws.row_dimensions[r].height = data_row_height

    out_ws.auto_filter.ref = None
    if out_ws.sheet_view.selection:
        out_ws.sheet_view.selection[0].activeCell = "A1"
        out_ws.sheet_view.selection[0].sqref = "A1"
    # Шаблон образца склада хранит ширины и высоты далеко за пределами
    # пяти столбцов. В готовом файле их быть не должно.
    last_row = 1 + len(records)
    for letter in list(out_ws.column_dimensions.keys()):
        if not letter or len(letter) != 1 or letter > "E":
            del out_ws.column_dimensions[letter]
    for idx in list(out_ws.row_dimensions.keys()):
        if isinstance(idx, int) and idx > last_row:
            del out_ws.row_dimensions[idx]
    out_wb.save(OUTPUT_FILE)

    qty_sum = sum(r["qty"] for r in records if isinstance(r["qty"], (int, float)))
    print("Готово: %s" % OUTPUT_FILE)
    print("  строк в источнике (непустых):      %d" % stats["total"])
    print("  убрано строк Сириус:               %d" % stats["dropped_sirius"])
    print("  отброшено по цвету заливки:        %d" % stats["dropped_color"])
    print("  отброшено по цвету шрифта:         %d" % stats["dropped_font"])
    print("  строк после фильтрации:            %d" % rows_before_merge)
    print("  объединено дубликатов PN+состояние:%d" % (rows_before_merge - rows_after_merge))
    print("  записано в результат (строк):       %d" % rows_after_merge)
    print("  сумма количеств:                    %s" % qty_sum)
    return 0


if __name__ == "__main__":
    sys.exit(main())
