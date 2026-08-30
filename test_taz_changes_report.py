#!/usr/bin/env python3
"""Tests for TROUBLE/CANCEL HTML report layout and problem-text extraction."""

from __future__ import annotations

from datetime import date

import pandas as pd

from generate_taz_changes_report import (
    COL_CATEGORY,
    COL_COMMENT,
    COL_CUSTOMER,
    COL_DAYS_TO_DELIVER,
    COL_DEADLINE,
    COL_DESC,
    COL_FEE,
    COL_INVOICE,
    COL_ORDER_DATE,
    COL_PN,
    COL_PURCHASE,
    COL_QTY,
    COL_ROOT_SUPPLIER,
    COL_SALE,
    COL_STATUS,
    COL_SUPPLIER,
    COL_TRANSPORT_PLAN,
    COL_UNIQUE_UNIT,
    STATUS_FINISHED,
    STATUS_TROUBLE,
    compare_snapshots,
    extract_problem_notes,
    format_supplier_display,
    fmt_margin_delta_cell,
    render_html_report,
    render_trouble_table,
    split_unit_code,
)

STATUS_PAID = "2 PAID"


def taz_row(**kwargs) -> pd.Series:
    base = {
        COL_INVOICE: "100",
        COL_STATUS: STATUS_TROUBLE,
        COL_CUSTOMER: "Utair",
        COL_PN: "PN-1",
        COL_DESC: "WIDGET LH",
        COL_CATEGORY: "EXPENDABLE",
        COL_SUPPLIER: "ACME",
        COL_ROOT_SUPPLIER: "ACME",
        COL_UNIQUE_UNIT: "",
        COL_QTY: 1,
        COL_SALE: 1000,
        COL_PURCHASE: 800,
        COL_TRANSPORT_PLAN: 0,
        COL_FEE: 0,
        COL_DEADLINE: date(2026, 9, 1),
        COL_DAYS_TO_DELIVER: 10,
        COL_ORDER_DATE: date(2026, 8, 1),
        COL_COMMENT: "",
    }
    base.update(kwargs)
    return pd.Series(base)


def test_split_unit_code_keeps_reference_strips_problem() -> None:
    assert split_unit_code("#3793 - cancel") == ("#3793", "cancel")
    assert split_unit_code("№1359 - cancelled") == ("№1359", "cancelled")
    assert split_unit_code("#5713-2") == ("#5713-2", "")
    assert split_unit_code("№2248") == ("№2248", "")
    assert split_unit_code("№1847 - cancelled") == ("№1847", "cancelled")


def test_bos_lh_supplier_shows_only_reference() -> None:
    shown = format_supplier_display(
        "BLUE OCEAN SOLUTIONS",
        "PROPONENT",
        "№1359 - cancelled",
    )
    assert shown == "BLUE OCEAN SOLUTIONS · №1359"
    assert "cancelled" not in shown

    shown_lh = format_supplier_display("LUFTHANSA", "AVIALL", "#3793 - cancel")
    assert shown_lh == "LUFTHANSA · #3793"
    assert "cancel" not in shown_lh


def test_extract_problem_from_comment_and_unit_and_other_column() -> None:
    prev = taz_row()
    curr = taz_row(
        **{
            COL_COMMENT: "supplier cancelled, waiting for quote",
            COL_UNIQUE_UNIT: "№1359 - cancelled",
            COL_SUPPLIER: "BLUE OCEAN SOLUTIONS",
            "Заметка склада": "AOG delay, no stock",
        }
    )
    notes = extract_problem_notes(prev, curr)
    joined = " | ".join(notes).lower()
    assert "waiting for quote" in joined
    assert "cancelled" in joined
    assert "aog delay" in joined
    assert "widget" not in joined


def test_description_is_not_treated_as_problem() -> None:
    row = taz_row(**{COL_DESC: "WINDSHIELD LH", COL_COMMENT: "SAMPLE"})
    notes = extract_problem_notes(row, row)
    assert notes == []


def _events(prev: pd.Series, curr: pd.Series):
    return compare_snapshots({"k": prev}, {"k": curr}, period_days=7)


def test_new_trouble_has_zero_margin_and_no_days_or_delta_term() -> None:
    prev = taz_row(**{COL_STATUS: STATUS_PAID, COL_PURCHASE: 500})
    curr = taz_row(**{COL_STATUS: STATUS_TROUBLE, COL_PURCHASE: 900, COL_COMMENT: "нет в наличии"})
    trouble, *_ = _events(prev, curr)
    assert len(trouble) == 1
    ev = trouble[0]
    assert ev.change_kind == "entered"
    assert fmt_margin_delta_cell(ev) == ("0", "")
    html = render_trouble_table(trouble, section="entered")
    assert "Δ срок" not in html
    assert "Счёт" not in html
    assert "В TROUBLE" not in html
    assert "Δ маржа" not in html
    assert "<th>P/N</th>" in html
    assert "<th>Описание</th>" in html
    assert "нет в наличии" in html
    assert "WIDGET LH" in html


def test_unresolved_keeps_days_in_trouble_column() -> None:
    prev = taz_row()
    curr = taz_row(**{COL_UNIQUE_UNIT: "#3793 - cancel", COL_SUPPLIER: "LUFTHANSA"})
    trouble, *_ = _events(prev, curr)
    assert trouble[0].change_kind == "ongoing"
    html = render_trouble_table(trouble, section="ongoing")
    assert "В TROUBLE" in html
    assert "Δ срок" not in html
    assert "Δ маржа" not in html
    assert "Счёт" not in html
    assert "cancel" in html
    assert "LUFTHANSA · #3793" in html
    assert "≥7 дн." in html


def test_resolved_shows_margin_not_days() -> None:
    prev = taz_row(**{COL_STATUS: STATUS_TROUBLE, COL_PURCHASE: 800, COL_SUPPLIER: "JET TECHNIC"})
    curr = taz_row(
        **{
            COL_STATUS: STATUS_FINISHED,
            COL_PURCHASE: 500,
            COL_SUPPLIER: "IBERIA",
            COL_ROOT_SUPPLIER: "LTB AEROSPACE",
        }
    )
    trouble, *_ = _events(prev, curr)
    assert trouble[0].change_kind == "resolved"
    txt, css = fmt_margin_delta_cell(trouble[0])
    assert css == "pos"
    assert txt != "0"
    html = render_trouble_table(trouble, section="resolved")
    assert "Δ маржа" in html
    assert "В TROUBLE" not in html
    assert "Δ срок" not in html
    assert "Счёт" not in html


def test_full_report_hides_margin_pill_for_new_and_unresolved() -> None:
    prev_rows = {
        "new": taz_row(
            **{COL_STATUS: STATUS_PAID, COL_PN: "NEW-1", COL_PURCHASE: 100, COL_SALE: 5000}
        ),
        "open": taz_row(**{COL_PN: "OPEN-1", COL_SALE: 2000}),
        "done": taz_row(
            **{
                COL_STATUS: STATUS_TROUBLE,
                COL_PN: "DONE-1",
                COL_PURCHASE: 800,
                COL_SALE: 1500,
                COL_SUPPLIER: "JET TECHNIC",
            }
        ),
    }
    curr_rows = {
        "new": taz_row(
            **{
                COL_STATUS: STATUS_TROUBLE,
                COL_PN: "NEW-1",
                COL_PURCHASE: 400,
                COL_SALE: 5000,
                COL_COMMENT: "срыв поставки",
            }
        ),
        "open": taz_row(**{COL_PN: "OPEN-1", COL_SALE: 2000}),
        "done": taz_row(
            **{
                COL_STATUS: STATUS_FINISHED,
                COL_PN: "DONE-1",
                COL_PURCHASE: 500,
                COL_SALE: 1500,
                COL_SUPPLIER: "IBERIA",
                COL_ROOT_SUPPLIER: "LTB",
            }
        ),
    }
    trouble, cancellations, refunds, warranty = compare_snapshots(prev_rows, curr_rows, 7)
    html = render_html_report(
        date(2026, 8, 21),
        date(2026, 8, 28),
        trouble,
        cancellations,
        refunds,
        warranty,
    )
    start = html.index('group-name">Новые TROUBLE')
    mid = html.index('group-name">Нерешённые TROUBLE')
    new_block = html[start:mid]
    assert "Δ маржа" not in new_block
    assert "срыв поставки" in html
    assert "Δ срок" not in html
    assert ">Счёт<" not in html and "th>Счёт" not in html
    assert "В TROUBLE" in html  # unresolved section
    resolved_block = html.split("Решённые TROUBLE", 1)[1].split("Отмены и возвраты", 1)[0]
    assert "Δ маржа" in resolved_block


if __name__ == "__main__":
    tests = [v for k, v in globals().items() if k.startswith("test_")]
    failed = 0
    for fn in tests:
        try:
            fn()
            print(f"ok  {fn.__name__}")
        except Exception as exc:
            failed += 1
            print(f"FAIL {fn.__name__}: {exc}")
    if failed:
        raise SystemExit(1)
    print(f"{len(tests)} tests passed")
