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
    compare_full_period,
    compare_snapshots,
    duration_in_trouble,
    extract_problem_notes,
    format_supplier_display,
    fmt_hung_days,
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


def test_po_numbers_are_not_problem_text() -> None:
    assert extract_problem_notes(taz_row(**{COL_COMMENT: "PO108630"}), taz_row(**{COL_COMMENT: "PO108630"})) == []
    assert extract_problem_notes(taz_row(**{COL_COMMENT: "P3761426"}), taz_row(**{COL_COMMENT: "P3761426"})) == []
    notes = extract_problem_notes(
        taz_row(**{COL_COMMENT: "PO 107844, stock out, клиент проинформирован"}),
        taz_row(**{COL_COMMENT: "PO 107844, stock out, клиент проинформирован"}),
    )
    assert notes and "stock out" in notes[0].lower()


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
    assert "маржа было" in html
    assert "маржа стало" in html
    assert "Δ маржа" not in html
    assert "В TROUBLE" in html
    assert "Δ срок" not in html
    assert "Счёт" not in html


def test_section_headers_show_plan_or_delta_margin() -> None:
    prev_rows = {
        "new": taz_row(
            **{COL_STATUS: STATUS_PAID, COL_PN: "NEW-1", COL_PURCHASE: 100, COL_SALE: 5000}
        ),
        "open": taz_row(**{COL_PN: "OPEN-1", COL_SALE: 2000, COL_PURCHASE: 1200}),
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
        "open": taz_row(**{COL_PN: "OPEN-1", COL_SALE: 2000, COL_PURCHASE: 1200}),
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
    assert "продажная итого 5 000 USD" in new_block
    assert "маржа" in new_block
    assert "1 кли." in new_block
    assert "срыв поставки" in html
    assert "Δ срок" not in html
    assert "В TROUBLE" in html
    resolved_block = html.split("Решённые TROUBLE", 1)[1].split("Отмены и возвраты", 1)[0]
    assert "маржа было" in resolved_block
    assert "маржа стало" in resolved_block
    assert "Δ маржа" not in resolved_block
    from generate_taz_changes_report import fmt_money, sum_margin_prev, sum_plan_margin

    resolved = [e for e in trouble if e.change_kind == "resolved"]
    was = fmt_money(sum_margin_prev(resolved), 0)
    became = fmt_money(sum_plan_margin(resolved), 0)
    assert f"маржа было {was} USD" in resolved_block
    assert f"маржа стало {became} USD" in resolved_block
    assert "маржа было (решённые)" in html
    assert 'class="period"' in html
    assert "font-weight:700" in html


def test_duration_in_trouble_from_history() -> None:
    d1, d2, d3 = date(2026, 8, 14), date(2026, 8, 21), date(2026, 8, 28)
    row_t = taz_row()
    row_p = taz_row(**{COL_STATUS: STATUS_PAID})
    min_d, max_d, first, last_t, closed = duration_in_trouble(
        "k",
        [(d1, {"k": row_t}), (d2, {"k": row_t}), (d3, {"k": row_t})],
    )
    assert (min_d, max_d, first, last_t, closed) == (14, None, d1, d3, None)

    min_d, max_d, first, last_t, closed = duration_in_trouble(
        "k",
        [(d1, {"k": row_p}), (d2, {"k": row_t}), (d3, {"k": row_t})],
    )
    assert min_d == 7
    assert max_d == 14
    assert first == d2
    assert last_t == d3
    assert closed is None


def test_weekly_ongoing_uses_history_days() -> None:
    d1, d2, d3 = date(2026, 8, 14), date(2026, 8, 21), date(2026, 8, 28)
    row_t = taz_row()
    hist = [
        (d1, {"k": row_t}),
        (d2, {"k": row_t}),
        (d3, {"k": row_t}),
    ]
    trouble, *_ = compare_snapshots(hist[1][1], hist[2][1], 7, history=hist)
    assert trouble[0].change_kind == "ongoing"
    assert trouble[0].trouble_min_days == 14
    html = render_trouble_table(trouble, section="ongoing")
    assert "≥14 дн." in html


def test_resolved_duration_from_history() -> None:
    d1, d2, d3 = date(2026, 8, 14), date(2026, 8, 21), date(2026, 8, 28)
    row_t = taz_row()
    row_done = taz_row(**{COL_STATUS: STATUS_FINISHED, COL_PURCHASE: 500})
    hist = [
        (d1, {"k": row_t}),
        (d2, {"k": row_t}),
        (d3, {"k": row_done}),
    ]
    min_d, max_d, first, last_t, closed = duration_in_trouble("k", hist)
    assert first == d1
    assert last_t == d2
    assert closed == d3
    assert min_d == 7
    assert max_d is None
    trouble, *_ = compare_snapshots(hist[1][1], hist[2][1], 7, history=hist)
    ev = trouble[0]
    assert ev.change_kind == "resolved"
    assert ev.trouble_min_days == 7
    html = render_trouble_table(trouble, section="resolved")
    assert "В TROUBLE" in html
    assert "≥7 дн." in html
    assert "висела с 14.08 (или раньше) по 21.08" in html
    assert "в отчёте 28.08 уже не TROUBLE" in html


def test_resolved_hung_from_first_to_last_trouble() -> None:
    """User example: TROUBLE from ~25.07 through 15.08, gone in 21.08 → ~18–20 days."""
    d_before = date(2026, 7, 17)
    d_first = date(2026, 7, 27)
    d_mid = date(2026, 8, 7)
    d_last = date(2026, 8, 14)
    d_closed = date(2026, 8, 21)
    row_ok = taz_row(**{COL_STATUS: STATUS_PAID})
    row_t = taz_row()
    row_done = taz_row(**{COL_STATUS: STATUS_FINISHED, COL_PURCHASE: 500})
    hist = [
        (d_before, {"k": row_ok}),
        (d_first, {"k": row_t}),
        (d_mid, {"k": row_t}),
        (d_last, {"k": row_t}),
        (d_closed, {"k": row_done}),
    ]
    hung, max_d, first, last_t, closed = duration_in_trouble("k", hist)
    assert first == d_first
    assert last_t == d_last
    assert closed == d_closed
    assert hung == (d_last - d_first).days  # 18
    assert max_d == hung
    assert fmt_hung_days(hung, unknown_start=False) == "~18 дн."

    trouble, *_ = compare_snapshots(hist[-2][1], hist[-1][1], 7, history=hist)
    ev = trouble[0]
    assert ev.change_kind == "resolved"
    assert ev.trouble_min_days == 18
    html = render_trouble_table(trouble, section="resolved")
    assert "~18 дн." in html
    assert "висела с 27.07 по 14.08" in html
    assert "в отчёте 21.08 уже не TROUBLE" in html


def test_resolved_single_snapshot_uses_close_date() -> None:
    d1, d2, d3 = date(2026, 8, 7), date(2026, 8, 14), date(2026, 8, 21)
    row_ok = taz_row(**{COL_STATUS: STATUS_PAID})
    row_t = taz_row()
    row_done = taz_row(**{COL_STATUS: STATUS_FINISHED, COL_PURCHASE: 500})
    hist = [
        (d1, {"k": row_ok}),
        (d2, {"k": row_t}),
        (d3, {"k": row_done}),
    ]
    hung, max_d, first, last_t, closed = duration_in_trouble("k", hist)
    assert (first, last_t, closed) == (d2, d2, d3)
    assert hung == 7
    assert max_d == 7
    html = render_trouble_table(
        compare_snapshots(hist[-2][1], hist[-1][1], 7, history=hist)[0],
        section="resolved",
    )
    assert "~7 дн." in html
    assert "висела с 14.08, в отчёте 21.08 уже не TROUBLE" in html


def test_full_period_classifies_mid_resolved() -> None:
    d1, d2, d3 = date(2026, 7, 17), date(2026, 8, 7), date(2026, 8, 28)
    snaps = [
        (d1, {"open": taz_row(), "mid": taz_row(**{COL_STATUS: STATUS_PAID, COL_PN: "MID"})}),
        (d2, {"open": taz_row(), "mid": taz_row(**{COL_PN: "MID"})}),
        (d3, {
            "open": taz_row(),
            "mid": taz_row(**{COL_STATUS: STATUS_FINISHED, COL_PN: "MID"}),
        }),
    ]
    trouble, *_ = compare_full_period(snaps)
    kinds = {e.pn: e.change_kind for e in trouble}
    assert kinds["PN-1"] == "ongoing"
    assert kinds["MID"] == "resolved"


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
