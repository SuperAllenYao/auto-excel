"""Tests for the formula workbook fixture and formula resolution infrastructure."""

from conftest import FORMULA_SHEET4_HEADERS, SOURCE_SHEET_HEADERS
from auto_excel.processor import resolve_formulas


# ---------------------------------------------------------------------------
# SUMIFS resolution tests
# ---------------------------------------------------------------------------


def test_resolve_sumifs_basic(make_formula_workbook):
    """SUMIFS formulas are resolved to aggregated values from source sheet."""
    wb = make_formula_workbook(
        rows=[{"笔记标题": "Note A", "笔记ID": "id1"}, {"笔记标题": "Note B", "笔记ID": "id2"}],
        source_rows=[
            {"笔记ID": "id1", "消费": 50.0,  "展现量": 300, "点击量": 10, "留资人数": 1},
            {"笔记ID": "id1", "消费": 30.0,  "展现量": 200, "点击量": 5,  "留资人数": 2},
            {"笔记ID": "id2", "消费": 100.0, "展现量": 800, "点击量": 40, "留资人数": 3},
        ],
    )
    ws = wb.worksheets[3]
    resolve_formulas(wb, ws)

    # id1: 消费 sum = 80, 展现量 sum = 500, 点击量 sum = 15, 留资人数 sum = 3
    assert ws.cell(2, 3).value == 80.0
    assert ws.cell(2, 4).value == 500
    assert ws.cell(2, 5).value == 15
    assert ws.cell(2, 6).value == 3

    # id2 sums
    assert ws.cell(3, 3).value == 100.0
    assert ws.cell(3, 4).value == 800


def test_resolve_sumifs_missing_id(make_formula_workbook):
    wb = make_formula_workbook(
        rows=[{"笔记标题": "Missing", "笔记ID": "id_missing"}],
        source_rows=[{"笔记ID": "id_other", "消费": 100.0, "展现量": 500, "点击量": 20, "留资人数": 1}],
    )
    ws = wb.worksheets[3]
    resolve_formulas(wb, ws)
    assert ws.cell(2, 3).value == 0


def test_resolve_sumifs_no_formulas(make_sample_workbook):
    wb = make_sample_workbook([{"花费": 100.0, "展现量": 1000, "点击量": 50, "留资人数": 5}])
    ws = wb.worksheets[3]
    resolve_formulas(wb, ws)
    assert ws.cell(2, 3).value == 100.0  # Value unchanged


def test_resolve_sumifs_multiple_source_rows(make_formula_workbook):
    wb = make_formula_workbook(
        rows=[{"笔记标题": "Multi", "笔记ID": "id1"}],
        source_rows=[
            {"笔记ID": "id1", "消费": 10.0, "展现量": 100, "点击量": 5, "留资人数": 0},
            {"笔记ID": "id1", "消费": 20.0, "展现量": 200, "点击量": 3, "留资人数": 1},
            {"笔记ID": "id1", "消费": 30.0, "展现量": 300, "点击量": 7, "留资人数": 0},
        ],
    )
    ws = wb.worksheets[3]
    resolve_formulas(wb, ws)
    assert ws.cell(2, 3).value == 60.0
    assert ws.cell(2, 4).value == 600
    assert ws.cell(2, 5).value == 15
    assert ws.cell(2, 6).value == 1


# ---------------------------------------------------------------------------
# Smoke test: fixture creates correct structure
# ---------------------------------------------------------------------------


def test_fixture_creates_formula_workbook(make_formula_workbook):
    wb = make_formula_workbook(
        rows=[{"笔记标题": "Test", "笔记ID": "id1"}],
        source_rows=[
            {
                "笔记ID": "id1",
                "日期": "2026-01-01",
                "消费": 100.0,
                "展现量": 500,
                "点击量": 20,
                "留资人数": 2,
            }
        ],
    )
    ws = wb.worksheets[3]
    assert ws.cell(1, 3).value == "花费"
    assert isinstance(ws.cell(2, 3).value, str)
    assert ws.cell(2, 3).value.startswith("=SUMIFS")


# ---------------------------------------------------------------------------
# Adversarial tests
# ---------------------------------------------------------------------------


# Category 1: Wrong data — missing 笔记ID key should not crash; formula
# row reference must still be correct.
def test_fixture_missing_notebook_id_defaults_to_empty(make_formula_workbook):
    """If a row dict omits 笔记ID, col B should be empty string (not raise)."""
    wb = make_formula_workbook(
        rows=[{"笔记标题": "Only title"}],
        source_rows=[],
    )
    ws = wb.worksheets[3]
    assert ws.cell(2, 2).value == ""


# Category 2: Structural integrity — workbook must have 5 sheets in the right
# order and with the right titles.
def test_fixture_sheet_structure(make_formula_workbook):
    """Formula sheet must be at index 3 with title '笔记id'; source at index 4 '源数据'."""
    wb = make_formula_workbook(rows=[], source_rows=[])
    assert len(wb.worksheets) == 5
    assert wb.worksheets[3].title == "笔记id"
    assert wb.worksheets[4].title == "源数据"


# Category 3: Formula string correctness — each formula column should
# reference the correct source column and the right row number.
def test_fixture_formula_strings_reference_correct_columns(make_formula_workbook):
    """All eight columns of row 2 must have the values/formulas described in the spec."""
    wb = make_formula_workbook(
        rows=[{"笔记标题": "Note A", "笔记ID": "abc"}],
        source_rows=[],
    )
    ws = wb.worksheets[3]

    # Literal columns
    assert ws.cell(2, 1).value == "Note A"
    assert ws.cell(2, 2).value == "abc"

    # SUMIFS formulas — verify both source column letter and row reference
    assert ws.cell(2, 3).value == "=SUMIFS('源数据'!C:C,'源数据'!A:A,B2)"  # 花费
    assert ws.cell(2, 4).value == "=SUMIFS('源数据'!D:D,'源数据'!A:A,B2)"  # 展现量
    assert ws.cell(2, 5).value == "=SUMIFS('源数据'!E:E,'源数据'!A:A,B2)"  # 点击量
    assert ws.cell(2, 6).value == "=SUMIFS('源数据'!F:F,'源数据'!A:A,B2)"  # 留资人数

    # Division formulas
    assert ws.cell(2, 7).value == "=C2/F2"   # 留资成本
    assert ws.cell(2, 8).value == "=C2/E2"   # 互动成本


# Category 4: Multiple rows — row index in formulas must advance correctly.
def test_fixture_multi_row_formula_row_references(make_formula_workbook):
    """Row 3 formulas must reference row 3, not row 2."""
    wb = make_formula_workbook(
        rows=[
            {"笔记标题": "A", "笔记ID": "id1"},
            {"笔记标题": "B", "笔记ID": "id2"},
        ],
        source_rows=[],
    )
    ws = wb.worksheets[3]
    assert ws.cell(3, 3).value == "=SUMIFS('源数据'!C:C,'源数据'!A:A,B3)"
    assert ws.cell(3, 7).value == "=C3/F3"
    assert ws.cell(3, 8).value == "=C3/E3"


# Category 5: Source sheet — headers and data rows must be written correctly.
def test_fixture_source_sheet_headers_and_data(make_formula_workbook):
    """Source sheet must have correct headers and data rows."""
    wb = make_formula_workbook(
        rows=[],
        source_rows=[
            {
                "笔记ID": "note01",
                "日期": "2026-03-01",
                "消费": 250.5,
                "展现量": 1000,
                "点击量": 50,
                "留资人数": 5,
            }
        ],
    )
    ws_src = wb.worksheets[4]

    # Header row
    for col_idx, expected_header in enumerate(SOURCE_SHEET_HEADERS, start=1):
        assert ws_src.cell(1, col_idx).value == expected_header

    # Data row — assert every column so swapped writes are caught.
    assert ws_src.cell(2, 1).value == "note01"
    assert ws_src.cell(2, 2).value == "2026-03-01"
    assert ws_src.cell(2, 3).value == 250.5
    assert ws_src.cell(2, 4).value == 1000
    assert ws_src.cell(2, 5).value == 50
    assert ws_src.cell(2, 6).value == 5


# Category 6: Header row of formula sheet must match FORMULA_SHEET4_HEADERS exactly.
def test_fixture_formula_sheet_headers(make_formula_workbook):
    """Formula sheet row 1 must exactly match FORMULA_SHEET4_HEADERS."""
    wb = make_formula_workbook(rows=[], source_rows=[])
    ws = wb.worksheets[3]
    # Read 9 cols (one past expected end) to also verify there are no extra headers.
    actual_headers = [ws.cell(1, col).value for col in range(1, 10)]
    assert actual_headers[:8] == FORMULA_SHEET4_HEADERS
    assert actual_headers[8] is None  # no stray header at col 9


# Category 7: Constant integrity — guard against the constant itself being truncated.
def test_formula_sheet4_headers_constant_integrity():
    """FORMULA_SHEET4_HEADERS must have exactly 8 items with known first and last."""
    assert len(FORMULA_SHEET4_HEADERS) == 8
    assert FORMULA_SHEET4_HEADERS[0] == "笔记标题"
    assert FORMULA_SHEET4_HEADERS[7] == "互动成本"


# Category 8: Numeric default behavior — missing numeric keys in source_rows default to 0.
def test_fixture_source_numeric_default_is_zero(make_formula_workbook):
    """Missing numeric keys in source_rows default to 0 so float() is safe downstream."""
    wb = make_formula_workbook(
        rows=[],
        source_rows=[{"笔记ID": "id9", "日期": "2026-01-01"}],  # numeric cols omitted
    )
    ws_src = wb.worksheets[4]
    assert ws_src.cell(2, 3).value == 0  # 消费
    assert ws_src.cell(2, 4).value == 0  # 展现量
    assert ws_src.cell(2, 5).value == 0  # 点击量
    assert ws_src.cell(2, 6).value == 0  # 留资人数
