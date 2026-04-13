import pytest
from auto_excel.processor import find_column, insert_calculated_column

def test_find_column_returns_correct_index(make_sample_workbook):
    wb = make_sample_workbook([{"花费": 100}])
    ws = wb.worksheets[3]
    idx = find_column(ws, "花费")
    assert ws.cell(1, idx).value == "花费"

def test_find_column_raises_on_missing(make_sample_workbook):
    wb = make_sample_workbook([{"花费": 100}])
    ws = wb.worksheets[3]
    with pytest.raises(ValueError):
        find_column(ws, "不存在的列")

def test_insert_column_adds_header_and_data(make_sample_workbook):
    wb = make_sample_workbook([{"花费": 113.6}, {"花费": 227.2}])
    ws = wb.worksheets[3]
    insert_calculated_column(ws, "花费", "实际花费", lambda val, row: val / 1.136)
    new_idx = find_column(ws, "实际花费")
    assert ws.cell(1, new_idx).value == "实际花费"
    assert abs(ws.cell(2, new_idx).value - 100.0) < 0.01
    assert abs(ws.cell(3, new_idx).value - 200.0) < 0.01

def test_adversarial_zero_division_fills_zero(make_sample_workbook):
    wb = make_sample_workbook([{"花费": 0}])
    ws = wb.worksheets[3]
    insert_calculated_column(ws, "花费", "test", lambda val, row: 1 / val)
    new_idx = find_column(ws, "test")
    assert ws.cell(2, new_idx).value == 0

def test_adversarial_insert_shifts_columns(make_sample_workbook):
    wb = make_sample_workbook([{"花费": 100}])
    ws = wb.worksheets[3]
    old_cols = ws.max_column
    insert_calculated_column(ws, "花费", "new_col", lambda val, row: val * 2)
    assert ws.max_column == old_cols + 1

def test_adversarial_find_after_insert(make_sample_workbook):
    wb = make_sample_workbook([{"花费": 100, "点击量": 50}])
    ws = wb.worksheets[3]
    insert_calculated_column(ws, "花费", "实际花费", lambda val, row: val / 1.136)
    idx = find_column(ws, "点击量")
    assert ws.cell(1, idx).value == "点击量"


def test_adversarial_new_column_is_immediately_after_source(make_sample_workbook):
    """Wrong implementation: insert_at = source_idx (BEFORE) instead of source_idx + 1 (AFTER)."""
    wb = make_sample_workbook([{"花费": 100}])
    ws = wb.worksheets[3]
    source_idx_before = find_column(ws, "花费")
    insert_calculated_column(ws, "花费", "新列", lambda val, row: val * 2)
    assert ws.cell(1, source_idx_before + 1).value == "新列"
    assert ws.cell(1, source_idx_before).value == "花费"


def test_adversarial_header_written_at_row1_insert_position(make_sample_workbook):
    """Wrong implementation: inserts column but forgets to write the header."""
    wb = make_sample_workbook([{"花费": 100}])
    ws = wb.worksheets[3]
    source_idx = find_column(ws, "花费")
    insert_calculated_column(ws, "花费", "新列头", lambda val, row: val)
    assert ws.cell(1, source_idx + 1).value == "新列头"


def test_adversarial_row_iteration_starts_at_row2(make_sample_workbook):
    """Wrong implementation: iterates from row 1 (overwrites header) or skips row 2."""
    wb = make_sample_workbook([{"花费": 50.0}, {"花费": 100.0}])
    ws = wb.worksheets[3]
    insert_calculated_column(ws, "花费", "双倍花费", lambda val, row: val * 2)
    new_idx = find_column(ws, "双倍花费")
    assert ws.cell(1, new_idx).value == "双倍花费"
    assert not isinstance(ws.cell(1, new_idx).value, (int, float))
    assert abs(ws.cell(2, new_idx).value - 100.0) < 0.01
    assert abs(ws.cell(3, new_idx).value - 200.0) < 0.01


def test_adversarial_insert_after_non_last_column(make_sample_workbook):
    """Verify AFTER semantics when source column is not the last column."""
    wb = make_sample_workbook([{"花费": 100, "点击量": 50}])
    ws = wb.worksheets[3]
    source_idx = find_column(ws, "花费")
    click_idx_before = find_column(ws, "点击量")
    insert_calculated_column(ws, "花费", "中间列", lambda val, row: val)
    assert ws.cell(1, source_idx + 1).value == "中间列"
    new_click_idx = find_column(ws, "点击量")
    assert new_click_idx == click_idx_before + 1
