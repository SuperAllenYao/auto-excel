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
