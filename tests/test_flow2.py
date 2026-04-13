from auto_excel.processor import sort_by_column, find_column

def test_sort_descending(make_sample_workbook):
    rows = [
        {"花费": 50,  "展现量": 100, "点击量": 10, "留资人数": 1, "留资成本": 50,  "互动成本": 5},
        {"花费": 200, "展现量": 100, "点击量": 10, "留资人数": 1, "留资成本": 200, "互动成本": 5},
        {"花费": 100, "展现量": 100, "点击量": 10, "留资人数": 1, "留资成本": 100, "互动成本": 5},
    ]
    wb = make_sample_workbook(rows)
    ws = wb.worksheets[3]
    sort_by_column(ws, "花费", descending=True)
    cost_idx = find_column(ws, "花费")
    assert ws.cell(2, cost_idx).value == 200
    assert ws.cell(3, cost_idx).value == 100
    assert ws.cell(4, cost_idx).value == 50

def test_sort_preserves_header(make_sample_workbook):
    rows = [{"花费": 999, "展现量": 100, "点击量": 10, "留资人数": 1, "留资成本": 50, "互动成本": 5}]
    wb = make_sample_workbook(rows)
    ws = wb.worksheets[3]
    sort_by_column(ws, "花费")
    assert ws.cell(1, find_column(ws, "花费")).value == "花费"  # header unchanged

def test_adversarial_sort_moves_entire_row(make_sample_workbook):
    rows = [
        {"花费": 50,  "展现量": 111, "点击量": 10, "留资人数": 1, "留资成本": 50,  "互动成本": 5},
        {"花费": 200, "展现量": 222, "点击量": 10, "留资人数": 1, "留资成本": 200, "互动成本": 5},
    ]
    wb = make_sample_workbook(rows)
    ws = wb.worksheets[3]
    sort_by_column(ws, "花费", descending=True)
    cost_idx = find_column(ws, "花费")
    imp_idx  = find_column(ws, "展现量")
    assert ws.cell(2, cost_idx).value == 200
    assert ws.cell(2, imp_idx).value  == 222  # entire row moved together

def test_adversarial_single_row(make_sample_workbook):
    wb = make_sample_workbook([{"花费": 100, "展现量": 100, "点击量": 10, "留资人数": 1, "留资成本": 100, "互动成本": 5}])
    ws = wb.worksheets[3]
    sort_by_column(ws, "花费")  # must not raise with 1 row
    assert ws.cell(2, find_column(ws, "花费")).value == 100


def test_adversarial_sort_ascending(make_sample_workbook):
    """descending=False must produce ascending order (not just ignore the flag)."""
    rows = [
        {"花费": 200, "展现量": 100, "点击量": 10, "留资人数": 1, "留资成本": 200, "互动成本": 5},
        {"花费": 50,  "展现量": 100, "点击量": 10, "留资人数": 1, "留资成本": 50,  "互动成本": 5},
        {"花费": 100, "展现量": 100, "点击量": 10, "留资人数": 1, "留资成本": 100, "互动成本": 5},
    ]
    wb = make_sample_workbook(rows)
    ws = wb.worksheets[3]
    sort_by_column(ws, "花费", descending=False)
    cost_idx = find_column(ws, "花费")
    assert ws.cell(2, cost_idx).value == 50
    assert ws.cell(3, cost_idx).value == 100
    assert ws.cell(4, cost_idx).value == 200


def test_adversarial_none_sorts_to_end(make_sample_workbook):
    """None values must sort to the end in descending order, and move with their row."""
    rows = [
        {"花费": None, "展现量": 111, "点击量": 10, "留资人数": 1, "留资成本": 0,   "互动成本": 5},
        {"花费": 100,  "展现量": 222, "点击量": 10, "留资人数": 1, "留资成本": 100, "互动成本": 5},
        {"花费": 50,   "展现量": 333, "点击量": 10, "留资人数": 1, "留资成本": 50,  "互动成本": 5},
    ]
    wb = make_sample_workbook(rows)
    ws = wb.worksheets[3]
    sort_by_column(ws, "花费", descending=True)
    cost_idx = find_column(ws, "花费")
    imp_idx  = find_column(ws, "展现量")
    assert ws.cell(2, cost_idx).value == 100
    assert ws.cell(3, cost_idx).value == 50
    assert ws.cell(4, cost_idx).value is None
    assert ws.cell(4, imp_idx).value == 111  # entire row moved with None value


def test_adversarial_invalid_column_raises(make_sample_workbook):
    """Sorting by a non-existent column must raise ValueError."""
    import pytest
    wb = make_sample_workbook([{"花费": 100, "展现量": 100, "点击量": 10, "留资人数": 1, "留资成本": 100, "互动成本": 5}])
    ws = wb.worksheets[3]
    with pytest.raises(ValueError):
        sort_by_column(ws, "不存在的列")
