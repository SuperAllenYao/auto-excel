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
