from auto_excel.processor import apply_calculated_columns, find_column

def test_all_four_columns_inserted(make_sample_workbook):
    wb = make_sample_workbook([{"花费": 113.6, "展现量": 1000, "点击量": 50, "留资人数": 5, "留资成本": 22.72, "互动成本": 10}])
    ws = wb.worksheets[3]
    apply_calculated_columns(ws)
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    assert "实际花费" in headers
    assert "点击率" in headers
    assert "CPC" in headers
    assert "实际成本" in headers

def test_calculated_values_are_correct(make_sample_workbook):
    wb = make_sample_workbook([{"花费": 113.6, "展现量": 1000, "点击量": 50, "留资人数": 5, "留资成本": 22.72, "互动成本": 10}])
    ws = wb.worksheets[3]
    apply_calculated_columns(ws)
    cost_idx = find_column(ws, "实际花费")
    ctr_idx = find_column(ws, "点击率")
    cpc_idx = find_column(ws, "CPC")
    actual_cost_idx = find_column(ws, "实际成本")
    assert abs(ws.cell(2, cost_idx).value - 100.0) < 0.01        # 113.6/1.136
    assert abs(ws.cell(2, ctr_idx).value - 0.05) < 0.001         # 50/1000
    assert abs(ws.cell(2, cpc_idx).value - 2.0) < 0.01           # 100/50
    assert abs(ws.cell(2, actual_cost_idx).value - 20.0) < 0.01  # 100/5

def test_adversarial_multiple_rows(make_sample_workbook):
    rows = [
        {"花费": 113.6, "展现量": 1000, "点击量": 50, "留资人数": 5, "留资成本": 22.72, "互动成本": 10},
        {"花费": 227.2, "展现量": 2000, "点击量": 100, "留资人数": 10, "留资成本": 22.72, "互动成本": 20},
    ]
    wb = make_sample_workbook(rows)
    ws = wb.worksheets[3]
    apply_calculated_columns(ws)
    cost_idx = find_column(ws, "实际花费")
    assert abs(ws.cell(2, cost_idx).value - 100.0) < 0.01
    assert abs(ws.cell(3, cost_idx).value - 200.0) < 0.01

def test_adversarial_zero_clicks(make_sample_workbook):
    wb = make_sample_workbook([{"花费": 100, "展现量": 1000, "点击量": 0, "留资人数": 5, "留资成本": 20, "互动成本": 10}])
    ws = wb.worksheets[3]
    apply_calculated_columns(ws)
    ctr_idx = find_column(ws, "点击率")
    cpc_idx = find_column(ws, "CPC")
    assert ws.cell(2, ctr_idx).value == 0   # 0/1000 = 0
    assert ws.cell(2, cpc_idx).value == 0   # division by zero → 0
