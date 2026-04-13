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

def test_adversarial_all_columns_all_rows(make_sample_workbook):
    """All 4 calculated columns must be correct for every row independently."""
    rows = [
        {"花费": 113.6, "展现量": 1000, "点击量": 50,  "留资人数": 5,  "留资成本": 22.72, "互动成本": 10},
        {"花费": 340.8, "展现量": 3000, "点击量": 60,  "留资人数": 12, "留资成本": 28.40, "互动成本": 20},
    ]
    wb = make_sample_workbook(rows)
    ws = wb.worksheets[3]
    apply_calculated_columns(ws)
    sjhf = find_column(ws, "实际花费")
    djl  = find_column(ws, "点击率")
    cpc  = find_column(ws, "CPC")
    sjcb = find_column(ws, "实际成本")
    # Row 2
    assert abs(ws.cell(2, sjhf).value - 100.0) < 0.01    # 113.6/1.136
    assert abs(ws.cell(2, djl).value  - 0.05)  < 0.001   # 50/1000
    assert abs(ws.cell(2, cpc).value  - 2.0)   < 0.01    # 100/50
    assert abs(ws.cell(2, sjcb).value - 20.0)  < 0.01    # 100/5
    # Row 3: different data, verifies row independence
    assert abs(ws.cell(3, sjhf).value - 300.0) < 0.01    # 340.8/1.136
    assert abs(ws.cell(3, djl).value  - 0.02)  < 0.001   # 60/3000
    assert abs(ws.cell(3, cpc).value  - 5.0)   < 0.01    # 300/60
    assert abs(ws.cell(3, sjcb).value - 25.0)  < 0.01    # 300/12

def test_adversarial_zero_clicks(make_sample_workbook):
    wb = make_sample_workbook([{"花费": 100, "展现量": 1000, "点击量": 0, "留资人数": 5, "留资成本": 20, "互动成本": 10}])
    ws = wb.worksheets[3]
    apply_calculated_columns(ws)
    ctr_idx = find_column(ws, "点击率")
    cpc_idx = find_column(ws, "CPC")
    assert ws.cell(2, ctr_idx).value == 0   # 0/1000 = 0
    assert ws.cell(2, cpc_idx).value == 0   # division by zero → 0

def test_adversarial_column_insertion_order(make_sample_workbook):
    """Each new column must be inserted immediately after its designated source column."""
    wb = make_sample_workbook([{"花费": 113.6, "展现量": 1000, "点击量": 50,
                                "留资人数": 5, "留资成本": 22.72, "互动成本": 10}])
    ws = wb.worksheets[3]
    apply_calculated_columns(ws)
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    hf_pos   = headers.index("花费")
    sjhf_pos = headers.index("实际花费")
    dj_pos   = headers.index("点击量")
    djl_pos  = headers.index("点击率")
    cpc_pos  = headers.index("CPC")
    lzcb_pos = headers.index("留资成本")
    sjcb_pos = headers.index("实际成本")
    assert sjhf_pos == hf_pos   + 1, "实际花费 must be immediately after 花费"
    assert djl_pos  == dj_pos   + 1, "点击率 must be immediately after 点击量"
    assert cpc_pos  == djl_pos  + 1, "CPC must be immediately after 点击率"
    assert sjcb_pos == lzcb_pos + 1, "实际成本 must be immediately after 留资成本"
