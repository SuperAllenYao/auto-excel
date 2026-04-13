def test_sample_workbook_has_4_sheets(make_sample_workbook):
    wb = make_sample_workbook([{"花费": 113.6, "展现量": 1000, "点击量": 50, "留资人数": 5, "留资成本": 22.72, "互动成本": 10}])
    assert len(wb.sheetnames) == 4

def test_sheet4_has_correct_headers(make_sample_workbook):
    wb = make_sample_workbook([{"花费": 100}])
    ws = wb.worksheets[3]
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    for name in ["花费", "展现量", "点击量", "留资人数", "留资成本", "互动成本"]:
        assert name in headers

def test_sheet4_data_rows(make_sample_workbook):
    rows = [{"花费": 100}, {"花费": 200}]
    wb = make_sample_workbook(rows)
    ws = wb.worksheets[3]
    assert ws.max_row == 3  # 1 header + 2 data

def test_adversarial_empty_rows(make_sample_workbook):
    wb = make_sample_workbook([])
    ws = wb.worksheets[3]
    assert ws.max_row == 1  # header only

def test_adversarial_missing_fields_get_defaults(make_sample_workbook):
    wb = make_sample_workbook([{"花费": 100}])  # other fields missing
    ws = wb.worksheets[3]
    assert ws.max_row == 2  # should still create the row

def test_adversarial_header_order(make_sample_workbook):
    """Headers must be in exact specified order."""
    wb = make_sample_workbook([])
    ws = wb.worksheets[3]
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    expected = ["日期", "计划名称", "花费", "展现量", "点击量", "留资人数", "留资成本", "互动成本"]
    assert headers == expected

def test_adversarial_tmp_dirs_all_exist(tmp_dirs):
    """All subdirectories in tmp_dirs must exist."""
    assert tmp_dirs["raw"].exists()
    assert tmp_dirs["new"].exists()
    assert tmp_dirs["log"].exists()
    assert tmp_dirs["raw"].parent == tmp_dirs["base"]
