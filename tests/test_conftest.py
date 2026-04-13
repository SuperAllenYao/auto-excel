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

def test_adversarial_missing_fields_default_to_zero(make_sample_workbook):
    """Missing fields must default to 0, not None or any other value."""
    HEADERS = ["日期", "计划名称", "花费", "展现量", "点击量", "留资人数", "留资成本", "互动成本"]
    wb = make_sample_workbook([{"花费": 100}])
    ws = wb.worksheets[3]
    花费_col = HEADERS.index("花费") + 1
    展现量_col = HEADERS.index("展现量") + 1
    assert ws.cell(2, 花费_col).value == 100
    assert ws.cell(2, 展现量_col).value == 0
    assert ws.cell(2, HEADERS.index("留资人数") + 1).value == 0

def test_adversarial_provided_values_written_correctly(make_sample_workbook):
    """Values provided in row_dict must be written to the correct cells."""
    HEADERS = ["日期", "计划名称", "花费", "展现量", "点击量", "留资人数", "留资成本", "互动成本"]
    wb = make_sample_workbook([
        {"花费": 113.6, "展现量": 2000, "留资人数": 5},
        {"花费": 999.9, "展现量": 0},
    ])
    ws = wb.worksheets[3]
    花费_col = HEADERS.index("花费") + 1
    展现量_col = HEADERS.index("展现量") + 1
    留资人数_col = HEADERS.index("留资人数") + 1
    assert ws.cell(2, 花费_col).value == 113.6
    assert ws.cell(2, 展现量_col).value == 2000
    assert ws.cell(2, 留资人数_col).value == 5
    assert ws.cell(3, 花费_col).value == 999.9
    assert ws.cell(3, 展现量_col).value == 0

def test_adversarial_tmp_dirs_are_directories(tmp_dirs):
    """tmp_dirs entries must be directories, not files."""
    assert tmp_dirs["raw"].is_dir()
    assert tmp_dirs["new"].is_dir()
    assert tmp_dirs["log"].is_dir()
