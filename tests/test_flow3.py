from auto_excel.processor import apply_calculated_columns, sort_by_column, group_and_merge, find_column

# Helper: 5-row workbook where 实际成本 will be: 100, 100, ~66.7, 25, 20
ROWS_5 = [
    {"花费": 113.6, "展现量": 1000, "点击量": 50, "留资人数": 1,   "留资成本": 100, "互动成本": 10},
    {"花费": 113.6, "展现量": 1000, "点击量": 50, "留资人数": 1,   "留资成本": 100, "互动成本": 10},
    {"花费": 113.6, "展现量": 1000, "点击量": 50, "留资人数": 1.5, "留资成本": 67,  "互动成本": 10},
    {"花费": 113.6, "展现量": 1000, "点击量": 50, "留资人数": 4,   "留资成本": 25,  "互动成本": 10},
    {"花费": 113.6, "展现量": 1000, "点击量": 50, "留资人数": 5,   "留资成本": 20,  "互动成本": 10},
]

def _apply_all(ws):
    apply_calculated_columns(ws)
    sort_by_column(ws, "实际成本", descending=True)
    group_and_merge(ws)

def test_group_inserts_zhanbi_column(make_sample_workbook):
    wb = make_sample_workbook(ROWS_5)
    ws = wb.worksheets[3]
    _apply_all(ws)
    assert find_column(ws, "占比") > 0

def test_group_merge_content(make_sample_workbook):
    wb = make_sample_workbook(ROWS_5)
    ws = wb.worksheets[3]
    _apply_all(ws)
    zb_col = find_column(ws, "占比")
    # After sort desc: rows 2-3 are ≥90, row 4 is 50-89, rows 5-6 are <50
    assert ws.cell(2, zb_col).value == "2/40%"
    assert ws.cell(4, zb_col).value == "1/20%"
    assert ws.cell(5, zb_col).value == "2/40%"

def test_adversarial_all_in_one_range(make_sample_workbook):
    """All rows ≥ 90 → single group '5/100%', no other groups."""
    rows = [
        {"花费": 113.6, "展现量": 1000, "点击量": 50, "留资人数": 1, "留资成本": 100, "互动成本": 10},
    ] * 5
    wb = make_sample_workbook(rows)
    ws = wb.worksheets[3]
    _apply_all(ws)
    zb_col = find_column(ws, "占比")
    assert ws.cell(2, zb_col).value == "5/100%"
    # Rows 3-6 are merged cells — their value should be None (not "5/100%")
    assert ws.cell(3, zb_col).value is None

def test_adversarial_empty_range(make_sample_workbook):
    """No medium range rows → only high and low groups appear."""
    rows = [
        {"花费": 113.6, "展现量": 1000, "点击量": 50, "留资人数": 1, "留资成本": 100, "互动成本": 10},
        {"花费": 113.6, "展现量": 1000, "点击量": 50, "留资人数": 1, "留资成本": 100, "互动成本": 10},
        {"花费": 113.6, "展现量": 1000, "点击量": 50, "留资人数": 5, "留资成本": 20,  "互动成本": 10},
        {"花费": 113.6, "展现量": 1000, "点击量": 50, "留资人数": 4, "留资成本": 20,  "互动成本": 10},
    ]
    wb = make_sample_workbook(rows)
    ws = wb.worksheets[3]
    _apply_all(ws)
    zb_col = find_column(ws, "占比")
    # Rows 2-3: high (≥90), 2 rows, 50%
    assert ws.cell(2, zb_col).value == "2/50%"
    # Rows 4-5: low (<50), 2 rows, 50%
    assert ws.cell(4, zb_col).value == "2/50%"

def test_adversarial_single_row(make_sample_workbook):
    """Single data row → '1/100%', no merge attempted."""
    wb = make_sample_workbook([
        {"花费": 113.6, "展现量": 1000, "点击量": 50, "留资人数": 1, "留资成本": 100, "互动成本": 10}
    ])
    ws = wb.worksheets[3]
    _apply_all(ws)
    zb_col = find_column(ws, "占比")
    assert ws.cell(2, zb_col).value == "1/100%"
