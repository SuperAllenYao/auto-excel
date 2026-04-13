"""End-to-end integration test for the full auto-excel pipeline."""
import json
from openpyxl import load_workbook
from typer.testing import CliRunner
from auto_excel.cli import app

runner = CliRunner()

ROWS = [
    {"花费": 100.0, "展现量": 3000,  "点击量": 100, "留资人数": 4},   # 实际成本≈22.0 (低)
    {"花费": 450.0, "展现量": 8000,  "点击量": 400, "留资人数": 4},   # 实际成本≈99.0 (高)
    {"花费": 200.0, "展现量": 5000,  "点击量": 200, "留资人数": 3},   # 实际成本≈58.7 (中)
    {"花费": 150.0, "展现量": 4000,  "点击量": 150, "留资人数": 5},   # 实际成本≈26.4 (低)
    {"花费": 600.0, "展现量": 10000, "点击量": 500, "留资人数": 5},   # 实际成本≈105.6 (高)
]


def test_full_pipeline(tmp_dirs, monkeypatch, make_sample_workbook):
    """Full pipeline: Raw/ → process → New/ with all transformations verified."""
    import auto_excel.config as cfg
    import auto_excel.state as st

    state_file = tmp_dirs["log"] / "processed.json"
    for attr, val in [("RAW_DIR", tmp_dirs["raw"]), ("NEW_DIR", tmp_dirs["new"]),
                      ("LOG_DIR", tmp_dirs["log"]), ("STATE_FILE", state_file)]:
        monkeypatch.setattr(cfg, attr, val)
    monkeypatch.setattr(st, "STATE_FILE", state_file)

    # 1. Save sample workbook to Raw/
    wb = make_sample_workbook(ROWS)
    src = tmp_dirs["raw"] / "integration_test.xlsx"
    wb.save(src)

    # 2. Run CLI
    result = runner.invoke(app, ["on"])
    assert result.exit_code == 0, f"CLI failed: {result.output}"

    # 3. Output file exists in New/
    out_path = tmp_dirs["new"] / "integration_test.xlsx"
    assert out_path.exists()

    # 4. State recorded
    assert state_file.exists()
    state = json.loads(state_file.read_text())
    assert "integration_test.xlsx" in state

    # 5. Open output and check Sheet 4
    wb_out = load_workbook(out_path)
    ws = wb_out.worksheets[3]
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]

    # 6. Verify all 4 calculated columns exist (Flow 1)
    for col in ["实际花费", "点击率", "CPC", "实际成本"]:
        assert col in headers, f"Missing column: {col}"

    # 7. Verify 占比 column exists (Flow 3)
    assert "占比" in headers, "Missing 占比 column"

    # 8. Verify data is sorted descending by 实际成本 (Flow 2)
    cost_col = headers.index("实际成本") + 1
    costs = [ws.cell(r, cost_col).value for r in range(2, ws.max_row + 1)
             if ws.cell(r, cost_col).value is not None]
    assert len(costs) == 5
    assert costs == sorted(costs, reverse=True), f"Not sorted desc: {costs}"

    # 9. Verify 实际花费 values (spot check first row after sort)
    shf_col = headers.index("实际花费") + 1
    first_shf = ws.cell(2, shf_col).value
    # First row after sort has highest cost: 花费=600/1.136 ≈ 528.169
    assert first_shf is not None
    assert abs(float(first_shf) - 600.0 / 1.136) < 0.01, f"实际花费 wrong: {first_shf}"

    # 10. Verify 占比 content with exact values for all 3 groups
    zb_col = headers.index("占比") + 1
    # High group (rows 2-3): merged, top-left cell has value
    assert ws.cell(2, zb_col).value == "2/40%", f"High group 占比 wrong: {ws.cell(2, zb_col).value}"
    # Medium group (row 4): single cell
    assert ws.cell(4, zb_col).value == "1/20%", f"Medium group 占比 wrong: {ws.cell(4, zb_col).value}"
    # Low group (rows 5-6): merged, top-left has value
    assert ws.cell(5, zb_col).value == "2/40%", f"Low group 占比 wrong: {ws.cell(5, zb_col).value}"
