"""Tests for the CLI entry point."""
import json
from typer.testing import CliRunner
from auto_excel.cli import app

runner = CliRunner()

_SAMPLE_ROWS = [
    {"花费": 113.6, "展现量": 1000, "点击量": 50, "留资人数": 5, "留资成本": 22.72, "互动成本": 10},
    {"花费": 227.2, "展现量": 2000, "点击量": 100, "留资人数": 10, "留资成本": 22.72, "互动成本": 10},
    {"花费": 340.8, "展现量": 3000, "点击量": 150, "留资人数": 15, "留资成本": 22.72, "互动成本": 10},
]


def test_on_no_files(tmp_dirs, monkeypatch):
    import auto_excel.config as cfg
    monkeypatch.setattr(cfg, "RAW_DIR", tmp_dirs["raw"])
    monkeypatch.setattr(cfg, "NEW_DIR", tmp_dirs["new"])
    monkeypatch.setattr(cfg, "LOG_DIR", tmp_dirs["log"])
    monkeypatch.setattr(cfg, "STATE_FILE", tmp_dirs["log"] / "processed.json")
    import auto_excel.state as st
    monkeypatch.setattr(st, "STATE_FILE", tmp_dirs["log"] / "processed.json")
    result = runner.invoke(app, ["on"])
    assert result.exit_code == 0
    assert "没有" in result.output or "待处理" in result.output


def test_on_processes_file(tmp_dirs, monkeypatch, make_sample_workbook):
    import auto_excel.config as cfg
    state_file = tmp_dirs["log"] / "processed.json"
    monkeypatch.setattr(cfg, "RAW_DIR", tmp_dirs["raw"])
    monkeypatch.setattr(cfg, "NEW_DIR", tmp_dirs["new"])
    monkeypatch.setattr(cfg, "LOG_DIR", tmp_dirs["log"])
    monkeypatch.setattr(cfg, "STATE_FILE", state_file)
    import auto_excel.state as st
    monkeypatch.setattr(st, "STATE_FILE", state_file)
    wb = make_sample_workbook(_SAMPLE_ROWS)
    wb.save(tmp_dirs["raw"] / "test_report.xlsx")
    result = runner.invoke(app, ["on"])
    assert result.exit_code == 0
    assert (tmp_dirs["new"] / "test_report.xlsx").exists()
    state = json.loads(state_file.read_text())
    assert "test_report.xlsx" in state


def test_adversarial_skips_already_processed(tmp_dirs, monkeypatch, make_sample_workbook):
    import auto_excel.config as cfg
    state_file = tmp_dirs["log"] / "processed.json"
    monkeypatch.setattr(cfg, "RAW_DIR", tmp_dirs["raw"])
    monkeypatch.setattr(cfg, "NEW_DIR", tmp_dirs["new"])
    monkeypatch.setattr(cfg, "LOG_DIR", tmp_dirs["log"])
    monkeypatch.setattr(cfg, "STATE_FILE", state_file)
    import auto_excel.state as st
    monkeypatch.setattr(st, "STATE_FILE", state_file)
    state_file.write_text(json.dumps({"already.xlsx": {"processed_at": "2026-01-01T00:00:00", "status": "success"}}))
    wb = make_sample_workbook(_SAMPLE_ROWS[:2])
    wb.save(tmp_dirs["raw"] / "already.xlsx")
    result = runner.invoke(app, ["on"])
    assert result.exit_code == 0
    assert "没有" in result.output or "待处理" in result.output


def test_adversarial_continues_after_error(tmp_dirs, monkeypatch, make_sample_workbook):
    import auto_excel.config as cfg
    state_file = tmp_dirs["log"] / "processed.json"
    monkeypatch.setattr(cfg, "RAW_DIR", tmp_dirs["raw"])
    monkeypatch.setattr(cfg, "NEW_DIR", tmp_dirs["new"])
    monkeypatch.setattr(cfg, "LOG_DIR", tmp_dirs["log"])
    monkeypatch.setattr(cfg, "STATE_FILE", state_file)
    import auto_excel.state as st
    monkeypatch.setattr(st, "STATE_FILE", state_file)
    (tmp_dirs["raw"] / "bad.xlsx").write_bytes(b"not an excel file")
    wb = make_sample_workbook(_SAMPLE_ROWS[:2])
    wb.save(tmp_dirs["raw"] / "good.xlsx")
    result = runner.invoke(app, ["on"])
    assert result.exit_code == 0
    assert (tmp_dirs["new"] / "good.xlsx").exists()
    if state_file.exists():
        state = json.loads(state_file.read_text())
        assert "bad.xlsx" not in state
