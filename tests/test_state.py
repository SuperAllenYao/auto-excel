from auto_excel.state import load_state, save_file_state, is_processed


def test_load_state_returns_empty_when_no_file(tmp_dirs, monkeypatch):
    monkeypatch.setattr("auto_excel.state.STATE_FILE", tmp_dirs["log"] / "processed.json")
    assert load_state() == {}


def test_save_and_load_roundtrip(tmp_dirs, monkeypatch):
    monkeypatch.setattr("auto_excel.state.STATE_FILE", tmp_dirs["log"] / "processed.json")
    save_file_state("test.xlsx")
    state = load_state()
    assert "test.xlsx" in state
    assert state["test.xlsx"]["status"] == "success"
    assert "processed_at" in state["test.xlsx"]


def test_is_processed(tmp_dirs, monkeypatch):
    monkeypatch.setattr("auto_excel.state.STATE_FILE", tmp_dirs["log"] / "processed.json")
    assert is_processed("test.xlsx") is False
    save_file_state("test.xlsx")
    assert is_processed("test.xlsx") is True


def test_adversarial_save_preserves_existing_entries(tmp_dirs, monkeypatch):
    monkeypatch.setattr("auto_excel.state.STATE_FILE", tmp_dirs["log"] / "processed.json")
    save_file_state("a.xlsx")
    save_file_state("b.xlsx")
    state = load_state()
    assert "a.xlsx" in state and "b.xlsx" in state  # first not overwritten


def test_adversarial_corrupted_json(tmp_dirs, monkeypatch):
    state_file = tmp_dirs["log"] / "processed.json"
    monkeypatch.setattr("auto_excel.state.STATE_FILE", state_file)
    state_file.write_text("not valid json")
    state = load_state()
    assert state == {}  # treat corrupted as empty
