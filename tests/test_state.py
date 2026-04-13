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


def test_adversarial_is_processed_key_specificity(tmp_dirs, monkeypatch):
    """is_processed must check the exact filename key, not just file existence."""
    monkeypatch.setattr("auto_excel.state.STATE_FILE", tmp_dirs["log"] / "processed.json")
    save_file_state("other.xlsx")
    assert is_processed("test.xlsx") is False
    assert is_processed("other.xlsx") is True


def test_adversarial_processed_at_is_iso8601(tmp_dirs, monkeypatch):
    """processed_at must be parseable as ISO 8601 datetime."""
    from datetime import datetime
    monkeypatch.setattr("auto_excel.state.STATE_FILE", tmp_dirs["log"] / "processed.json")
    save_file_state("test.xlsx")
    state = load_state()
    ts = state["test.xlsx"]["processed_at"]
    parsed = datetime.fromisoformat(ts)
    assert isinstance(parsed, datetime)


def test_adversarial_repeated_save_overwrites_not_duplicates(tmp_dirs, monkeypatch):
    """Saving same filename twice must overwrite, not create duplicates."""
    monkeypatch.setattr("auto_excel.state.STATE_FILE", tmp_dirs["log"] / "processed.json")
    save_file_state("dup.xlsx")
    first_ts = load_state()["dup.xlsx"]["processed_at"]
    save_file_state("dup.xlsx")
    state = load_state()
    assert list(state.keys()).count("dup.xlsx") == 1
    assert state["dup.xlsx"]["processed_at"] >= first_ts


def test_adversarial_empty_filename(tmp_dirs, monkeypatch):
    """Empty string must be treated as a valid (if unusual) key."""
    monkeypatch.setattr("auto_excel.state.STATE_FILE", tmp_dirs["log"] / "processed.json")
    assert is_processed("") is False
    save_file_state("")
    assert is_processed("") is True
    assert is_processed("real.xlsx") is False


def test_adversarial_corrupted_file_then_save(tmp_dirs, monkeypatch):
    """After load_state() returns {} for corrupted file, save_file_state must write valid JSON."""
    import json
    state_file = tmp_dirs["log"] / "processed.json"
    monkeypatch.setattr("auto_excel.state.STATE_FILE", state_file)
    state_file.write_text("{{invalid json")
    save_file_state("after_corrupt.xlsx")
    content = json.loads(state_file.read_text())
    assert "after_corrupt.xlsx" in content
    assert content["after_corrupt.xlsx"]["status"] == "success"
