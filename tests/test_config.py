"""Tests for the config module."""

from auto_excel import config


def test_base_dir_is_on_desktop():
    assert "Desktop" in str(config.BASE_DIR)
    assert config.BASE_DIR.name == "marketing analysis"


def test_subdirs_are_under_base():
    assert config.RAW_DIR == config.BASE_DIR / "Raw"
    assert config.NEW_DIR == config.BASE_DIR / "New"
    assert config.LOG_DIR == config.BASE_DIR / "log"


def test_state_file_is_in_log_dir():
    assert config.STATE_FILE == config.LOG_DIR / "processed.json"


def test_adversarial_paths_are_absolute():
    for p in [config.BASE_DIR, config.RAW_DIR, config.NEW_DIR, config.LOG_DIR, config.STATE_FILE]:
        assert p.is_absolute()


def test_adversarial_state_file_has_json_extension():
    assert config.STATE_FILE.suffix == ".json"


def test_adversarial_dir_names_are_correct():
    """Verify that each directory and file has the correct name."""
    assert config.RAW_DIR.name == "Raw"
    assert config.NEW_DIR.name == "New"
    assert config.LOG_DIR.name == "log"
    assert config.STATE_FILE.name == "processed.json"


def test_adversarial_state_file_name():
    """Verify that STATE_FILE has the correct complete filename."""
    assert config.STATE_FILE.name == "processed.json"


def test_adversarial_constants_are_path_instances():
    """Verify that all constants are Path instances."""
    from pathlib import Path

    for p in [config.BASE_DIR, config.RAW_DIR, config.NEW_DIR, config.LOG_DIR, config.STATE_FILE]:
        assert isinstance(p, Path)
