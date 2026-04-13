"""Tests for the display module."""
from io import StringIO
from rich.console import Console
from auto_excel.display import (
    print_start, print_file_start, print_step, print_file_done,
    print_file_error, print_report, print_no_files, print_exit,
)


def make_test_console():
    output = StringIO()
    console = Console(file=output, highlight=False)
    return console, output


def test_print_start_shows_count():
    console, output = make_test_console()
    print_start(3, console=console)
    assert "3" in output.getvalue()


def test_print_step_shows_message():
    console, output = make_test_console()
    print_step("正在计算", console=console)
    assert "正在计算" in output.getvalue()


def test_print_report_shows_table():
    console, output = make_test_console()
    results = [{"filename": "test.xlsx", "status": "success", "duration": 1.2}]
    print_report(results, console=console)
    text = output.getvalue()
    assert "test.xlsx" in text


def test_adversarial_print_report_mixed_status():
    results = [
        {"filename": "a.xlsx", "status": "success", "duration": 1.0},
        {"filename": "b.xlsx", "status": "error", "duration": 0},
    ]
    console, output = make_test_console()
    print_report(results, console=console)
    text = output.getvalue()
    assert "a.xlsx" in text and "b.xlsx" in text


def test_adversarial_print_no_files():
    console, output = make_test_console()
    print_no_files(console=console)
    assert len(output.getvalue()) > 0  # should output something


def test_adversarial_print_report_status_mapping():
    """Wrong impl that maps all rows to same status must be caught."""
    results = [
        {"filename": "ok.xlsx", "status": "success", "duration": 1.0},
        {"filename": "bad.xlsx", "status": "error", "duration": 0},
    ]
    console, output = make_test_console()
    print_report(results, console=console)
    text = output.getvalue()
    assert "✅" in text, "success row must show success symbol"
    assert "❌" in text, "error row must show failure symbol"
    assert "1.00" in text, "success row must show numeric duration"


def test_adversarial_print_start_varies_with_count():
    """Two different counts must produce different outputs."""
    console1, output1 = make_test_console()
    print_start(1, console=console1)
    console2, output2 = make_test_console()
    print_start(99, console=console2)
    assert "1" in output1.getvalue()
    assert "99" in output2.getvalue()
    assert output1.getvalue() != output2.getvalue()
