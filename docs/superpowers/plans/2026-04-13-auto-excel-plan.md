# Auto Excel Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers-pro:subagent-driven-development (recommended) or superpowers-pro:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a macOS CLI tool (`auto-excel on`) that automatically processes marketing Excel files — calculating derived columns, sorting, and generating group statistics — with real-time terminal progress and one-command installation.

**Architecture:** Python CLI application. Typer handles the `on` command, openpyxl manipulates Excel files, Rich renders terminal progress. The CLI orchestrator scans `~/Desktop/marketing analysis/Raw/`, processes each file through three sequential flows (column calc → sort → group), saves results to `New/`, and tracks state in `log/processed.json`.

**Tech Stack:** Python 3.11+, uv 0.11+, openpyxl 3.1.5+, Typer 0.9+, Rich 14.1+

---

## Part 1: Requirement Specification

### Sub-function overview

1. **File scanning** — Scan Raw/ for unprocessed .xlsx files, ordered by creation time
2. **Column calculation (Flow 1)** — Insert 4 calculated columns in Sheet 4
3. **Sorting (Flow 2)** — Sort all data rows by 实际成本 descending
4. **Grouping (Flow 3)** — Insert 占比 column, merge cells by cost range, fill count/percentage
5. **State tracking** — Maintain processed.json to skip already-processed files
6. **Progress display** — Real-time Chinese terminal output with emoji
7. **Installation** — One-command `curl | bash` install

### Sub-function details

#### 1. File scanning
- **Input:** `~/Desktop/marketing analysis/Raw/` directory
- **Behavior:** List `.xlsx` files → exclude files in `processed.json` → sort by creation time (oldest first)
- **No files:** Print friendly message, exit 0

#### 2. Column calculation (Flow 1)
- **Target:** Sheet 4 (index 3), Row 1 = headers
- **Required source columns:** 花费, 展现量, 点击量, 留资人数, 留资成本
- **Columns inserted in order (re-scan headers after each insertion):**

| New column | Insert after | Formula | Notes |
|-----------|-------------|---------|-------|
| 实际花费 | 花费 | 花费 ÷ 1.136 | Skip row 1 |
| 点击率 | 点击量 | 点击量 ÷ 展现量 | Standard CTR |
| CPC | 点击率 | 实际花费 ÷ 点击量 | Cost per click |
| 实际成本 | 留资成本 | 实际花费 ÷ 留资人数 | Cost per lead |

- **Zero division:** Cell = 0, log warning with row number
- **Column lookup:** By header name in row 1 (not fixed index)

#### 3. Sorting (Flow 2)
- Sort by 实际成本 descending; row 1 (header) fixed; entire rows move together

#### 4. Grouping (Flow 3)
- Insert 占比 column right of 互动成本 (header = "占比")
- Three ranges (contiguous after sort):

| Range | Condition | Cell content |
|-------|-----------|-------------|
| High | 实际成本 ≥ 90 | `count/percentage%` |
| Medium | 50 ≤ 实际成本 < 90 | `count/percentage%` |
| Low | 实际成本 < 50 | `count/percentage%` |

- Merge 占比 cells for each range; percentage = count ÷ total_rows
- If a range has 0 rows, skip (no merge)

#### 5. State tracking
- File: `log/processed.json`
- Format: `{"filename.xlsx": {"processed_at": "ISO8601", "status": "success"}}`
- Write immediately after each file succeeds; failed files not written

#### 6. Progress display
- Chinese, emoji-decorated, non-technical per-step output
- Summary table at end (Rich Table)
- Errors shown briefly; full stack trace only in log file

#### 7. Installation
- `curl -sSL <url> | bash` → clone to `~/.auto-excel/`, install uv, uv sync, register CLI globally, create desktop folders

### Non-functional requirements
- One file failure must not block subsequent files
- All terminal output in Chinese with emoji

## Part 2: CLI Internal Architecture

### Module interfaces

#### config.py
```python
BASE_DIR: Path    # ~/Desktop/marketing analysis
RAW_DIR: Path     # BASE_DIR / "Raw"
NEW_DIR: Path     # BASE_DIR / "New"
LOG_DIR: Path     # BASE_DIR / "log"
STATE_FILE: Path  # LOG_DIR / "processed.json"
```

#### state.py
```python
def load_state() -> dict[str, dict]
def save_file_state(filename: str) -> None
def is_processed(filename: str) -> bool
```

#### processor.py
```python
def find_column(ws: Worksheet, name: str) -> int
def insert_calculated_column(ws: Worksheet, after_col_name: str, new_col_name: str, formula_fn: Callable) -> None
def sort_by_column(ws: Worksheet, col_name: str, descending: bool = True) -> None
def group_and_merge(ws: Worksheet, value_col_name: str, insert_after_col: str, new_col_name: str, ranges: list[tuple]) -> None
def process_file(src: Path, dst: Path, on_step: Callable[[str], None] | None = None) -> None
```

#### display.py
```python
def print_start(file_count: int) -> None
def print_file_start(index: int, total: int, filename: str) -> None
def print_step(message: str) -> None
def print_file_done(duration: float) -> None
def print_file_error(filename: str, error: str) -> None
def print_report(results: list[dict]) -> None
def print_no_files() -> None
def print_exit() -> None
```

#### cli.py
```python
app = typer.Typer()
@app.command()
def on() -> None  # Scan → filter → for each: process_file() → state → report
```

---

### Task 1: Project Scaffolding

**Type:** Non-code task

**Context Brief:**
- **Files:** `pyproject.toml` (create), `.python-version` (create), `src/auto_excel/__init__.py` (create)
- **Constraints:** Use uv for package management; register CLI command `auto-excel` via `[project.scripts]`

**Files:**
- Create: `pyproject.toml`
- Create: `.python-version`
- Create: `src/auto_excel/__init__.py`

- [ ] **Step 1: Create pyproject.toml**

```toml
[build-system]
requires = ["hatchling"]
build-backend = "hatchling.build"

[project]
name = "auto-excel"
version = "0.1.0"
description = "Marketing Excel automation CLI"
requires-python = ">=3.11"
dependencies = [
    "openpyxl>=3.1.5",
    "typer>=0.9.0",
    "rich>=14.0.0",
]

[project.scripts]
auto-excel = "auto_excel.cli:app"

[dependency-groups]
dev = ["pytest>=8.0"]
```

- [ ] **Step 2: Create .python-version**

Content: `3.11`

- [ ] **Step 3: Create src/auto_excel/__init__.py**

```python
__version__ = "0.1.0"
```

- [ ] **Step 4: Run uv sync**

Run: `cd ~/work_project/auto-excel && uv sync`
Expected: `.venv/` created, all dependencies installed, no errors

- [ ] **Step 5: Verify CLI command registers**

Run: `cd ~/work_project/auto-excel && uv run auto-excel --help`
Expected: Shows Typer help output (will fail with import error since cli.py doesn't exist yet — that's OK, just verify uv resolves the command)

- [ ] **Step 6: Commit**

```bash
git add pyproject.toml .python-version src/auto_excel/__init__.py
git commit -m "chore: scaffold project with pyproject.toml and uv"
```

---

### Task 2: Config Module

**Type:** Code task

**Context Brief:**
- **Files:** `src/auto_excel/config.py` (create), `tests/test_config.py` (create)
- **Interface:** Module-level `Path` constants: `BASE_DIR`, `RAW_DIR`, `NEW_DIR`, `LOG_DIR`, `STATE_FILE`
- **Constraints:** `BASE_DIR` must use `Path.home() / "Desktop" / "marketing analysis"` for portability across macOS users

**Files:**
- Create: `src/auto_excel/config.py`
- Test: `tests/test_config.py`

- [ ] **Step 1: Write the failing test**

`tests/test_config.py`:
```python
def test_base_dir_is_on_desktop():
    assert "Desktop" in str(config.BASE_DIR)
    assert config.BASE_DIR.name == "marketing analysis"

def test_subdirs_are_under_base():
    assert config.RAW_DIR == config.BASE_DIR / "Raw"
    assert config.NEW_DIR == config.BASE_DIR / "New"
    assert config.LOG_DIR == config.BASE_DIR / "log"

def test_state_file_is_in_log_dir():
    assert config.STATE_FILE == config.LOG_DIR / "processed.json"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd ~/work_project/auto-excel && uv run pytest tests/test_config.py -v`
Expected: FAIL (module not found)

- [ ] **Step 3: Write minimal implementation**

`src/auto_excel/config.py` — define 5 `Path` constants using `Path.home()`.

- [ ] **Step 4: Run test to verify it passes**

Run: `cd ~/work_project/auto-excel && uv run pytest tests/test_config.py -v`
Expected: All 3 tests PASS

- [ ] **Step 5: Write adversarial tests**

```python
def test_adversarial_paths_are_absolute():
    for p in [config.BASE_DIR, config.RAW_DIR, config.NEW_DIR, config.LOG_DIR]:
        assert p.is_absolute()

def test_adversarial_state_file_has_json_extension():
    assert config.STATE_FILE.suffix == ".json"
```

- [ ] **Step 6: Run adversarial quality gate**

Run: `cd ~/work_project/auto-excel && uv run pytest tests/test_config.py -v`
Expected: All PASS

- [ ] **Step 7: Commit**

```bash
git add src/auto_excel/config.py tests/test_config.py
git commit -m "feat: add config module with path constants"
```

---

### Task 3: Test Fixture Factory

**Type:** Code task

**Context Brief:**
- **Files:** `tests/conftest.py` (create)
- **Interface:** `make_sample_workbook(rows: list[dict]) -> Workbook` — creates an openpyxl Workbook with 4 sheets; Sheet 4 has marketing data columns with provided rows
- **Dependencies:** Used by all subsequent processor tests
- **Constraints:** Sheet 4 must have headers: 日期, 计划名称, 花费, 展现量, 点击量, 留资人数, 留资成本, 互动成本 (in row 1). Must also provide `tmp_dirs` fixture that creates temp Raw/New/log directories.

**Files:**
- Create: `tests/conftest.py`
- Test: `tests/test_conftest.py`

- [ ] **Step 1: Write test for the fixture**

`tests/test_conftest.py`:
```python
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd ~/work_project/auto-excel && uv run pytest tests/test_conftest.py -v`
Expected: FAIL (fixture not found)

- [ ] **Step 3: Implement conftest.py**

`tests/conftest.py` — define `make_sample_workbook` fixture that returns a factory function. The factory creates a Workbook with 4 sheets (Sheet1-Sheet3 empty, Sheet4 with marketing headers and data rows). Also define `tmp_dirs` fixture using `tmp_path` that creates Raw/, New/, log/ subdirectories and returns a dict of paths.

- [ ] **Step 4: Run test to verify it passes**

Run: `cd ~/work_project/auto-excel && uv run pytest tests/test_conftest.py -v`
Expected: All 3 tests PASS

- [ ] **Step 5: Write adversarial tests**

```python
def test_adversarial_empty_rows(make_sample_workbook):
    wb = make_sample_workbook([])
    ws = wb.worksheets[3]
    assert ws.max_row == 1  # header only

def test_adversarial_missing_fields_get_defaults(make_sample_workbook):
    wb = make_sample_workbook([{"花费": 100}])  # other fields missing
    ws = wb.worksheets[3]
    assert ws.max_row == 2  # should still create the row
```

- [ ] **Step 6: Run adversarial quality gate**

Run: `cd ~/work_project/auto-excel && uv run pytest tests/test_conftest.py -v`
Expected: All PASS

- [ ] **Step 7: Commit**

```bash
git add tests/conftest.py tests/test_conftest.py
git commit -m "test: add sample workbook fixture factory"
```

---

### Task 4: State Management Module

**Type:** Code task

**Context Brief:**
- **Files:** `src/auto_excel/state.py` (create), `tests/test_state.py` (create)
- **Interface:** `load_state() -> dict`, `save_file_state(filename: str) -> None`, `is_processed(filename: str) -> bool`
- **Dependencies:** Reads `config.STATE_FILE` from `config.py`
- **Constraints:** `save_file_state` must write immediately (not batch). File not in state = unprocessed. State file may not exist on first run.

**Files:**
- Create: `src/auto_excel/state.py`
- Test: `tests/test_state.py`

- [ ] **Step 1: Write failing tests**

`tests/test_state.py`:
```python
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd ~/work_project/auto-excel && uv run pytest tests/test_state.py -v`
Expected: FAIL (import error)

- [ ] **Step 3: Implement state.py**

Implement three functions using `json` module. `load_state` reads JSON file or returns `{}` if missing. `save_file_state` loads existing state, adds entry with ISO timestamp and `"success"` status, writes back. `is_processed` delegates to `load_state`.

- [ ] **Step 4: Run test to verify it passes**

Run: `cd ~/work_project/auto-excel && uv run pytest tests/test_state.py -v`
Expected: All 3 PASS

- [ ] **Step 5: Write adversarial tests**

```python
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
    # Should handle gracefully — return empty or raise clear error
    state = load_state()
    assert state == {}  # treat corrupted as empty
```

- [ ] **Step 6: Run adversarial quality gate**

Run: `cd ~/work_project/auto-excel && uv run pytest tests/test_state.py -v`
Expected: All PASS

- [ ] **Step 7: Commit**

```bash
git add src/auto_excel/state.py tests/test_state.py
git commit -m "feat: add state management for tracking processed files"
```

---

### Task 5: Excel Column Helpers

**Type:** Code task

**Context Brief:**
- **Files:** `src/auto_excel/processor.py` (create), `tests/test_column_helpers.py` (create)
- **Interface:** `find_column(ws, name: str) -> int` returns 1-based column index; raises `ValueError` if not found. `insert_calculated_column(ws, after_col_name: str, new_col_name: str, formula_fn: Callable[[float], float]) -> None` inserts a new column, writes header, fills data rows.
- **Dependencies:** Uses openpyxl `Worksheet`
- **Constraints:** `find_column` scans row 1 only. `insert_calculated_column` must handle zero-division in `formula_fn` by catching `ZeroDivisionError` and filling 0.

**Files:**
- Create: `src/auto_excel/processor.py`
- Test: `tests/test_column_helpers.py`

- [ ] **Step 1: Write failing tests**

`tests/test_column_helpers.py`:
```python
def test_find_column_returns_correct_index(make_sample_workbook):
    wb = make_sample_workbook([{"花费": 100}])
    ws = wb.worksheets[3]
    idx = find_column(ws, "花费")
    assert ws.cell(1, idx).value == "花费"

def test_find_column_raises_on_missing(make_sample_workbook):
    wb = make_sample_workbook([{"花费": 100}])
    ws = wb.worksheets[3]
    with pytest.raises(ValueError):
        find_column(ws, "不存在的列")

def test_insert_column_adds_header_and_data(make_sample_workbook):
    wb = make_sample_workbook([{"花费": 113.6}, {"花费": 227.2}])
    ws = wb.worksheets[3]
    insert_calculated_column(ws, "花费", "实际花费", lambda val, row: val / 1.136)
    new_idx = find_column(ws, "实际花费")
    assert ws.cell(1, new_idx).value == "实际花费"
    assert abs(ws.cell(2, new_idx).value - 100.0) < 0.01
    assert abs(ws.cell(3, new_idx).value - 200.0) < 0.01
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd ~/work_project/auto-excel && uv run pytest tests/test_column_helpers.py -v`
Expected: FAIL (import error)

- [ ] **Step 3: Implement find_column and insert_calculated_column**

`src/auto_excel/processor.py`:
- `find_column`: iterate row 1 cells, return column index where `cell.value == name`
- `insert_calculated_column`: call `find_column` to locate `after_col_name`, `ws.insert_cols(idx + 1)`, write `new_col_name` in row 1, iterate rows 2..max_row, read source value, apply `formula_fn`, write result. Catch `ZeroDivisionError` → write 0.

- [ ] **Step 4: Run test to verify it passes**

Run: `cd ~/work_project/auto-excel && uv run pytest tests/test_column_helpers.py -v`
Expected: All 3 PASS

- [ ] **Step 5: Write adversarial tests**

```python
def test_adversarial_zero_division_fills_zero(make_sample_workbook):
    wb = make_sample_workbook([{"花费": 100, "点击量": 0}])
    ws = wb.worksheets[3]
    click_idx = find_column(ws, "点击量")
    insert_calculated_column(ws, "花费", "test", lambda val, row: ws.cell(row, click_idx).value / val if val else 0)
    # Should not raise, should fill 0

def test_adversarial_insert_shifts_columns(make_sample_workbook):
    wb = make_sample_workbook([{"花费": 100}])
    ws = wb.worksheets[3]
    old_cols = ws.max_column
    insert_calculated_column(ws, "花费", "new_col", lambda val, row: val * 2)
    assert ws.max_column == old_cols + 1

def test_adversarial_find_after_insert(make_sample_workbook):
    wb = make_sample_workbook([{"花费": 100, "点击量": 50}])
    ws = wb.worksheets[3]
    insert_calculated_column(ws, "花费", "实际花费", lambda val, row: val / 1.136)
    # 点击量 should still be findable after insertion
    idx = find_column(ws, "点击量")
    assert ws.cell(1, idx).value == "点击量"
```

- [ ] **Step 6: Run adversarial quality gate**

Run: `cd ~/work_project/auto-excel && uv run pytest tests/test_column_helpers.py -v`
Expected: All PASS

- [ ] **Step 7: Commit**

```bash
git add src/auto_excel/processor.py tests/test_column_helpers.py
git commit -m "feat: add Excel column helpers (find and insert)"
```

---

### Task 6: Flow 1 — Column Calculation

**Type:** Code task

**Context Brief:**
- **Files:** `src/auto_excel/processor.py` (modify), `tests/test_flow1.py` (create)
- **Interface:** `apply_calculated_columns(ws: Worksheet) -> None` — inserts all 4 columns in order: 实际花费, 点击率, CPC, 实际成本. Each insertion re-scans headers.
- **Dependencies:** Calls `find_column` and `insert_calculated_column` from same module
- **Constraints:** Formulas: 实际花费=花费÷1.136, 点击率=点击量÷展现量, CPC=实际花费÷点击量, 实际成本=实际花费÷留资人数. CPC and 实际成本 read from the just-inserted 实际花费 column.

**Files:**
- Modify: `src/auto_excel/processor.py`
- Test: `tests/test_flow1.py`

- [ ] **Step 1: Write failing tests**

`tests/test_flow1.py` — create workbook with known data, call `apply_calculated_columns(ws)`, then verify:
```python
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd ~/work_project/auto-excel && uv run pytest tests/test_flow1.py -v`
Expected: FAIL (function not found)

- [ ] **Step 3: Implement apply_calculated_columns**

Add to `processor.py`. The function calls `insert_calculated_column` four times in sequence. For CPC and 实际成本, the `formula_fn` must read the already-inserted 实际花费 column value from the current row (use `find_column(ws, "实际花费")` inside the lambda/closure to get the current index).

- [ ] **Step 4: Run test to verify it passes**

Run: `cd ~/work_project/auto-excel && uv run pytest tests/test_flow1.py -v`
Expected: All PASS

- [ ] **Step 5: Write adversarial tests**

```python
def test_adversarial_multiple_rows(make_sample_workbook):
    rows = [
        {"花费": 113.6, "展现量": 1000, "点击量": 50, "留资人数": 5, "留资成本": 22.72, "互动成本": 10},
        {"花费": 227.2, "展现量": 2000, "点击量": 100, "留资人数": 10, "留资成本": 22.72, "互动成本": 20},
    ]
    wb = make_sample_workbook(rows)
    ws = wb.worksheets[3]
    apply_calculated_columns(ws)
    cost_idx = find_column(ws, "实际花费")
    assert abs(ws.cell(2, cost_idx).value - 100.0) < 0.01
    assert abs(ws.cell(3, cost_idx).value - 200.0) < 0.01  # row 3

def test_adversarial_zero_clicks(make_sample_workbook):
    wb = make_sample_workbook([{"花费": 100, "展现量": 1000, "点击量": 0, "留资人数": 5, "留资成本": 20, "互动成本": 10}])
    ws = wb.worksheets[3]
    apply_calculated_columns(ws)
    ctr_idx = find_column(ws, "点击率")
    cpc_idx = find_column(ws, "CPC")
    assert ws.cell(2, ctr_idx).value == 0  # 0/1000 = 0
    assert ws.cell(2, cpc_idx).value == 0  # division by zero → 0
```

- [ ] **Step 6: Run adversarial quality gate**

Run: `cd ~/work_project/auto-excel && uv run pytest tests/test_flow1.py -v`
Expected: All PASS

- [ ] **Step 7: Commit**

```bash
git add src/auto_excel/processor.py tests/test_flow1.py
git commit -m "feat: implement Flow 1 — insert 4 calculated columns"
```

---

### Task 7: Flow 2 — Sorting

**Type:** Code task

**Context Brief:**
- **Files:** `src/auto_excel/processor.py` (modify), `tests/test_flow2.py` (create)
- **Interface:** `sort_by_column(ws: Worksheet, col_name: str, descending: bool = True) -> None` — sorts data rows (2+) by named column, row 1 fixed
- **Dependencies:** Calls `find_column` from same module
- **Constraints:** Entire rows must move together (all columns). Must handle mixed numeric types.

**Files:**
- Modify: `src/auto_excel/processor.py`
- Test: `tests/test_flow2.py`

- [ ] **Step 1: Write failing tests**

`tests/test_flow2.py`:
```python
def test_sort_descending(make_sample_workbook):
    rows = [
        {"花费": 50, "展现量": 100, "点击量": 10, "留资人数": 1, "留资成本": 50, "互动成本": 5},
        {"花费": 200, "展现量": 100, "点击量": 10, "留资人数": 1, "留资成本": 200, "互动成本": 5},
        {"花费": 100, "展现量": 100, "点击量": 10, "留资人数": 1, "留资成本": 100, "互动成本": 5},
    ]
    wb = make_sample_workbook(rows)
    ws = wb.worksheets[3]
    sort_by_column(ws, "花费", descending=True)
    cost_idx = find_column(ws, "花费")
    assert ws.cell(2, cost_idx).value == 200
    assert ws.cell(3, cost_idx).value == 100
    assert ws.cell(4, cost_idx).value == 50

def test_sort_preserves_header(make_sample_workbook):
    rows = [{"花费": 999, "展现量": 100, "点击量": 10, "留资人数": 1, "留资成本": 50, "互动成本": 5}]
    wb = make_sample_workbook(rows)
    ws = wb.worksheets[3]
    sort_by_column(ws, "花费")
    assert ws.cell(1, find_column(ws, "花费")).value == "花费"  # header unchanged
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd ~/work_project/auto-excel && uv run pytest tests/test_flow2.py -v`
Expected: FAIL

- [ ] **Step 3: Implement sort_by_column**

Add to `processor.py`. Read all data rows (2..max_row) into list of row-value-lists, sort by target column value, write back. Row 1 untouched.

- [ ] **Step 4: Run test to verify it passes**

Run: `cd ~/work_project/auto-excel && uv run pytest tests/test_flow2.py -v`
Expected: All PASS

- [ ] **Step 5: Write adversarial tests**

```python
def test_adversarial_sort_moves_entire_row(make_sample_workbook):
    rows = [
        {"花费": 50, "展现量": 111, "点击量": 10, "留资人数": 1, "留资成本": 50, "互动成本": 5},
        {"花费": 200, "展现量": 222, "点击量": 10, "留资人数": 1, "留资成本": 200, "互动成本": 5},
    ]
    wb = make_sample_workbook(rows)
    ws = wb.worksheets[3]
    sort_by_column(ws, "花费", descending=True)
    # Row with 花费=200 should be first, and its 展现量 should be 222
    cost_idx = find_column(ws, "花费")
    imp_idx = find_column(ws, "展现量")
    assert ws.cell(2, cost_idx).value == 200
    assert ws.cell(2, imp_idx).value == 222  # entire row moved

def test_adversarial_single_row(make_sample_workbook):
    wb = make_sample_workbook([{"花费": 100, "展现量": 100, "点击量": 10, "留资人数": 1, "留资成本": 100, "互动成本": 5}])
    ws = wb.worksheets[3]
    sort_by_column(ws, "花费")  # should not error with 1 row
    assert ws.cell(2, find_column(ws, "花费")).value == 100
```

- [ ] **Step 6: Run adversarial quality gate**

Run: `cd ~/work_project/auto-excel && uv run pytest tests/test_flow2.py -v`
Expected: All PASS

- [ ] **Step 7: Commit**

```bash
git add src/auto_excel/processor.py tests/test_flow2.py
git commit -m "feat: implement Flow 2 — sort by column descending"
```

---

### Task 8: Flow 3 — Grouping & Merging

**Type:** Code task

**Context Brief:**
- **Files:** `src/auto_excel/processor.py` (modify), `tests/test_flow3.py` (create)
- **Interface:** `group_and_merge(ws: Worksheet) -> None` — inserts 占比 column right of 互动成本, groups rows by 实际成本 ranges, merges 占比 cells per range, fills `count/percentage%`
- **Dependencies:** Calls `find_column` from same module; uses `openpyxl.worksheet.merge_cells`
- **Constraints:** Three ranges: ≥90, 50–89, <50. Data must be pre-sorted (contiguous ranges). Ranges with 0 rows are skipped. Percentage = count ÷ total data rows, formatted as integer percent.

**Files:**
- Modify: `src/auto_excel/processor.py`
- Test: `tests/test_flow3.py`

- [ ] **Step 1: Write failing tests**

`tests/test_flow3.py`:
```python
def test_group_inserts_zhanbi_column(make_sample_workbook):
    # Pre-sorted by 实际成本 desc. Need 实际成本 column to exist.
    wb = make_sample_workbook([...])  # see step details below
    ws = wb.worksheets[3]
    apply_calculated_columns(ws)
    sort_by_column(ws, "实际成本")
    group_and_merge(ws)
    assert find_column(ws, "占比")  # column exists

def test_group_merge_content():
    # Create workbook with data that produces known ranges:
    # 2 rows with 实际成本 >= 90, 1 row with 50-89, 2 rows < 50 (total 5)
    # After merge: "2/40%", "1/20%", "2/40%"
```

Construct 5 rows where 花费 and 留资人数 are chosen so that 实际花费÷留资人数 yields known 实际成本 values. For example:
- Row A: 花费=113.6, 留资人数=1 → 实际成本=100 (≥90)
- Row B: 花费=113.6, 留资人数=1 → 实际成本=100 (≥90)
- Row C: 花费=113.6, 留资人数=1.5 → 实际成本≈66.7 (50-89)
- Row D: 花费=113.6, 留资人数=5 → 实际成本=20 (<50)
- Row E: 花费=113.6, 留资人数=4 → 实际成本=25 (<50)

Assert merged cell values match `"2/40%"`, `"1/20%"`, `"2/40%"`.

- [ ] **Step 2: Run test to verify it fails**

Run: `cd ~/work_project/auto-excel && uv run pytest tests/test_flow3.py -v`
Expected: FAIL

- [ ] **Step 3: Implement group_and_merge**

Add to `processor.py`:
1. `find_column(ws, "互动成本")` → insert column at idx+1, write "占比" header
2. `find_column(ws, "实际成本")` → read values from rows 2..max_row
3. Partition rows into three contiguous groups based on value
4. For each non-empty group: `ws.merge_cells(...)` spanning the 占比 column rows, write `f"{count}/{percentage}%"` to top-left cell of merged range

- [ ] **Step 4: Run test to verify it passes**

Run: `cd ~/work_project/auto-excel && uv run pytest tests/test_flow3.py -v`
Expected: All PASS

- [ ] **Step 5: Write adversarial tests**

```python
def test_adversarial_all_in_one_range():
    # All rows have 实际成本 >= 90 → only one merged group, "5/100%"

def test_adversarial_empty_range():
    # No rows in medium range → only high and low groups, no merge for medium

def test_adversarial_single_row():
    # Only 1 data row → one group, "1/100%", no actual merge needed
```

- [ ] **Step 6: Run adversarial quality gate**

Run: `cd ~/work_project/auto-excel && uv run pytest tests/test_flow3.py -v`
Expected: All PASS

- [ ] **Step 7: Commit**

```bash
git add src/auto_excel/processor.py tests/test_flow3.py
git commit -m "feat: implement Flow 3 — group by cost range with merged cells"
```

---

### Task 9: Display Module

**Type:** Code task

**Context Brief:**
- **Files:** `src/auto_excel/display.py` (create), `tests/test_display.py` (create)
- **Interface:** `print_start(count)`, `print_file_start(idx, total, name)`, `print_step(msg)`, `print_file_done(duration)`, `print_file_error(name, err)`, `print_report(results)`, `print_no_files()`, `print_exit()`
- **Constraints:** All output in Chinese with emoji. Use `rich.console.Console` for styled output. Use `rich.table.Table` for the report. The console instance should be module-level for testability.

**Files:**
- Create: `src/auto_excel/display.py`
- Test: `tests/test_display.py`

- [ ] **Step 1: Write failing tests**

`tests/test_display.py` — use `rich.console.Console(file=StringIO())` to capture output:
```python
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd ~/work_project/auto-excel && uv run pytest tests/test_display.py -v`
Expected: FAIL

- [ ] **Step 3: Implement display.py**

All functions accept optional `console` parameter (defaults to module-level `Console()`). Each function uses `console.print()` with appropriate emoji and formatting. `print_report` builds a `rich.table.Table` with columns: 文件名, 状态, 耗时.

- [ ] **Step 4: Run test to verify it passes**

Run: `cd ~/work_project/auto-excel && uv run pytest tests/test_display.py -v`
Expected: All PASS

- [ ] **Step 5: Write adversarial tests**

```python
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
```

- [ ] **Step 6: Run adversarial quality gate**

Run: `cd ~/work_project/auto-excel && uv run pytest tests/test_display.py -v`
Expected: All PASS

- [ ] **Step 7: Commit**

```bash
git add src/auto_excel/display.py tests/test_display.py
git commit -m "feat: add Rich display module for terminal progress"
```

---

### Task 10: CLI Entry Point & Orchestration

**Type:** Code task

**Context Brief:**
- **Files:** `src/auto_excel/cli.py` (create), `tests/test_cli.py` (create)
- **Interface:** `app = typer.Typer()`, `on()` command — scans Raw/, filters by state, processes each file, displays progress, outputs report
- **Dependencies:** Calls `config.*`, `state.*`, `processor.process_file()`, `display.*`
- **Constraints:** Processing failures must not stop other files. Logging to `log/YYYY-MM-DD.log` via Python `logging` module.

**Files:**
- Create: `src/auto_excel/cli.py`
- Test: `tests/test_cli.py`

- [ ] **Step 1: Add process_file to processor.py**

First add the top-level orchestration function to `processor.py`:
```python
def process_file(src: Path, dst: Path, on_step: Callable[[str], None] | None = None) -> None
```
This function: copies src→dst, opens dst workbook, gets Sheet 4, calls `apply_calculated_columns`, `sort_by_column`, `group_and_merge`, saves workbook. Calls `on_step(message)` before each major operation.

- [ ] **Step 2: Write failing test for CLI**

`tests/test_cli.py` — use `typer.testing.CliRunner`:
```python
def test_on_no_files(tmp_dirs, monkeypatch):
    # Monkeypatch config paths to tmp_dirs
    result = runner.invoke(app, ["on"])
    assert result.exit_code == 0
    assert "待处理" in result.output or "没有" in result.output

def test_on_processes_file(tmp_dirs, monkeypatch, make_sample_workbook):
    # Save sample workbook to tmp Raw/ dir
    # Monkeypatch config paths
    result = runner.invoke(app, ["on"])
    assert result.exit_code == 0
    # Verify New/ has the output file
    # Verify processed.json has the entry
```

- [ ] **Step 3: Run test to verify it fails**

Run: `cd ~/work_project/auto-excel && uv run pytest tests/test_cli.py -v`
Expected: FAIL

- [ ] **Step 4: Implement cli.py**

Create Typer app with `on` command. The command:
1. Ensure directories exist
2. Scan Raw/ for .xlsx files
3. Filter unprocessed via `state.is_processed`
4. Sort by `os.path.getctime`
5. If none → `display.print_no_files()`, return
6. `display.print_start(count)`
7. For each file: try/except around `process_file(src, dst, on_step=display.print_step)`, track result, timing
8. `display.print_report(results)`
9. `display.print_exit()`

Also set up Python `logging` to write to `log/YYYY-MM-DD.log`.

- [ ] **Step 5: Run test to verify it passes**

Run: `cd ~/work_project/auto-excel && uv run pytest tests/test_cli.py -v`
Expected: All PASS

- [ ] **Step 6: Write adversarial tests**

```python
def test_adversarial_skips_already_processed(tmp_dirs, monkeypatch):
    # Pre-populate processed.json, verify file is skipped

def test_adversarial_continues_after_error(tmp_dirs, monkeypatch):
    # Place one valid and one invalid Excel file
    # Verify both are attempted, one succeeds
```

- [ ] **Step 7: Run adversarial quality gate**

Run: `cd ~/work_project/auto-excel && uv run pytest tests/test_cli.py -v`
Expected: All PASS

- [ ] **Step 8: Commit**

```bash
git add src/auto_excel/cli.py src/auto_excel/processor.py tests/test_cli.py
git commit -m "feat: add CLI entry point with full orchestration"
```

---

### Task 11: Install Script

**Type:** Code task

**Context Brief:**
- **Files:** `install.sh` (create)
- **Constraints:** Must be a single `curl | bash` entry point. Clones repo to `~/.auto-excel/`, installs uv if missing, runs `uv sync`, creates wrapper in `~/.local/bin/auto-excel`, ensures PATH, creates desktop folders. Must be idempotent (safe to run multiple times).

**Files:**
- Create: `install.sh`

- [ ] **Step 1: Write install.sh**

The script must:
1. `set -e` for fail-fast
2. Define color output helpers (info/success/error)
3. Clone repo to `~/.auto-excel/` (skip if already exists, `git pull` instead)
4. Check `command -v uv` → if missing, install via `curl -LsSf https://astral.sh/uv/install.sh | sh`
5. `cd ~/.auto-excel && uv sync`
6. Create `~/.local/bin/auto-excel` wrapper script:
   ```bash
   #!/bin/bash
   cd ~/.auto-excel && uv run auto-excel "$@"
   ```
7. `chmod +x ~/.local/bin/auto-excel`
8. Check if `~/.local/bin` in PATH → if not, append `export PATH="$HOME/.local/bin:$PATH"` to `~/.zshrc`
9. `mkdir -p ~/Desktop/marketing\ analysis/{Raw,New,log}`
10. Print success message with usage instructions

- [ ] **Step 2: Make executable and test locally**

Run: `chmod +x ~/work_project/auto-excel/install.sh`
Verify: `bash -n ~/work_project/auto-excel/install.sh` (syntax check, no execution)

- [ ] **Step 3: Commit**

```bash
git add install.sh
git commit -m "feat: add one-command install script"
```

---

### Task 12: End-to-End Integration Test

**Type:** Code task

**Context Brief:**
- **Files:** `tests/test_integration.py` (create)
- **Interface:** Full pipeline test: create sample Excel → save to temp Raw/ → invoke `auto-excel on` → verify New/ output has all expected columns, sorting, and merged cells
- **Dependencies:** All modules (config, state, processor, display, cli)
- **Constraints:** Must use temp directories (monkeypatched config paths). Must verify: 4 new columns exist with correct values, data sorted by 实际成本 desc, 占比 column has merged cells with correct content.

**Files:**
- Create: `tests/test_integration.py`

- [ ] **Step 1: Write integration test**

`tests/test_integration.py`:
```python
def test_full_pipeline(tmp_dirs, monkeypatch, make_sample_workbook):
    # 1. Create workbook with 5 rows of known data
    # 2. Save to tmp Raw/ directory
    # 3. Monkeypatch all config paths to tmp_dirs
    # 4. Invoke CLI: runner.invoke(app, ["on"])
    # 5. Assert exit_code == 0
    # 6. Open output file from New/
    # 7. Verify Sheet 4 has 实际花费, 点击率, CPC, 实际成本 columns
    # 8. Verify values are correct
    # 9. Verify sorted by 实际成本 descending
    # 10. Verify 占比 column exists with merged cells
    # 11. Verify processed.json has the file entry
```

Use 5 data rows with values that produce:
- 2 rows ≥90, 1 row 50-89, 2 rows <50 实际成本

- [ ] **Step 2: Run integration test**

Run: `cd ~/work_project/auto-excel && uv run pytest tests/test_integration.py -v`
Expected: PASS

- [ ] **Step 3: Run full test suite**

Run: `cd ~/work_project/auto-excel && uv run pytest -v`
Expected: All tests PASS

- [ ] **Step 4: Commit**

```bash
git add tests/test_integration.py
git commit -m "test: add end-to-end integration test for full pipeline"
```
