# Formula-Aware Excel Preprocessing Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers-pro:subagent-driven-development (recommended) or superpowers-pro:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Fix three related bugs: formula cells reading as 0, workbook formulas destroyed on save, and empty rows polluting group statistics.

**Architecture:** Insert a formula resolution layer (`resolve_formulas`) and empty row filter (`remove_empty_rows`) into the existing `process_file` pipeline, between `load_workbook` and `apply_calculated_columns`. Switch from `data_only=True` to `data_only=False` to preserve workbook formulas.

**Tech Stack:** Python 3, openpyxl, pytest

---

## Part 1: Requirement Specification

### Sub-function overview

1. **SUMIFS formula resolution** — detect SUMIFS formulas in Sheet 4, read source sheet, aggregate by key, fill computed values
2. **Division formula resolution** — detect `=X/Y` formulas, compute after SUMIFS are resolved
3. **Empty row filtering** — remove rows where title and ID are both None
4. **process_file integration** — wire new functions into the pipeline, switch to `data_only=False`

### Sub-function details

**1. SUMIFS formula resolution**
- **Input:** Workbook (wb), target worksheet (ws)
- **Detection:** Scan row 2 for cells matching `=SUMIFS('<sheet>'!<col>:<col>,'<sheet>'!<col>:<col>,<col><row>)`
- **Aggregation:** Read source sheet, group by key column (B = 笔记ID), sum each referenced column
- **Fill:** For each data row, replace formula cell with aggregated value; fill 0 if key not found
- **Boundary:** If no SUMIFS formulas detected → return immediately (no-op)
- **Error:** If referenced sheet doesn't exist → log warning, fill 0 for all affected cells

**2. Division formula resolution**
- **Detection:** Scan row 2 for cells matching `=<COL><ROW>/<COL><ROW>` (e.g. `=I2/L2`)
- **Compute:** After SUMIFS values are filled, calculate numerator/denominator for each row
- **Boundary:** Division by zero → fill 0
- **Error:** Non-numeric values after SUMIFS resolution → fill 0

**3. Empty row filtering**
- **Criteria:** `ws.cell(row, 1).value is None and ws.cell(row, 2).value is None`
- **Implementation:** Iterate from max_row down to 2, delete matching rows
- **Boundary:** All rows empty → sheet left with header only; no empty rows → no-op

**4. process_file integration**
- Change `load_workbook(dst, data_only=True)` → `load_workbook(dst, data_only=False)`
- Call `resolve_formulas(wb, ws)` before `apply_calculated_columns(ws)`
- Call `remove_empty_rows(ws)` before `apply_calculated_columns(ws)`

### Non-functional requirements
- Backward compatible: pure-value files process identically to current behavior
- No new dependencies: uses only openpyxl (already a dependency)
- Other sheets' formulas must be preserved in output file

---

## Part 2: Internal Interface Specification

### Function signatures

| # | Function | Module | Description |
|---|----------|--------|-------------|
| 1 | `resolve_formulas(wb, ws)` | processor.py | Detect and resolve all formula cells in ws |
| 2 | `remove_empty_rows(ws)` | processor.py | Delete rows where col 1 and col 2 are both None |

### resolve_formulas(wb, ws) details

**Parameters:**
- `wb: Workbook` — the full workbook (needed to access source sheets)
- `ws: Worksheet` — the target worksheet (Sheet 4)

**Returns:** `None` (modifies ws in-place)

**Algorithm:**
1. Scan row 2 cols 1..max_column for SUMIFS pattern → build `{col_idx: (sheet_name, sum_col_letter)}` map
2. If map empty → return
3. Extract key column from formula (the `B<row>` part) → determines the match column
4. Try `wb[sheet_name]` → if KeyError, log warning, fill 0 for all formula cells, return
5. Read source sheet: for each row, get key value from col B, accumulate sums by key
6. For each row 2..max_row in ws: look up key, fill each SUMIFS col with aggregated value (or 0)
7. Scan row 2 for division formulas `=<COL><ROW>/<COL><ROW>` → parse numerator/denominator col indices
8. For each row: compute division (or 0 on div-by-zero)
9. Final sweep: any remaining formula strings (start with `=`) → replace with 0

### remove_empty_rows(ws) details

**Parameters:** `ws: Worksheet`
**Returns:** `None` (modifies ws in-place)
**Algorithm:** `for row in range(ws.max_row, 1, -1): if col1 is None and col2 is None: ws.delete_rows(row)`

---

### Task 1: Add formula workbook test fixture

**Type:** Code task

**Context Brief:**
- **Files:** `tests/conftest.py` (modify — add new fixture alongside existing ones)
- **Interface:** `make_formula_workbook(rows, source_rows)` → `Workbook` with Sheet 4 containing formula strings and a source data sheet
- **Dependencies:** Used by all subsequent test tasks; follows pattern of existing `make_sample_workbook` fixture
- **Constraints:** Must create formula STRINGS (not actual Excel formulas that compute), since openpyxl tests operate in-memory without an Excel engine

**Files:**
- Modify: `tests/conftest.py`

- [ ] **Step 1: Write the fixture**

Add a new fixture `make_formula_workbook` to `tests/conftest.py`. It should:
- Accept `rows: list[dict]` for Sheet 4 data and `source_rows: list[dict]` for source sheet data
- Create a Workbook with at least 4 sheets
- Sheet 4 (index 3, title "笔记id") headers: `["笔记标题", "笔记ID", "花费", "展现量", "点击量", "留资人数", "留资成本", "互动成本"]`
- For each row in `rows`: write 笔记标题 and 笔记ID as values; write 花费 through 互动成本 as SUMIFS formula strings like `=SUMIFS('源数据'!C:C,'源数据'!A:A,B{row})`
- Create source data sheet (title "源数据") with headers: `["笔记ID", "日期", "消费", "展现量", "点击量", "留资人数"]`
- Write `source_rows` as data rows

Also add a constant `FORMULA_SHEET4_HEADERS` for the formula workbook headers.

- [ ] **Step 2: Write a smoke test to verify the fixture**

Create `tests/test_formula_resolution.py` with a test that creates a formula workbook and verifies:
- Sheet 4 exists at index 3 with correct headers
- Row 2 花费 cell contains a formula string starting with `=SUMIFS`
- Source data sheet exists with correct headers and data

```python
def test_fixture_creates_formula_workbook(make_formula_workbook):
    wb = make_formula_workbook(
        rows=[{"笔记标题": "Test", "笔记ID": "id1"}],
        source_rows=[{"笔记ID": "id1", "日期": "2026-01-01", "消费": 100.0, "展现量": 500, "点击量": 20, "留资人数": 2}],
    )
    ws = wb.worksheets[3]
    assert ws.cell(1, 3).value == "花费"
    assert isinstance(ws.cell(2, 3).value, str)
    assert ws.cell(2, 3).value.startswith("=SUMIFS")
```

- [ ] **Step 3: Run test**

Run: `cd /Users/allenyao/work_project/auto-excel && .venv/bin/pytest tests/test_formula_resolution.py::test_fixture_creates_formula_workbook -v`
Expected: PASS

- [ ] **Step 4: Commit**

```bash
git add tests/conftest.py tests/test_formula_resolution.py
git commit -m "test: add formula workbook fixture for formula resolution tests"
```

---

### Task 2: Implement SUMIFS formula resolution

**Type:** Code task

**Context Brief:**
- **Files:** `tests/test_formula_resolution.py` (modify — add tests), `src/auto_excel/processor.py` (modify — add resolve_formulas)
- **Interface:** `resolve_formulas(wb: Workbook, ws: Worksheet) -> None` — detects SUMIFS formulas in ws row 2, reads source sheet from wb, aggregates by key, fills values
- **Dependencies:** Uses `re` for regex parsing; uses fixture from Task 1; called by `process_file()`
- **Constraints:** Must handle formula strings (not computed formulas) since openpyxl stores them as strings when `data_only=False`

**Files:**
- Modify: `src/auto_excel/processor.py:1-10` (add imports: `re`, `collections.defaultdict`)
- Modify: `src/auto_excel/processor.py` (add `resolve_formulas` function before `process_file`)
- Test: `tests/test_formula_resolution.py`

- [ ] **Step 1: Write the failing test for SUMIFS resolution**

In `tests/test_formula_resolution.py`, add:

```python
from auto_excel.processor import resolve_formulas

def test_resolve_sumifs_basic(make_formula_workbook):
    """SUMIFS formulas are resolved to aggregated values from source sheet."""
    wb = make_formula_workbook(
        rows=[{"笔记标题": "Note A", "笔记ID": "id1"}, {"笔记标题": "Note B", "笔记ID": "id2"}],
        source_rows=[
            {"笔记ID": "id1", "消费": 50.0,  "展现量": 300, "点击量": 10, "留资人数": 1},
            {"笔记ID": "id1", "消费": 30.0,  "展现量": 200, "点击量": 5,  "留资人数": 2},
            {"笔记ID": "id2", "消费": 100.0, "展现量": 800, "点击量": 40, "留资人数": 3},
        ],
    )
    ws = wb.worksheets[3]
    resolve_formulas(wb, ws)

    # id1: 消费 sum = 80, 展现量 sum = 500, 点击量 sum = 15, 留资人数 sum = 3
    assert ws.cell(2, 3).value == 80.0   # 花费
    assert ws.cell(2, 4).value == 500    # 展现量
    assert ws.cell(2, 5).value == 15     # 点击量
    assert ws.cell(2, 6).value == 3      # 留资人数

    # id2: 消费 sum = 100, 展现量 = 800, 点击量 = 40, 留资人数 = 3
    assert ws.cell(3, 3).value == 100.0
    assert ws.cell(3, 4).value == 800
```

- [ ] **Step 2: Run test to verify it fails**

Run: `.venv/bin/pytest tests/test_formula_resolution.py::test_resolve_sumifs_basic -v`
Expected: FAIL — `ImportError: cannot import name 'resolve_formulas'`

- [ ] **Step 3: Implement resolve_formulas (SUMIFS phase)**

In `src/auto_excel/processor.py`, add `resolve_formulas(wb, ws)`:
1. Add imports: `import re` and `from collections import defaultdict`
2. Define regex: `SUMIFS_RE = re.compile(r"=SUMIFS\('([^']+)'!([A-Z]+):[A-Z]+,'[^']*'![A-Z]+:[A-Z]+,([A-Z]+)\d+\)")`
3. Scan row 2: for each col, match against SUMIFS_RE → extract (sheet_name, sum_col_letter, key_col_letter)
4. Build col_letter_to_idx helper (A→1, B→2, ..., Z→26, AA→27)
5. Try `wb[sheet_name]` → KeyError → log warning, fill 0, return
6. Read source sheet: for each row, accumulate `aggregated[key_id][col_idx] += float(val)`
7. For each target row: look up key → fill each SUMIFS column with aggregated value (or 0)

- [ ] **Step 4: Run test to verify it passes**

Run: `.venv/bin/pytest tests/test_formula_resolution.py -v`
Expected: ALL PASS

- [ ] **Step 5: Write adversarial tests**

Add to `tests/test_formula_resolution.py`:

```python
def test_resolve_sumifs_missing_id(make_formula_workbook):
    """IDs in Sheet 4 not in source data get 0 (matches Excel SUMIFS behavior)."""
    wb = make_formula_workbook(
        rows=[{"笔记标题": "Missing", "笔记ID": "id_missing"}],
        source_rows=[{"笔记ID": "id_other", "消费": 100.0, "展现量": 500, "点击量": 20, "留资人数": 1}],
    )
    ws = wb.worksheets[3]
    resolve_formulas(wb, ws)
    assert ws.cell(2, 3).value == 0  # No match → 0

def test_resolve_sumifs_no_formulas(make_sample_workbook):
    """Pure-value workbook: resolve_formulas is a no-op."""
    wb = make_sample_workbook([{"花费": 100.0, "展现量": 1000, "点击量": 50, "留资人数": 5}])
    ws = wb.worksheets[3]
    resolve_formulas(wb, ws)
    assert ws.cell(2, 3).value == 100.0  # Value unchanged

def test_resolve_sumifs_multiple_source_rows(make_formula_workbook):
    """Multiple source rows for same ID are correctly summed."""
    wb = make_formula_workbook(
        rows=[{"笔记标题": "Multi", "笔记ID": "id1"}],
        source_rows=[
            {"笔记ID": "id1", "消费": 10.0, "展现量": 100, "点击量": 5, "留资人数": 0},
            {"笔记ID": "id1", "消费": 20.0, "展现量": 200, "点击量": 3, "留资人数": 1},
            {"笔记ID": "id1", "消费": 30.0, "展现量": 300, "点击量": 7, "留资人数": 0},
        ],
    )
    ws = wb.worksheets[3]
    resolve_formulas(wb, ws)
    assert ws.cell(2, 3).value == 60.0   # 10+20+30
    assert ws.cell(2, 4).value == 600    # 100+200+300
    assert ws.cell(2, 5).value == 15     # 5+3+7
    assert ws.cell(2, 6).value == 1      # 0+1+0
```

- [ ] **Step 6: Run adversarial tests**

Run: `.venv/bin/pytest tests/test_formula_resolution.py -v`
Expected: ALL PASS

- [ ] **Step 7: Commit**

```bash
git add src/auto_excel/processor.py tests/test_formula_resolution.py
git commit -m "feat: add SUMIFS formula resolution in resolve_formulas"
```

---

### Task 3: Implement division formula resolution and residual cleanup

**Type:** Code task

**Context Brief:**
- **Files:** `src/auto_excel/processor.py` (modify — extend `resolve_formulas` with Phase 2 and residual sweep), `tests/test_formula_resolution.py` (modify — add tests)
- **Interface:** Extends `resolve_formulas(wb, ws)` to also handle `=X<row>/Y<row>` division formulas and clean up any remaining formula strings
- **Dependencies:** Phase 2 runs AFTER Phase 1 (SUMIFS) since division formulas reference SUMIFS-resolved cells
- **Constraints:** Division regex: `=([A-Z]+)\d+/([A-Z]+)\d+`; division by zero → 0; residual formula strings → 0

**Files:**
- Modify: `src/auto_excel/processor.py` (extend `resolve_formulas`)
- Test: `tests/test_formula_resolution.py`

- [ ] **Step 1: Write failing tests for division resolution**

```python
def test_resolve_division_formulas(make_formula_workbook):
    """Division formulas (e.g. 留资成本=花费/留资人数) are computed after SUMIFS."""
    # Need fixture that also creates division formula columns
    # Test: after resolve, 留资成本 = 花费 / 留资人数
    ...

def test_resolve_division_by_zero():
    """Division by zero fills 0."""
    ...

def test_resolve_residual_formula_strings():
    """Any formula string remaining after resolution is replaced with 0."""
    ...
```

The fixture from Task 1 should include division formula columns (留资成本, 互动成本). If not already included, extend the fixture to add `=C{row}/F{row}` (花费/留资人数) for 留资成本 and `=C{row}/G{row}` style for 互动成本.

Key assertions:
- After resolve, 留资成本 cell contains a number (花费/留资人数), not a formula string
- When 留资人数 = 0, 留资成本 = 0
- Any cell still containing `=...` after full resolution → replaced with 0

- [ ] **Step 2: Run tests to verify they fail**

Run: `.venv/bin/pytest tests/test_formula_resolution.py -k "division or residual" -v`
Expected: FAIL

- [ ] **Step 3: Extend resolve_formulas with division phase + residual sweep**

In `resolve_formulas`, after SUMIFS phase, add:

Phase 2 — Division formulas:
1. Define regex: `DIV_RE = re.compile(r"=([A-Z]+)\d+/([A-Z]+)\d+")`
2. Scan row 2: for each col matching DIV_RE → extract (numerator_col_letter, denominator_col_letter)
3. Convert letters to indices
4. For each row 2..max_row: compute `num_val / den_val` (or 0 on ZeroDivisionError/TypeError)
5. Write result

Phase 3 — Residual sweep:
1. For each cell in rows 2..max_row, cols 1..max_column: if value is a string starting with `=` → replace with 0

- [ ] **Step 4: Run tests to verify they pass**

Run: `.venv/bin/pytest tests/test_formula_resolution.py -v`
Expected: ALL PASS

- [ ] **Step 5: Write adversarial tests**

```python
def test_adversarial_division_depends_on_sumifs():
    """Division formula using SUMIFS-resolved values produces correct result."""
    # 花费=80 (from SUMIFS), 留资人数=3 (from SUMIFS) → 留资成本=80/3≈26.67
    ...

def test_adversarial_all_zero_source_data():
    """All source data is 0 → all SUMIFS = 0 → all divisions = 0."""
    ...
```

- [ ] **Step 6: Run all formula resolution tests**

Run: `.venv/bin/pytest tests/test_formula_resolution.py -v`
Expected: ALL PASS

- [ ] **Step 7: Commit**

```bash
git add src/auto_excel/processor.py tests/test_formula_resolution.py
git commit -m "feat: add division formula resolution and residual cleanup"
```

---

### Task 4: Implement remove_empty_rows

**Type:** Code task

**Context Brief:**
- **Files:** `src/auto_excel/processor.py` (modify — add function), `tests/test_empty_rows.py` (create)
- **Interface:** `remove_empty_rows(ws: Worksheet) -> None` — deletes rows where col 1 and col 2 are both None
- **Dependencies:** Called by `process_file()` after `resolve_formulas()` and before `apply_calculated_columns()`
- **Constraints:** Must iterate in reverse (max_row down to 2) to avoid row-number shifting; must never delete row 1 (header)

**Files:**
- Modify: `src/auto_excel/processor.py` (add `remove_empty_rows` function)
- Create: `tests/test_empty_rows.py`

- [ ] **Step 1: Write failing tests**

Create `tests/test_empty_rows.py`:

```python
from openpyxl import Workbook
from auto_excel.processor import remove_empty_rows

def test_remove_empty_rows_basic():
    wb = Workbook()
    ws = wb.active
    ws.append(["标题", "ID", "数据"])      # row 1: header
    ws.append(["Note1", "id1", 100])        # row 2: data
    ws.append([None, None, None])           # row 3: empty
    ws.append(["Note2", "id2", 200])        # row 4: data
    ws.append([None, None, None])           # row 5: empty
    remove_empty_rows(ws)
    assert ws.max_row == 3  # header + 2 data rows
    assert ws.cell(2, 1).value == "Note1"
    assert ws.cell(3, 1).value == "Note2"

def test_remove_empty_rows_no_empty():
    wb = Workbook()
    ws = wb.active
    ws.append(["标题", "ID"])
    ws.append(["A", "id1"])
    ws.append(["B", "id2"])
    remove_empty_rows(ws)
    assert ws.max_row == 3  # unchanged

def test_remove_empty_rows_all_empty():
    wb = Workbook()
    ws = wb.active
    ws.append(["标题", "ID"])
    ws.append([None, None])
    ws.append([None, None])
    remove_empty_rows(ws)
    assert ws.max_row == 1  # header only
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `.venv/bin/pytest tests/test_empty_rows.py -v`
Expected: FAIL — `ImportError: cannot import name 'remove_empty_rows'`

- [ ] **Step 3: Implement remove_empty_rows**

In `src/auto_excel/processor.py`, add:

```python
def remove_empty_rows(ws: Worksheet) -> None:
    """Delete rows where column 1 and column 2 are both None. Row 1 (header) is never deleted."""
    for row in range(ws.max_row, 1, -1):
        if ws.cell(row, 1).value is None and ws.cell(row, 2).value is None:
            ws.delete_rows(row)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `.venv/bin/pytest tests/test_empty_rows.py -v`
Expected: ALL PASS

- [ ] **Step 5: Write adversarial tests**

```python
def test_adversarial_title_but_no_id():
    """Row with title but no ID is kept (not empty)."""
    wb = Workbook()
    ws = wb.active
    ws.append(["标题", "ID"])
    ws.append(["Has Title", None])
    remove_empty_rows(ws)
    assert ws.max_row == 2
    assert ws.cell(2, 1).value == "Has Title"

def test_adversarial_id_but_no_title():
    """Row with ID but no title is kept."""
    wb = Workbook()
    ws = wb.active
    ws.append(["标题", "ID"])
    ws.append([None, "has_id"])
    remove_empty_rows(ws)
    assert ws.max_row == 2
    assert ws.cell(2, 2).value == "has_id"

def test_adversarial_empty_rows_scattered():
    """Empty rows scattered among data rows are all removed, data order preserved."""
    wb = Workbook()
    ws = wb.active
    ws.append(["标题", "ID"])
    ws.append([None, None])       # empty
    ws.append(["A", "id1"])       # data
    ws.append([None, None])       # empty
    ws.append([None, None])       # empty
    ws.append(["B", "id2"])       # data
    ws.append([None, None])       # empty
    remove_empty_rows(ws)
    assert ws.max_row == 3
    assert ws.cell(2, 1).value == "A"
    assert ws.cell(3, 1).value == "B"

def test_adversarial_header_never_deleted():
    """Row 1 is never deleted even if it looks empty."""
    wb = Workbook()
    ws = wb.active
    ws.append([None, None])  # row 1: looks empty but is header
    ws.append(["A", "id1"])
    remove_empty_rows(ws)
    assert ws.max_row == 2  # row 1 kept
```

- [ ] **Step 6: Run adversarial tests**

Run: `.venv/bin/pytest tests/test_empty_rows.py -v`
Expected: ALL PASS

- [ ] **Step 7: Commit**

```bash
git add src/auto_excel/processor.py tests/test_empty_rows.py
git commit -m "feat: add remove_empty_rows to filter out empty data rows"
```

---

### Task 5: Wire into process_file and integration test

**Type:** Code task

**Context Brief:**
- **Files:** `src/auto_excel/processor.py:171-183` (modify — `process_file` function), `tests/test_integration.py` (modify — add formula integration test)
- **Interface:** `process_file(src, dst, on_step)` — change `data_only=True` to `data_only=False`, add `resolve_formulas` and `remove_empty_rows` calls
- **Dependencies:** `resolve_formulas` and `remove_empty_rows` from Tasks 2-4; existing `apply_calculated_columns`, `sort_by_column`, `group_and_merge` unchanged
- **Constraints:** Must preserve on_step callback pattern; existing integration test must still pass; new test must verify formula workbook produces correct numeric output

**Files:**
- Modify: `src/auto_excel/processor.py:171-183`
- Modify: `tests/test_integration.py`

- [ ] **Step 1: Write failing integration test for formula workbook**

Add to `tests/test_integration.py`:

```python
def test_full_pipeline_with_formulas(tmp_dirs, monkeypatch, make_formula_workbook):
    """Full pipeline with formula-based data produces correct numeric output."""
    import auto_excel.config as cfg
    import auto_excel.state as st

    state_file = tmp_dirs["log"] / "processed.json"
    for attr, val in [("RAW_DIR", tmp_dirs["raw"]), ("NEW_DIR", tmp_dirs["new"]),
                      ("LOG_DIR", tmp_dirs["log"]), ("STATE_FILE", state_file)]:
        monkeypatch.setattr(cfg, attr, val)
    monkeypatch.setattr(st, "STATE_FILE", state_file)

    wb = make_formula_workbook(
        rows=[
            {"笔记标题": "High Cost Note",  "笔记ID": "id1"},
            {"笔记标题": "Medium Cost Note", "笔记ID": "id2"},
            {"笔记标题": "Low Cost Note",    "笔记ID": "id3"},
        ],
        source_rows=[
            {"笔记ID": "id1", "消费": 500.0, "展现量": 5000, "点击量": 200, "留资人数": 2},
            {"笔记ID": "id1", "消费": 300.0, "展现量": 3000, "点击量": 100, "留资人数": 1},
            {"笔记ID": "id2", "消费": 200.0, "展现量": 2000, "点击量": 80,  "留资人数": 3},
            {"笔记ID": "id3", "消费": 50.0,  "展现量": 1000, "点击量": 30,  "留资人数": 5},
        ],
    )
    src = tmp_dirs["raw"] / "formula_test.xlsx"
    wb.save(src)

    result = runner.invoke(app, ["on"])
    assert result.exit_code == 0

    out_path = tmp_dirs["new"] / "formula_test.xlsx"
    assert out_path.exists()
    wb_out = load_workbook(out_path)
    ws = wb_out.worksheets[3]

    # Verify formula columns are resolved to numbers, not formula strings or None
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    huafei_col = headers.index("花费") + 1
    # id1 花费 = 500+300 = 800 (resolved from source data, not None/0)
    # After sort by 实际成本 desc, id1 should be first (highest cost)
    first_huafei = ws.cell(2, huafei_col).value
    assert first_huafei is not None
    assert isinstance(first_huafei, (int, float))
    assert float(first_huafei) > 0  # Must not be 0 — the original bug
```

- [ ] **Step 2: Run test to verify it fails**

Run: `.venv/bin/pytest tests/test_integration.py::test_full_pipeline_with_formulas -v`
Expected: FAIL — formula cells not resolved (either 0 or formula strings)

- [ ] **Step 3: Modify process_file**

In `src/auto_excel/processor.py`, change `process_file`:

```python
def process_file(src: Path, dst: Path, on_step=None) -> None:
    if on_step: on_step("正在复制文件...")
    shutil.copy2(src, dst)
    wb = load_workbook(dst, data_only=False)  # Changed from True
    ws = wb.worksheets[3]
    if on_step: on_step("正在解析公式...")
    resolve_formulas(wb, ws)
    if on_step: on_step("正在过滤空行...")
    remove_empty_rows(ws)
    if on_step: on_step("正在计算列...")
    apply_calculated_columns(ws)
    if on_step: on_step("正在排序...")
    sort_by_column(ws, "实际成本")
    if on_step: on_step("正在生成分组统计...")
    group_and_merge(ws)
    wb.save(dst)
```

- [ ] **Step 4: Run all tests**

Run: `.venv/bin/pytest tests/ -v`
Expected: ALL PASS (both old and new integration tests)

- [ ] **Step 5: Write adversarial integration test**

```python
def test_full_pipeline_with_formulas_and_empty_rows(tmp_dirs, monkeypatch, make_formula_workbook):
    """Formula workbook with empty rows: empty rows filtered, formulas resolved, groups correct."""
    # Create workbook with 3 data rows + 2 empty rows (None title + None ID)
    # Verify: output has only 3 rows, no empty rows polluting groups
    ...
```

- [ ] **Step 6: Verify existing integration test still passes**

Run: `.venv/bin/pytest tests/test_integration.py -v`
Expected: ALL PASS (backward compatibility confirmed)

- [ ] **Step 7: Commit**

```bash
git add src/auto_excel/processor.py tests/test_integration.py
git commit -m "feat: wire resolve_formulas and remove_empty_rows into process_file pipeline"
```

---

### Task 6: Edge case tests and missing-sheet handling

**Type:** Code task

**Context Brief:**
- **Files:** `tests/test_formula_resolution.py` (modify — add edge case tests), `src/auto_excel/processor.py` (modify if needed — ensure missing-sheet handling works)
- **Interface:** No new functions; tests verify `resolve_formulas` handles edge cases gracefully
- **Dependencies:** `resolve_formulas` from Task 2-3; `make_formula_workbook` from Task 1
- **Constraints:** Missing source sheet → log warning + fill 0 (not crash); empty source sheet → all 0; None IDs in source → skipped

**Files:**
- Modify: `tests/test_formula_resolution.py`
- Modify: `src/auto_excel/processor.py` (if edge case handling needs fixes)

- [ ] **Step 1: Write edge case tests**

```python
def test_resolve_missing_source_sheet():
    """If the source sheet referenced by SUMIFS doesn't exist, fill 0 and don't crash."""
    wb = Workbook()
    for _ in range(3): wb.create_sheet()
    ws = wb.worksheets[3]
    ws.cell(1, 1).value = "标题"
    ws.cell(1, 2).value = "ID"
    ws.cell(1, 3).value = "花费"
    ws.cell(2, 1).value = "Note"
    ws.cell(2, 2).value = "id1"
    ws.cell(2, 3).value = "=SUMIFS('不存在Sheet'!C:C,'不存在Sheet'!A:A,B2)"
    resolve_formulas(wb, ws)
    assert ws.cell(2, 3).value == 0  # Graceful fallback

def test_resolve_empty_source_sheet(make_formula_workbook):
    """Source sheet exists but has no data rows → all SUMIFS = 0."""
    wb = make_formula_workbook(
        rows=[{"笔记标题": "Note", "笔记ID": "id1"}],
        source_rows=[],
    )
    ws = wb.worksheets[3]
    resolve_formulas(wb, ws)
    assert ws.cell(2, 3).value == 0

def test_resolve_source_with_none_ids(make_formula_workbook):
    """Source rows with None IDs are skipped during aggregation."""
    wb = make_formula_workbook(
        rows=[{"笔记标题": "Note", "笔记ID": "id1"}],
        source_rows=[
            {"笔记ID": "id1", "消费": 100.0, "展现量": 500, "点击量": 20, "留资人数": 1},
            {"笔记ID": None,  "消费": 999.0, "展现量": 999, "点击量": 99, "留资人数": 9},
        ],
    )
    ws = wb.worksheets[3]
    resolve_formulas(wb, ws)
    assert ws.cell(2, 3).value == 100.0  # Only id1 data, not None row

def test_resolve_empty_rows_have_no_formula():
    """Rows with no formula and no value are left as-is (handled by remove_empty_rows later)."""
    wb = Workbook()
    for _ in range(3): wb.create_sheet()
    ws = wb.worksheets[3]
    ws.cell(1, 1).value = "标题"
    ws.cell(1, 2).value = "ID"
    ws.cell(1, 3).value = "花费"
    ws.cell(2, 1).value = "Note"
    ws.cell(2, 2).value = "id1"
    ws.cell(2, 3).value = "=SUMIFS('源'!C:C,'源'!A:A,B2)"
    ws.cell(3, 1).value = None  # empty row
    ws.cell(3, 2).value = None
    ws.cell(3, 3).value = None  # no formula, no value
    src = wb.create_sheet("源")
    src.cell(1, 1).value = "笔记ID"
    src.cell(1, 3).value = "消费"
    src.cell(2, 1).value = "id1"
    src.cell(2, 3).value = 50.0
    resolve_formulas(wb, ws)
    assert ws.cell(2, 3).value == 50.0  # formula resolved
    assert ws.cell(3, 3).value is None  # empty row left as-is
```

- [ ] **Step 2: Run tests**

Run: `.venv/bin/pytest tests/test_formula_resolution.py -v`
Expected: ALL PASS (if missing-sheet handling is already implemented correctly). If any fail, fix `resolve_formulas` accordingly.

- [ ] **Step 3: Fix any failing edge cases in resolve_formulas**

Review test output. If missing-sheet or empty-source tests fail, add the necessary error handling.

- [ ] **Step 4: Run full test suite**

Run: `.venv/bin/pytest tests/ -v`
Expected: ALL PASS

- [ ] **Step 5: Commit**

```bash
git add tests/test_formula_resolution.py src/auto_excel/processor.py
git commit -m "test: add edge case tests for formula resolution (missing sheet, empty source, None IDs)"
```

---

### Task 7: Real-file validation test

**Type:** Code task

**Context Brief:**
- **Files:** `tests/test_formula_resolution.py` (modify — add real-file simulation test)
- **Interface:** No new functions; a comprehensive test that simulates the exact structure of the real `0415 .xlsx` file
- **Dependencies:** All functions from Tasks 2-5
- **Constraints:** Must simulate: 12 sheets, SUMIFS formulas in Sheet 4, source data in Sheet 5, multiple rows per ID, division formulas, empty rows mixed with data rows

**Files:**
- Modify: `tests/test_formula_resolution.py`

- [ ] **Step 1: Write real-file simulation test**

Create a test that builds a workbook matching the real file structure:
- 4+ sheets, Sheet 4 (index 3) with "笔记id" title
- Headers matching real file: 笔记标题, 笔记ID, ... 花费(SUMIFS), 展现量(SUMIFS), ... 留资成本(division), 互动成本(division)
- Mix of data rows (with formulas) and empty rows (None everywhere)
- Source data sheet with multiple rows per 笔记ID
- Some IDs with no source data match

Key assertions:
- After full pipeline (resolve + remove_empty + calculated + sort + group): output has correct row count (only data rows, no empties)
- 花费 values are numeric and match expected SUMIFS sums
- 实际成本 column is sorted descending
- 占比 percentages are based on data rows only (not inflated by empty rows)

```python
def test_real_file_simulation():
    """Simulate the exact structure of 0415.xlsx to verify end-to-end correctness."""
    ...
    # After pipeline: verify
    assert ws.max_row == expected_data_count + 1  # +1 for header
    # Verify no formula strings remain
    for row in range(2, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            val = ws.cell(row, col).value
            assert not (isinstance(val, str) and val.startswith("=")), \
                f"Residual formula at row {row} col {col}: {val}"
```

- [ ] **Step 2: Run the test**

Run: `.venv/bin/pytest tests/test_formula_resolution.py::test_real_file_simulation -v`
Expected: PASS

- [ ] **Step 3: Commit**

```bash
git add tests/test_formula_resolution.py
git commit -m "test: add real-file simulation test matching 0415.xlsx structure"
```

---

### Task 8: Run full test suite and verify backward compatibility

**Type:** Non-code task

**Context Brief:**
- **Files:** All test files (read-only — run only)
- **Dependencies:** All previous tasks must be complete
- **Constraints:** Every existing test must still pass; no regressions

**Files:**
- Read-only: all files in `tests/` and `src/auto_excel/`

- [ ] **Step 1: Run complete test suite**

Run: `.venv/bin/pytest tests/ -v --tb=short`
Expected: ALL tests pass, including all pre-existing tests (backward compatibility)

- [ ] **Step 2: Verify no formula strings in pure-value workbook output**

Run: `.venv/bin/pytest tests/test_integration.py::test_full_pipeline -v`
Expected: PASS — confirms pure-value files still work identically

- [ ] **Step 3: Verify no import errors or warnings**

Run: `.venv/bin/python -c "from auto_excel.processor import resolve_formulas, remove_empty_rows, process_file; print('All imports OK')"`
Expected: `All imports OK`

- [ ] **Step 4: Commit any final fixes if needed**

If any test failed and was fixed, commit the fix.
