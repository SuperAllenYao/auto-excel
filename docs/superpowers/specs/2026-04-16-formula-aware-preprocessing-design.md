# Formula-Aware Excel Preprocessing

**Date:** 2026-04-16
**Status:** Approved
**Scope:** processor.py — formula resolution, workbook formula preservation, empty row filtering

## Problem

Three related bugs cause incorrect output when processing xlsx files with formula-based data columns:

### BUG 1: Formula cells read as None/0 (Critical)

`processor.py:174` uses `load_workbook(dst, data_only=True)`. When data columns contain Excel formulas (e.g., `=SUMIFS(...)`) without cached computed values, openpyxl returns `None` for all formula cells. The code then treats `None` as `0`, producing all-zero output.

**Evidence:** Source file `0415 .xlsx` Sheet 4 has 125 rows where columns 9-22 (花费 through 互动成本) are all SUMIFS/division formulas referencing Sheet "日笔记投放数据". With `data_only=True`, all 125 × 14 = 1,750 formula cells read as `None`.

### BUG 2: `data_only=True` destroys all workbook formulas (Severe)

Loading with `data_only=True` then saving replaces ALL formulas in ALL sheets with their cached values (or `None`). This silently destroys:
- 48 VLOOKUP formulas in Sheet "日笔记投放数据"
- 78 calculation formulas in Sheet "分日数据"

### BUG 3: Empty rows pollute grouping statistics (Severe)

Sheet 4 has 405 data rows: 125 with formulas, 280 completely empty. Empty rows produce `实际成本 = 0` → classified as "low" group → group percentages are meaningless (e.g., "405/100%" instead of actual distribution).

## Design

### Architecture Change

Insert a formula resolution layer and empty row filter into the existing pipeline:

```
process_file()
  ├─ shutil.copy2()
  ├─ load_workbook(dst, data_only=False)      ← CHANGED: preserve formulas
  ├─ resolve_formulas(wb, ws)                  ← NEW: formula preprocessing
  ├─ remove_empty_rows(ws)                     ← NEW: filter empty rows
  ├─ apply_calculated_columns(ws)              (unchanged)
  ├─ sort_by_column(ws, "实际成本")             (unchanged)
  ├─ group_and_merge(ws)                       (unchanged)
  └─ wb.save(dst)
```

### New Function: `resolve_formulas(wb, ws)`

**Purpose:** Detect formula cells in the target worksheet, compute their values from source data, and replace formulas with computed numbers.

**Two-phase resolution:**

**Phase 1 — SUMIFS formulas:**

1. Scan row 2 of the worksheet for SUMIFS formulas
2. Parse each formula with regex: `=SUMIFS('<sheet>'!<sum_col>:<sum_col>,'<sheet>'!<key_col>:<key_col>,<match_col><row>)`
3. Extract: source sheet name, sum column letter, match key column letter
4. Read the source sheet data, aggregate by match key (笔记ID in column B):
   - For each unique ID, sum each referenced column
5. For each data row in the target sheet:
   - Look up the row's key value (笔记ID) in the aggregated data
   - Replace the formula cell with the computed sum
   - If no match found → fill 0 (matches Excel SUMIFS behavior)

**Phase 2 — Simple division formulas:**

1. Scan row 2 for `=<col><row>/<col><row>` patterns (e.g., `=I2/L2`)
2. After SUMIFS values are filled, compute each division
3. Division by zero → fill 0

**Safety rules:**
- If no formulas detected → return immediately (backward compatible with pure-value files)
- If referenced source sheet doesn't exist → log warning, fill 0 for all affected cells
- After resolution, scan for any residual formula strings and replace with 0

### New Function: `remove_empty_rows(ws)`

**Purpose:** Delete rows where both column 1 (笔记标题) and column 2 (笔记ID) are `None`.

**Implementation:** Iterate from `max_row` down to row 2, deleting matching rows. Reverse iteration avoids row-number shifting issues.

### Changes to `process_file()`

1. `load_workbook(dst, data_only=True)` → `load_workbook(dst, data_only=False)`
2. Add `resolve_formulas(wb, ws)` call before `apply_calculated_columns`
3. Add `remove_empty_rows(ws)` call before `apply_calculated_columns`
4. Add corresponding `on_step` progress callbacks

### Compatibility Matrix

| File type | Behavior |
|-----------|----------|
| Formulas + no cache | Resolves from source sheet ✓ |
| Formulas + cached values | Still resolves from source (more accurate) ✓ |
| Pure numeric values | No formulas detected → skip, existing flow ✓ |
| Mixed (some formula, some value) | Per-cell detection, only resolve formulas ✓ |

### Formula Patterns (Verified)

**SUMIFS (7 columns):**

| Column | Formula | Source Sheet Col | Aggregation |
|--------|---------|-----------------|-------------|
| 花费 | `=SUMIFS('日笔记投放数据'!D:D,B:B,B<row>)` | D (消费) | SUM by 笔记ID |
| 展现量 | `=SUMIFS('日笔记投放数据'!E:E,B:B,B<row>)` | E (展现量) | SUM by 笔记ID |
| 点击量 | `=SUMIFS('日笔记投放数据'!F:F,B:B,B<row>)` | F (点击量) | SUM by 笔记ID |
| 进线量 | `=SUMIFS('日笔记投放数据'!J:J,B:B,B<row>)` | J (私信进线数) | SUM by 笔记ID |
| 开口人数 | `=SUMIFS('日笔记投放数据'!L:L,B:B,B<row>)` | L (私信开口数) | SUM by 笔记ID |
| 留资人数 | `=SUMIFS('日笔记投放数据'!N:N,B:B,B<row>)` | N (私信留资数) | SUM by 笔记ID |
| 互动量 | `=SUMIFS('日笔记投放数据'!P:P,B:B,B<row>)` | P (互动量) | SUM by 笔记ID |

**Simple division (7 columns):**

| Column | Formula | Meaning |
|--------|---------|---------|
| 进线成本 | `=I<row>/L<row>` | 花费/进线量 |
| 进线率 | `=L<row>/K<row>` | 进线量/点击量 |
| 开口成本 | `=I<row>/O<row>` | 花费/开口人数 |
| 开口率 | `=O<row>/L<row>` | 开口人数/进线量 |
| 留资成本 | `=I<row>/R<row>` | 花费/留资人数 |
| 留资率 | `=R<row>/O<row>` | 留资人数/开口人数 |
| 互动成本 | `=I<row>/U<row>` | 花费/互动量 |

### Edge Cases

| Scenario | Handling |
|----------|----------|
| 笔记ID in Sheet 4 not in source data (10 cases) | Fill 0, matches Excel SUMIFS |
| Source data has 1059 null-ID rows | Skipped during aggregation |
| Division by zero (e.g., 留资人数=0) | Fill 0 |
| Residual formula strings after resolution | Replaced with 0 |
| Referenced source sheet doesn't exist | Log warning, fill 0 |
| File has < 4 sheets | Existing IndexError behavior (unchanged) |

### Files Modified

- `src/auto_excel/processor.py` — add `resolve_formulas()`, `remove_empty_rows()`, modify `process_file()`

### Files NOT Modified

- `cli.py`, `config.py`, `state.py`, `display.py` — no changes needed
- All existing tests remain valid; new tests needed for formula resolution and empty row filtering
