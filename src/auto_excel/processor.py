from __future__ import annotations
import logging
import shutil
from pathlib import Path
from typing import Callable
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet


def find_column(ws: Worksheet, name: str) -> int:
    """Return 1-based column index of the header named `name` in row 1."""
    for col in range(1, ws.max_column + 1):
        if ws.cell(1, col).value == name:
            return col
    raise ValueError(f"Column '{name}' not found in row 1")


def insert_calculated_column(
    ws: Worksheet,
    after_col_name: str,
    new_col_name: str,
    formula_fn: Callable[[float, int], float],
) -> None:
    """Insert a new column immediately after `after_col_name`.

    Writes `new_col_name` as header in row 1.
    For each data row (2..max_row), reads the source column value,
    applies formula_fn(val, row), fills result. ZeroDivisionError → 0.
    """
    source_idx = find_column(ws, after_col_name)
    insert_at = source_idx + 1
    ws.insert_cols(insert_at)
    ws.cell(1, insert_at).value = new_col_name
    # max_row was captured before insert (insert_cols doesn't change row count)
    for row in range(2, ws.max_row + 1):
        val = ws.cell(row, source_idx).value or 0
        try:
            result = formula_fn(float(val), row)  # type: ignore[arg-type]
        except (ZeroDivisionError, TypeError, ValueError):
            logging.warning("row %d: cannot compute %s from value %r, using 0", row, new_col_name, val)
            result = 0
        ws.cell(row, insert_at).value = result  # type: ignore[assignment]


def apply_calculated_columns(ws: Worksheet) -> None:
    """Insert 4 calculated columns into Sheet 4 in sequence.

    Re-scans column indices after each insertion to account for shifting.
    """
    # Column 1: 实际花费 = 花费 / 1.136
    insert_calculated_column(ws, "花费", "实际花费", lambda val, row: val / 1.136)

    # Column 2: 点击率 = 点击量 / 展现量
    # Capture 展现量 index now (after 实际花费 insertion shifted columns right of 花费)
    zx_col = find_column(ws, "展现量")
    insert_calculated_column(
        ws, "点击量", "点击率",
        lambda val, row, _zx=zx_col: val / (ws.cell(row, _zx).value or 1)
        if (ws.cell(row, _zx).value or 0) != 0 else 0
    )

    # Column 3: CPC = 实际花费 / 点击量
    # Re-scan after 点击率 insertion
    sjhf_col = find_column(ws, "实际花费")
    jl_col = find_column(ws, "点击量")
    insert_calculated_column(
        ws, "点击率", "CPC",
        lambda val, row, _jf=sjhf_col, _jl=jl_col: (
            (ws.cell(row, _jf).value or 0) / (ws.cell(row, _jl).value or 1)
            if (ws.cell(row, _jl).value or 0) != 0 else 0
        )
    )

    # Column 4: 实际成本 = 实际花费 / 留资人数
    # Re-scan after CPC insertion
    sjhf_col = find_column(ws, "实际花费")
    lzrs_col = find_column(ws, "留资人数")
    insert_calculated_column(
        ws, "留资成本", "实际成本",
        lambda val, row, _jf=sjhf_col, _lzrs=lzrs_col: (
            (ws.cell(row, _jf).value or 0) / (ws.cell(row, _lzrs).value or 1)
            if (ws.cell(row, _lzrs).value or 0) != 0 else 0
        )
    )


def sort_by_column(ws: Worksheet, col_name: str, descending: bool = True) -> None:
    """Sort data rows (row 2 to max_row) by the named column.

    Row 1 (header) is never moved. Entire rows move together.
    """
    sort_col = find_column(ws, col_name)
    max_col = ws.max_column
    max_row = ws.max_row

    # Read all data rows into memory
    data_rows = []
    for row in range(2, max_row + 1):
        row_values = [ws.cell(row, col).value for col in range(1, max_col + 1)]
        data_rows.append(row_values)

    # Sort by the target column; None values always go to the end
    none_rows = [r for r in data_rows if r[sort_col - 1] is None]
    value_rows = [r for r in data_rows if r[sort_col - 1] is not None]
    value_rows.sort(key=lambda r: r[sort_col - 1], reverse=descending)
    data_rows = value_rows + none_rows

    # Write sorted rows back
    for row_idx, row_values in enumerate(data_rows, start=2):
        for col_idx, value in enumerate(row_values, start=1):
            ws.cell(row_idx, col_idx).value = value


def group_and_merge(ws: Worksheet) -> None:
    """Insert 占比 column, group rows by 实际成本 ranges, merge and fill."""
    # Insert 占比 column after 互动成本
    hd_col = find_column(ws, "互动成本")
    zb_col = hd_col + 1
    ws.insert_cols(zb_col)
    ws.cell(1, zb_col).value = "占比"

    # Read 实际成本 values (column may have shifted after insert — re-scan)
    sjcb_col = find_column(ws, "实际成本")
    max_row = ws.max_row
    total_rows = max_row - 1  # row 1 is header

    if total_rows == 0:
        return

    # Partition into contiguous groups (data is pre-sorted descending)
    # Groups: high (≥90), medium (50-89), low (<50)
    groups = []
    current_range = None
    current_start = None

    for row in range(2, max_row + 1):
        val = ws.cell(row, sjcb_col).value or 0
        if val >= 90:
            rng = "high"
        elif val >= 50:
            rng = "medium"
        else:
            rng = "low"

        if rng != current_range:
            if current_range is not None:
                groups.append((current_range, current_start, row - 1))
            current_range = rng
            current_start = row

    if current_range is not None:
        groups.append((current_range, current_start, max_row))

    # Merge and fill each group
    for _rng, start_row, end_row in groups:
        count = end_row - start_row + 1
        pct = round(count / total_rows * 100)
        text = f"{count}/{pct}%"

        if start_row == end_row:
            # Single row — just write, no merge needed
            ws.cell(start_row, zb_col).value = text
        else:
            ws.merge_cells(
                start_row=start_row, start_column=zb_col,
                end_row=end_row, end_column=zb_col
            )
            ws.cell(start_row, zb_col).value = text


def process_file(src: Path, dst: Path, on_step=None) -> None:
    if on_step: on_step("正在复制文件...")
    shutil.copy2(src, dst)
    wb = load_workbook(dst, data_only=True)
    ws = wb.worksheets[3]
    if on_step: on_step("正在计算列...")
    apply_calculated_columns(ws)
    if on_step: on_step("正在排序...")
    sort_by_column(ws, "实际成本")
    if on_step: on_step("正在生成分组统计...")
    group_and_merge(ws)
    wb.save(dst)
