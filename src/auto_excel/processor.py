from __future__ import annotations
from typing import Callable
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
            result = formula_fn(float(val), row)
        except (ZeroDivisionError, TypeError):
            result = 0
        ws.cell(row, insert_at).value = result


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
