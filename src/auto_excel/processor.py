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
        except ZeroDivisionError:
            result = 0
        ws.cell(row, insert_at).value = result
