import pytest
from openpyxl import Workbook

SHEET4_HEADERS = ["日期", "计划名称", "花费", "展现量", "点击量", "留资人数", "留资成本", "互动成本"]


@pytest.fixture
def make_sample_workbook():
    """Factory fixture: returns a function that creates an in-memory Workbook.

    The returned workbook has 4 sheets; Sheet 4 (index 3) contains
    marketing data with SHEET4_HEADERS in row 1 and the provided rows
    as subsequent data rows.
    """

    def _factory(rows: list[dict]) -> Workbook:
        wb = Workbook()
        # Workbook starts with 1 active sheet; add 3 more to reach 4 total.
        for name in ["Sheet2", "Sheet3", "Sheet4"]:
            wb.create_sheet(name)

        ws = wb.worksheets[3]  # Sheet4

        # Write header row
        for col_idx, header in enumerate(SHEET4_HEADERS, start=1):
            ws.cell(row=1, column=col_idx, value=header)

        # Write data rows
        for row_idx, row_dict in enumerate(rows, start=2):
            for col_idx, header in enumerate(SHEET4_HEADERS, start=1):
                ws.cell(row=row_idx, column=col_idx, value=row_dict.get(header, 0))

        return wb

    return _factory


@pytest.fixture
def tmp_dirs(tmp_path):
    """Fixture that creates temp Raw/, New/, and log/ subdirectories.

    Returns a dict with keys: base, raw, new, log.
    """
    raw = tmp_path / "Raw"
    new = tmp_path / "New"
    log = tmp_path / "log"

    raw.mkdir()
    new.mkdir()
    log.mkdir()

    return {"base": tmp_path, "raw": raw, "new": new, "log": log}
