from excelstyler.headers import create_header
import openpyxl


def test_header_creation():
    wb = openpyxl.Workbook()
    ws = wb.active
    create_header(ws, ["A", "B"], 1, 1)
    assert ws.cell(1, 1).value == "A"
