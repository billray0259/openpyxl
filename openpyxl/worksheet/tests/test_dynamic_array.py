from openpyxl.workbook import Workbook


def test_add_dynamic_array():
    wb = Workbook()
    ws = wb.active
    da = ws.add_dynamic_array("FILTER(A1:A3,B1:B3>0)", "D1")
    assert ws["D1"].cm == 1
    assert da.formula == "FILTER(A1:A3,B1:B3>0)"
