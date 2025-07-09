import pytest
from io import BytesIO

from openpyxl import Workbook, load_workbook


def test_dynamic_array_roundtrip():
    wb = Workbook()
    ws = wb.active
    ws.add_dynamic_array("SEQUENCE(3)", anchor="A1")
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    wb2 = load_workbook(bio)
    ws2 = wb2.active
    assert len(ws2._arrays) == 1
    da = ws2._arrays[0]
    assert da.formula.startswith("SEQUENCE")
    assert ws2[da.anchor].cm == 1
