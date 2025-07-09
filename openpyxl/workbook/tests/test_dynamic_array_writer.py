from openpyxl.workbook import Workbook
from openpyxl.writer.excel import save_workbook
from openpyxl.reader.workbook import WorkbookParser
from openpyxl.xml.constants import ARC_WORKBOOK, ARC_WORKBOOK_RELS
from zipfile import ZipFile


def test_metadata_written(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.add_dynamic_array("FILTER(A1:A3,B1:B3>0)", "D1")
    filename = tmp_path / "da.xlsx"
    save_workbook(wb, filename)
    with ZipFile(filename) as z:
        assert "xl/metadata.xml" in z.namelist()
        rels = z.read("xl/_rels/workbook.xml.rels").decode()
        assert "sheetMetadata" in rels
