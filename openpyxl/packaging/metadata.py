from openpyxl.xml.functions import Element, SubElement
from openpyxl.xml.constants import SHEET_MAIN_NS


class MetadataPart:
    """Representation of xl/metadata.xml for dynamic arrays."""

    path = "/xl/metadata.xml"
    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.metadata+xml"
    rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata"

    def __init__(self, count=0):
        self.count = count

    @classmethod
    def from_tree(cls, node):
        cm = node.find("{%s}cellMetadata" % SHEET_MAIN_NS)
        count = int(cm.get("count", "0")) if cm is not None else 0
        return cls(count=count)

    def to_tree(self, arrays):
        root = Element("metadata", xmlns=SHEET_MAIN_NS)
        mt = SubElement(root, "metadataTypes", count="1")
        SubElement(mt, "metadataType", name="XLDAPR", id="1", cellMeta="1", copy="1", pasteAll="1")
        fm = SubElement(root, "futureMetadata", name="XLDAPR", count="1")
        bk = SubElement(fm, "bk")
        extLst = SubElement(bk, "extLst")
        ext = SubElement(extLst, "ext", uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}")
        SubElement(ext, "{http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray}dynamicArrayProperties", fDynamic="1", fCollapsed="0")
        cm = SubElement(root, "cellMetadata", count=str(len(arrays)))
        bk = SubElement(cm, "bk")
        for _ in arrays:
            SubElement(bk, "rc", t="1", v="0")
        return root
