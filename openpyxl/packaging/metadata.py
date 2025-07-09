from openpyxl.xml.functions import Element, SubElement, QName
from openpyxl.xml.constants import SHEET_MAIN_NS, XML_NS

XDA_NS = "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"


class MetadataPart:
    """Represents xl/metadata.xml for dynamic arrays."""

    path = "/xl/metadata.xml"
    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"
    rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata"

    def __init__(self, cell_indices=None):
        self.cell_indices = cell_indices or []

    def to_tree(self):
        root = Element("metadata", xmlns=SHEET_MAIN_NS)
        root.set(QName(XML_NS, "xda"), XDA_NS)
        mt = SubElement(root, "metadataTypes", count="1")
        SubElement(mt, "metadataType", name="XLDAPR", cellMeta="1", copy="1", pasteAll="1")
        fm = SubElement(root, "futureMetadata", name="XLDAPR", count="1")
        bk = SubElement(fm, "bk")
        extLst = SubElement(bk, "extLst")
        ext = SubElement(extLst, "ext", uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}")
        SubElement(ext, f"{{{XDA_NS}}}dynamicArrayProperties", fDynamic="1", fCollapsed="0")
        cm = SubElement(root, "cellMetadata", count=str(len(self.cell_indices)))
        bk = SubElement(cm, "bk")
        for idx in self.cell_indices:
            SubElement(bk, "rc", t="1", v=str(idx))
        return root

    @classmethod
    def from_tree(cls, node):
        cm = node.find('{%s}cellMetadata' % SHEET_MAIN_NS)
        indices = []
        if cm is not None:
            bk = cm.find('{%s}bk' % SHEET_MAIN_NS)
            if bk is not None:
                for rc in bk.findall('{%s}rc' % SHEET_MAIN_NS):
                    v = rc.get('v')
                    if v is not None:
                        indices.append(int(v))
        return cls(indices)
