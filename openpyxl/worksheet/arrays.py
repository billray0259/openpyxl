from dataclasses import dataclass

@dataclass
class DynamicArrayAnchor:
    """Information about a dynamic array formula anchor."""
    formula: str
    anchor: str
    cm: int | None = None


def add_dynamic_array(ws, formula, anchor):
    """Helper to add a dynamic array formula to a worksheet."""
    da = DynamicArrayAnchor(formula=formula, anchor=anchor)
    if not hasattr(ws, "_arrays"):
        ws._arrays = []
    if not hasattr(ws.parent, "_arrays"):
        ws.parent._arrays = []
    da.cm = len(ws.parent._arrays) + 1
    ws._arrays.append(da)
    ws.parent._arrays.append(da)
    cell = ws[anchor]
    from .formula import ArrayFormula
    cell.value = ArrayFormula(ref=anchor, text=f"=_xlfn._xlws.{formula}")
    cell.cm = da.cm
    return da
