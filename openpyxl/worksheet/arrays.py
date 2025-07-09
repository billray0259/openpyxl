class DynamicArrayAnchor:
    """Represents a dynamic array formula anchor."""

    def __init__(self, formula=None, anchor=None, ref=None, cm=None):
        self.formula = formula
        self.anchor = anchor
        self.ref = ref
        self.cm = cm

    def __repr__(self):
        return f"<DynamicArrayAnchor {self.anchor} {self.formula}>"
