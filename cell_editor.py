"""Multi-line in-cell editor: a Text widget that answers like an Entry (character positions),
so code written for the one-line formula bar (pointing at cells, reference highlights) works with both."""
import tkinter as tk

from theme import C


class CellEditor(tk.Text):
    def __init__(self, master, font, wrap):
        super().__init__(master, font=font, relief="solid", bd=1, highlightthickness=0, undo=False,
                         wrap="word" if wrap else "none", padx=2, pady=1,
                         background=C["field_bg"], foreground=C["fg"], insertbackground=C["fg"])

    @staticmethod
    def _ix(i):
        """Entry-style position (int or 'end') -> Text index; other Text indexes pass through."""
        if isinstance(i, int):
            return f"1.0+{i}c"
        return "end-1c" if i == "end" else i

    def get(self, *args):
        return super().get(self._ix(args[0]), self._ix(args[1])) if args else super().get("1.0", "end-1c")

    def insert(self, index, chars, *tags):
        super().insert(self._ix(index), chars, *tags)

    def delete(self, first, last=None):
        super().delete(self._ix(first), self._ix(last) if last is not None else None)

    def index(self, i):
        """'insert' etc. -> character offset from the start (like Entry.index)."""
        return len(super().get("1.0", self._ix(i)))

    def icursor(self, i):
        self.mark_set("insert", self._ix(i))

    def select_range(self, first, last):
        self.tag_remove("sel", "1.0", "end")
        self.tag_add("sel", self._ix(first), self._ix(last))

    def lines(self):
        """Lines as displayed (counting word-wrapped ones)."""
        n = self.count("1.0", "end-1c", "displaylines")
        return (n[0] if n else 0) + 1


class FormulaBar(CellEditor):
    """The formula bar: one line that grows (up to MAX_LINES) for a long formula, with each cell reference
    coloured like its outline on the grid."""
    MAX_LINES = 5

    def __init__(self, master, font):
        super().__init__(master, font, wrap=True)
        self.configure(height=1, padx=4, pady=2)
        self.bind("<KeyRelease>", lambda e: self.refresh(), add="+")
        self.bind("<Configure>", lambda e: self.fit(), add="+")
        self.bind("<Tab>", lambda e: "break")            # Tab doesn't type a tab character here

    def insert(self, index, chars, *tags):
        super().insert(index, chars, *tags)
        self.refresh()

    def delete(self, first, last=None):
        super().delete(first, last)
        self.refresh()

    def fit(self):
        try:
            self.configure(height=max(1, min(self.lines(), self.MAX_LINES)))
        except tk.TclError:
            pass

    def refresh(self):
        self.color_refs()
        self.fit()

    def color_refs(self):
        """Same colours as the dashed boxes on the grid (first reference blue, then red, purple…)."""
        from formula import lex, parse_ref
        from grid import REF_COLORS
        for tag in self.tag_names():
            if tag.startswith("ref"):
                self.tag_delete(tag)
        text = self.get()
        if not text.startswith("="):
            return
        colors, pos = {}, 1
        for kind, tok in lex(text[1:]):
            if kind in ("ref", "cols", "rows"):
                try:
                    key = parse_ref(kind, tok)
                except (AttributeError, ValueError):
                    key = None
                if key is not None:
                    color = colors.setdefault(key, REF_COLORS[len(colors) % len(REF_COLORS)])
                    tag = f"ref{list(colors).index(key) + 1}"     # the same reference: the same tag / colour
                    self.tag_configure(tag, foreground=color)
                    super().tag_add(tag, f"1.0+{pos}c", f"1.0+{pos + len(tok)}c")
            pos += len(tok)
