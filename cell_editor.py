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
