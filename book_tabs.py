"""Several workbooks in one window: ☰ → New tab (Ctrl+T) / Open in new tab…, or Settings → "open files in
tabs". A bar of workbook tabs appears under the toolbar when more than one is open; each workbook keeps its
own sheets, undo history, file state and formula engine."""
import os
from tkinter import filedialog, ttk

from formula import Engine

STATE = ("path", "sheets", "sheet", "names", "disk_names", "base", "history", "hpos", "disk_sheets", "book_props",
         "disk_book_props", "lost_parts", "lost_ok", "template_src", "encrypted_from", "recovered", "recovered_from",
         "disk_stamp", "locked", "disk_asked", "recovery_id", "autosaved", "formula_home", "engine", "drive_file")


def capture(app):
    app.table.commit()
    if app.table.sheet is not None:      # remember the scroll / selection of the sheet on screen
        t = app.table
        t.sheet.view = (t.top, t.left, t.cur, t.anchor, t.ext)
    return {k: getattr(app, k, None) for k in STATE}


def restore(app, state):
    for k, v in state.items():
        setattr(app, k, v)
    app.table.sheet = None
    app.tabs.rebuild()
    app.show(app.sheet)
    app.refresh()


class BookBar(ttk.Frame):
    def __init__(self, app):
        super().__init__(app, padding=(4, 0, 4, 2))
        self.app = app
        self.visible = False

    def rebuild(self):
        for w in self.winfo_children():
            w.destroy()
        books = self.app.books
        if len(books) < 2:
            if self.visible:
                self.pack_forget()
                self.visible = False
            return
        for i, st in enumerate(books):
            live = i == self.app.book_index
            path = (self.app.path if live else st["path"]) or "Untitled"
            name = os.path.basename(path)
            b = ttk.Button(self, text=("● " if live else "") + name, style="Toolbutton",
                           command=lambda i=i: switch(self.app, i))
            b.pack(side="left")
            x = ttk.Button(self, text="×", width=2, style="Toolbutton", command=lambda i=i: close(self.app, i))
            x.pack(side="left", padx=(0, 6))
        if not self.visible:
            self.pack(fill="x", before=self.app.table)
            self.visible = True


def setup(app):
    app.books = [None]          # None = the workbook currently on screen (its state lives on the app)
    app.book_index = 0
    app.book_bar = BookBar(app)


def _store_current(app):
    app.books[app.book_index] = capture(app)


def switch(app, i):
    if i == app.book_index or not 0 <= i < len(app.books):
        return False
    app.autosave()
    _store_current(app)
    app.book_index = i
    restore(app, app.books[i])
    app.books[i] = None
    app.book_bar.rebuild()
    return True


def new_tab(app, path=None):
    """Another workbook in this window (empty, or `path`)."""
    _store_current(app)
    import recovery
    app.books.append(None)
    app.book_index = len(app.books) - 1
    app.engine = Engine()
    app.history, app.hpos, app.base = [], 0, {}
    app.recovery_id, app.autosaved = recovery.new_id(), None
    app.formula_home = None
    app.new_document()
    if path and not app.open_file(path):
        close(app, app.book_index, ask=False)
        return False
    app.book_bar.rebuild()
    return True


def open_in_tab(app, path=None):
    if path is None:
        from simple_sheets import OPEN_TYPES
        path = filedialog.askopenfilename(filetypes=OPEN_TYPES, parent=app)
        if not path:
            return False
    return new_tab(app, path)


def close(app, i, ask=True):
    """Close workbook tab i (asks to save it first)."""
    if i != app.book_index:
        switch(app, i)
    if ask and not app.confirm_discard():
        return False
    import recovery
    recovery.discard(app.recovery_id)
    if len(app.books) == 1:
        app.new_document()
        return True
    del app.books[i]
    app.book_index = min(i, len(app.books) - 1)
    restore(app, app.books[app.book_index])
    app.books[app.book_index] = None
    app.book_bar.rebuild()
    return True


def close_all(app):
    """Window closing: every open workbook gets its 'save changes?' question. False if one was cancelled."""
    for i in range(len(app.books) - 1, -1, -1):
        if i != app.book_index:
            switch(app, i)
        if not app.confirm_discard():
            return False
        import recovery
        recovery.discard(app.recovery_id)
        if len(app.books) > 1:
            del app.books[app.book_index]
            app.book_index = min(app.book_index, len(app.books) - 1)
            restore(app, app.books[app.book_index])
            app.books[app.book_index] = None
    return True


def count(app):
    return len(getattr(app, "books", [None]))

