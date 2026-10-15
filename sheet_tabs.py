"""Sheet tabs at the bottom: + to add, drag to reorder, double-click to rename in place,
right-click for add/rename/delete."""
import tkinter as tk
from tkinter import ttk

import theme

DRAG_START = 6   # px the mouse must move before a press becomes a drag


class SheetTabs(ttk.Frame):
    """app must provide: sheets (dict), sheet (current name), visible_sheets(), show(name),
    add_sheet(), rename_sheet(old, new), delete_sheet(name), move_sheet(name, index),
    hide_sheet(name, hide)."""

    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        self.drag = None      # (name, x at press, moved?)
        self.editor = None    # inline rename Entry
        add = ttk.Button(self, text="+", width=3, style="Icon.Toolbutton", command=app.add_sheet)
        add.pack(side="left", padx=(4, 2))
        from format_bar import Tooltip
        Tooltip(add, "Add sheet")
        self.nb = ttk.Notebook(self, style="Bottom.TNotebook", height=0)
        self.nb.pack(side="left", fill="x", expand=True)
        self.nb.bind("<<NotebookTabChanged>>", self.on_tab)
        self.nb.bind("<Double-1>", self.rename_at)
        self.nb.bind("<Button-3>", self.menu_at)
        self.nb.bind("<ButtonPress-1>", self.on_press, add="+")
        self.nb.bind("<B1-Motion>", self.on_drag, add="+")
        self.nb.bind("<ButtonRelease-1>", self.on_release, add="+")

    def rebuild(self):
        """Tabs = app.sheets, in order; selects app.sheet."""
        self.cancel_rename()
        for t in self.nb.tabs():
            self.nb.forget(t)
        self.swatches = {}
        for name in self.app.visible_sheets():   # hidden sheets get no tab
            color = self.app.sheets[name].props.get("tab")
            lock = "🔒 " if self.app.sheets[name].props.get("protect") else ""
            if color:   # tab colour: a coloured bar left of the name
                img = self.swatches.get(color)
                if img is None:
                    img = self.swatches[color] = tk.PhotoImage(master=self, width=6, height=14)
                    img.put(color, to=(0, 0, 6, 14))
                self.nb.add(ttk.Frame(self.nb, height=0), text=lock + name, image=img, compound="left")
            else:
                self.nb.add(ttk.Frame(self.nb, height=0), text=lock + name)
        self.select(self.app.sheet)

    def select(self, name):
        names = self.names()
        if name in names:
            self.nb.select(names.index(name))

    def names(self):
        return [self.nb.tab(t, "text").removeprefix("🔒 ") for t in self.nb.tabs()]

    def on_tab(self, _=None):
        if self.nb.select():
            name = self.nb.tab(self.nb.select(), "text").removeprefix("🔒 ")
            if name != self.app.sheet and name in self.app.sheets:
                self.app.show(name)

    def index_at(self, x, y):
        try:
            return self.nb.index(f"@{x},{y}")
        except tk.TclError:
            return None

    def tab_at(self, e):
        i = self.index_at(e.x, e.y)
        return self.names()[i] if i is not None else None

    def tab_box(self, name):
        """(x1, x2) of a tab inside the notebook (ttk has no bbox for tabs, so scan for it)."""
        i, y = self.names().index(name), self.nb.winfo_height() // 2
        xs = [x for x in range(0, self.nb.winfo_width(), 2) if self.index_at(x, y) == i]
        return (xs[0], xs[-1] + 2) if xs else (0, 80)

    # ---------- drag to reorder ----------

    def on_press(self, e):
        name = self.tab_at(e)
        self.drag = (name, e.x, False) if name else None

    def on_drag(self, e):
        if not self.drag:
            return
        name, x0, moved = self.drag
        if not moved and abs(e.x - x0) < DRAG_START:
            return
        self.drag = (name, x0, True)
        self.nb.configure(cursor="sb_h_double_arrow")
        i = self.index_at(max(e.x, 1), self.nb.winfo_height() // 2)
        if i is None:   # past the last tab
            i = len(self.names()) - 1 if e.x > 0 else 0
        names = self.names()
        if names.index(name) != i:   # live feedback: move the tab now
            self.nb.insert(i, self.nb.tabs()[names.index(name)])

    def on_release(self, _=None):
        drag, self.drag = self.drag, None
        self.nb.configure(cursor="")
        if drag and drag[2]:
            name = drag[0]
            new_index = self.names().index(name)
            if new_index != self.app.visible_sheets().index(name):
                self.app.move_sheet(name, new_index)
            else:
                self.rebuild()

    # ---------- rename in place ----------

    def rename_at(self, e):
        name = self.tab_at(e)
        if name:
            self.start_rename(name)
        return "break"

    def start_rename(self, name):
        """Edit box placed right over the tab's name."""
        if self.app.sheet_protected(name, "renamed"):
            return
        self.cancel_rename()
        self.app.show(name)
        self.select(name)
        self.update_idletasks()
        x1, x2 = self.tab_box(name)
        e = ttk.Entry(self.nb, justify="center")
        e.insert(0, name)
        e.select_range(0, "end")
        e.place(x=x1, y=0, width=max(x2 - x1, 90), height=self.nb.winfo_height())
        e.focus_set()
        e.old = name
        e.bind("<Return>", lambda ev: self.finish_rename())
        e.bind("<KP_Enter>", lambda ev: self.finish_rename())
        e.bind("<Escape>", lambda ev: self.cancel_rename())
        e.bind("<FocusOut>", lambda ev: self.after_idle(self.focus_left, e))
        self.editor = e

    def focus_left(self, e):
        """Finish when focus moves elsewhere in the app; not when switching to another program."""
        if self.editor is e and self.focus_get() not in (None, e):
            self.finish_rename()

    def finish_rename(self):
        e, self.editor = self.editor, None
        if e is None:
            return "break"
        old, new = e.old, e.get().strip()
        e.destroy()
        if new != old:
            self.app.rename_sheet(old, new)   # warns and keeps the old name if `new` is invalid
        self.app.table.cv.focus_set()
        return "break"

    def cancel_rename(self):
        e, self.editor = self.editor, None
        if e is not None:
            e.destroy()
        return "break"

    # ---------- right-click ----------

    def menu_at(self, e):
        name = self.tab_at(e)
        if not name:
            return
        self.app.show(name)
        self.select(name)
        m = tk.Menu(self, tearoff=False)
        m.add_command(label="Add sheet", command=self.app.add_sheet)
        locked = bool(self.app.sheets[name].props.get("protect"))   # protected: unprotect first
        m.add_command(label="Rename", command=lambda: self.start_rename(name),
                      state="disabled" if locked else "normal")
        several = len(self.app.visible_sheets()) > 1
        m.add_command(label="Delete", command=lambda: self.app.delete_sheet(name),
                      state="normal" if several and not locked else "disabled")
        m.add_separator()
        m.add_command(label="Hide", command=lambda: self.app.hide_sheet(name),
                      state="normal" if several and not locked else "disabled")
        hidden = [n for n, s in self.app.sheets.items() if s.hidden]
        unhide = tk.Menu(m, tearoff=False)
        for n in hidden:
            unhide.add_command(label=n, command=lambda n=n: self.app.hide_sheet(n, False))
        m.add_cascade(label="Unhide", menu=unhide, state="normal" if hidden else "disabled")
        m.add_separator()
        import workbook_tools as wt
        m.add_command(label="Duplicate", command=lambda: wt.duplicate_sheet(self.app, name))
        m.add_command(label="Copy to a new workbook…", command=lambda: wt.copy_to_new_workbook(self.app, name))
        colors = tk.Menu(m, tearoff=False)
        for label, c in wt.TAB_COLORS:
            colors.add_command(label=label, command=lambda c=c: wt.set_tab_color(self.app, name, c))
        colors.add_command(label="Other…", command=lambda: wt.pick_tab_color(self.app, name))
        colors.add_separator()
        colors.add_command(label="No colour", command=lambda: wt.set_tab_color(self.app, name, None))
        m.add_cascade(label="Tab colour", menu=colors)
        if self.app.sheets[name].props.get("protect"):
            m.add_command(label="Unprotect sheet…", command=lambda: wt.unprotect_sheet(self.app))
        else:
            m.add_command(label="Protect sheet…", command=lambda: wt.protect_sheet(self.app))
        self.menu = m   # kept for tests
        import i18n
        i18n.translate_menu(m)
        theme.style_menu(m)
        try:
            m.tk_popup(e.x_root, e.y_root)
        finally:
            m.grab_release()
