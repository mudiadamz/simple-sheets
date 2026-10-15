"""Command palette (Ctrl+Shift+P): every ☰ menu command in one list, found by typing; and first-run tips."""
import tkinter as tk
from tkinter import ttk

import settings


def commands(menu, path=()):
    """Walk a tk.Menu: [(label path, menu, index)] for each command / check item (submenus included)."""
    out = []
    end = menu.index("end")
    if end is None:
        return out
    for i in range(end + 1):
        kind = menu.type(i)
        if kind in ("command", "checkbutton", "radiobutton"):
            if str(menu.entrycget(i, "state")) == "disabled":
                continue
            out.append((path + (menu.entrycget(i, "label"),), menu, i))
        elif kind == "cascade":
            sub = menu.entrycget(i, "menu")
            if sub:
                sub_menu = menu.nametowidget(sub)
                try:                                  # menus filled when opened (recent files, Sheet)
                    post = str(sub_menu.cget("postcommand") or "")   # may come back as a Tcl object
                    if post:
                        sub_menu.tk.eval(post)
                except tk.TclError:
                    pass
                out += commands(sub_menu, path + (menu.entrycget(i, "label"),))
    return out


def score(query, label):
    """Every typed word must appear (in order of importance: start of a word > anywhere). None = no match."""
    text = label.casefold()
    total = 0
    for word in query.casefold().split():
        at = text.find(word)
        if at < 0:
            return None
        total += 0 if at == 0 or text[at - 1] in " ›(/" else 1
    return total


class Palette(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title("Command palette")
        self.transient(app)
        self.items = [(" › ".join(p), m, i) for p, m, i in commands(app.app_menu)]
        self.entry = ttk.Entry(self, width=60)
        self.entry.pack(fill="x", padx=8, pady=(8, 4))
        self.list = tk.Listbox(self, height=14, activestyle="dotbox", exportselection=False)
        self.list.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self.entry.bind("<KeyRelease>", lambda e: e.keysym in ("Up", "Down", "Return", "Escape") or self.fill())
        self.entry.bind("<Down>", lambda e: self.move(1))
        self.entry.bind("<Up>", lambda e: self.move(-1))
        self.entry.bind("<Return>", lambda e: self.run())
        self.list.bind("<Double-Button-1>", lambda e: self.run())
        self.bind("<Escape>", lambda e: self.destroy())
        x = app.winfo_rootx() + max((app.winfo_width() - 520) // 2, 0)
        self.geometry(f"+{x}+{app.winfo_rooty() + 60}")
        self.fill()
        self.entry.focus_set()

    def fill(self):
        q = self.entry.get().strip()
        found = [(score(q, label), k, label) for k, (label, _m, _i) in enumerate(self.items)] if q else \
            [(0, k, label) for k, (label, _m, _i) in enumerate(self.items)]
        found = sorted((s, k, label) for s, k, label in found if s is not None)
        self.shown = [k for _, k, _ in found]
        self.list.delete(0, "end")
        for _, _, label in found:
            self.list.insert("end", label)
        if self.shown:
            self.list.selection_set(0)
        return len(self.shown)

    def move(self, step):
        sel = self.list.curselection()
        i = max(0, min((sel[0] if sel else -1) + step, len(self.shown) - 1))
        self.list.selection_clear(0, "end")
        self.list.selection_set(i)
        self.list.see(i)
        return "break"

    def run(self):
        sel = self.list.curselection()
        if not sel:
            return None
        label, menu, i = self.items[self.shown[sel[0]]]
        self.destroy()
        menu.invoke(i)
        return label


def open_palette(app):
    app.table.commit()
    return Palette(app)


# ---------- tips ----------

TIPS = [
    "Ctrl+Shift+P finds any command by typing a few letters of its name.",
    "Ctrl+A selects the block of data around the cursor; press it again for the whole sheet.",
    "Type = to start a formula, then click cells (or use the arrows) to point at them.",
    "Drag the small square at the selection's corner to fill a series: 1, 2, 3… or Mon, Tue…",
    "Right-click a cell for insert / delete rows and columns, paste special and links.",
    "Ctrl+E (flash fill) finishes a column from one or two examples you typed.",
    "View → Page breaks shows where printed pages end; ☰ → Print preview… shows the pages.",
    "Drop a file on the window to open it. Ctrl+T opens a workbook in a tab of this window.",
    "F6 moves between the cells, the formula bar, the name box and the sheet tabs.",
]


class Tips(tk.Toplevel):
    def __init__(self, app, start=0):
        super().__init__(app)
        self.app, self.i = app, start % len(TIPS)
        self.title("Tips")
        self.transient(app)
        self.resizable(False, False)
        f = ttk.Frame(self, padding=14)
        f.pack(fill="both", expand=True)
        self.head = ttk.Label(f, text="", font=("Segoe UI", 10, "bold"))
        self.head.pack(anchor="w")
        self.text = ttk.Label(f, text="", wraplength=380, justify="left")
        self.text.pack(anchor="w", pady=(6, 12))
        self.again = tk.BooleanVar(master=self, value=bool(settings.get("tips_at_start", True)))
        ttk.Checkbutton(f, text="Show tips when starting", variable=self.again,
                        command=lambda: settings.set("tips_at_start", self.again.get())).pack(anchor="w")
        b = ttk.Frame(f)
        b.pack(fill="x", pady=(10, 0))
        ttk.Button(b, text="Close", command=self.close).pack(side="right")
        ttk.Button(b, text="Next", command=lambda: self.go(1)).pack(side="right", padx=4)
        ttk.Button(b, text="Back", command=lambda: self.go(-1)).pack(side="right")
        self.bind("<Escape>", lambda e: self.close())
        self.go(0)

    def go(self, step):
        self.i = (self.i + step) % len(TIPS)
        self.head["text"] = f"Tip {self.i + 1} of {len(TIPS)}"
        self.text["text"] = TIPS[self.i]

    def close(self):
        settings.set("tip_next", (self.i + 1) % len(TIPS))
        self.destroy()


def show_tips(app):
    return Tips(app, int(settings.get("tip_next", 0) or 0))


def first_run(app):
    """At start: the tips window, unless turned off (one new tip each time)."""
    if settings.get("tips_at_start", True):
        return show_tips(app)
    return None
