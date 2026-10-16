"""Help: keyboard shortcuts (F1), the user guide (assets/USER_GUIDE.md), About."""
import os
import sys
import tkinter as tk
from tkinter import ttk

from version import APP_NAME, APP_VERSION

SHORTCUTS = [
    ("Files", [("Ctrl+N", "New window"), ("Ctrl+O", "Open"), ("Ctrl+S", "Save"), ("Ctrl+Shift+S", "Save as"),
               ("Ctrl+P", "Print"), ("F10", "Open the ☰ menu")]),
    ("Moving", [("Arrows", "Move"), ("Ctrl+Arrows", "Jump to the edge of the data"), ("Ctrl+Home / End", "First / last cell"),
                ("Page Up / Down", "Scroll a page"), ("Ctrl+G", "Go to a cell, range or name"),
                ("Ctrl+Page Up / Down", "Previous / next sheet")]),
    ("Selecting", [("Shift+Arrows", "Extend the selection"), ("Ctrl+A", "Select the data block, again for all"),
                   ("Ctrl+click", "Add another block"), ("Shift+click", "Extend to the clicked cell")]),
    ("Editing", [("F2", "Edit the cell"), ("Enter / Tab", "Finish and move down / right"), ("Esc", "Cancel"),
                 ("Ctrl+Enter / Alt+Enter", "New line inside the cell"), ("Delete", "Clear the cells"),
                 ("Ctrl+Z / Ctrl+Y", "Undo / redo"), ("Ctrl+C / X / V", "Copy / cut / paste"),
                 ("Ctrl+D / Ctrl+R", "Fill down / right"), ("Ctrl+E", "Flash fill"),
                 ("Alt+↓", "Open a cell's drop-down list"), ("Ctrl+K", "Insert a link"),
                 ("Drag the selection's edge", "Move the cells"),
                 ("Drag a selected column / row header", "Move whole columns / rows")]),
    ("Formulas", [("=", "Start a formula; a list of functions appears"), ("Tab", "Take the suggested function"),
                  ("Arrows (after = ( , or an operator)", "Point at a cell; Shift+arrows for a range"),
                  ("F9", "Recalculate (manual calculation)"), ("Ctrl+`", "Show formulas")]),
    ("Formatting", [("Ctrl+B / I / U", "Bold / italic / underline"),
                    ("Ctrl+Shift+> / <", "Bigger / smaller text"), ("🖌", "Format painter"),
                    ("Alt+Shift+→ / ←", "Group / ungroup rows or columns")]),
    ("Find & view", [("Ctrl+F / Ctrl+H", "Find / replace"), ("F7", "Spelling"),
                     ("Ctrl + / Ctrl - / Ctrl+wheel", "Zoom"), ("F11", "Full screen"), ("F1", "This list")]),
    ("Keyboard only", [("Ctrl+Shift+P", "Command palette: type to find any command"),
                       ("Ctrl+Shift+A", "AI assistant"),
                       ("Ctrl+Alt+5", "Pick charts / pictures (Tab: the next one)"),
                       ("Arrows / Shift+Arrows / Delete", "Move / resize / remove the picked chart or picture"),
                       ("F6 / Shift+F6", "Cells → formula bar → name box → sheet tabs"),
                       ("Shift+F10 / Menu key", "Right-click menu"), ("Shift+F3", "Insert a function")]),
]


GROUPS = {g for g, _ in SHORTCUTS}


def resource(*parts):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, *parts)


def guide_text(name="USER_GUIDE.md"):
    try:
        with open(resource("assets", name), encoding="utf-8") as f:
            return f.read()
    except OSError:
        return "The guide wasn't found."


def shortcuts_text():
    lines = []
    for group, items in SHORTCUTS:
        lines.append(group)
        lines += [f"    {k:<38}{v}" for k, v in items]
        lines.append("")
    return "\n".join(lines)


class HelpWindow(tk.Toplevel):
    def __init__(self, app, title, text):
        super().__init__(app)
        from theme import C
        self.title(title)
        self.geometry(f"{int(640 * app.scale)}x{int(520 * app.scale)}")
        box = tk.Text(self, wrap="word", padx=12, pady=10, relief="flat", font=("Consolas", 10),
                      background=C["field_bg"], foreground=C["fg"])
        sb = ttk.Scrollbar(self, command=box.yview)
        box.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        box.pack(fill="both", expand=True)
        box.tag_configure("h", font=("Segoe UI", 12, "bold"), spacing1=8)
        box.tag_configure("code", background=C["header_bg"], foreground=C["fg"], lmargin1=16, lmargin2=16,
                          wrap="none")                       # code lines stay whole (scroll sideways if needed)
        in_code = False
        for line in text.split("\n"):
            if line.startswith("```"):             # fenced example: shaded, kept as typed
                in_code = not in_code
            elif in_code:
                box.insert("end", line + "\n", "code")
            elif line.startswith("#") or line in GROUPS:
                box.insert("end", line.lstrip("# ") + "\n", "h")
            else:
                box.insert("end", line + "\n")
        box.configure(state="disabled")
        self.text = box
        self.bind("<Escape>", lambda e: self.destroy())


def show(app, what):
    if what == "shortcuts":
        return HelpWindow(app, "Keyboard shortcuts", shortcuts_text())
    if what == "guide":
        return HelpWindow(app, "User guide", guide_text())
    if what == "scripts":
        w = HelpWindow(app, "Script guide", guide_text("SCRIPT_GUIDE.md"))
        w.geometry(f"{int(820 * app.scale)}x{int(640 * app.scale)}")
        return w
    from tkinter import messagebox
    messagebox.showinfo("About", f"{APP_NAME} {APP_VERSION}\n\nA light viewer / editor for spreadsheet files "
                        "(.xlsx, .xlsm, .xlsb, .xls, .ods, .csv).", parent=app)
    return None
