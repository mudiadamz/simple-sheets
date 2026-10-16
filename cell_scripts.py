"""Cell scripts: a Python script attached to a cell runs when the cell is clicked (a "button" in the sheet).

Kept in App.book_props["scripts"] = {sheet name: {"G2": code}} (an undoable "wp" change) and saved in the
.xlsx as custom document properties (sheet_props.read_book / apply_book), which Excel keeps.
A script from a file runs only after the user allows that file's scripts (remembered per file + scripts);
scripts written here are trusted. The script gets `clicked` ("G2"), plus everything a Tools → Script has."""
import hashlib
import json

import settings
from formula import col_index, col_letters

TRUST_KEY = "trusted_cell_scripts"


def ref_of(r, c):
    return f"{col_letters(c)}{r + 1}"


def parse_ref(ref):
    import re
    m = re.fullmatch(r"\$?([A-Za-z]{1,3})\$?(\d+)", ref.strip())
    if not m:
        raise ValueError(f"Not a cell address: {ref!r}")
    return int(m.group(2)) - 1, col_index(m.group(1).upper())


def all_scripts(app):
    return app.book_props.get("scripts") or {}


def sheet_name(app, sheet=None):
    sheet = sheet or app.current
    return next((n for n, s in app.sheets.items() if s is sheet), app.sheet)


def get(app, r, c, sheet=None):
    """The script on cell (r, c) of the sheet (default: the one on screen), or None."""
    return (all_scripts(app).get(sheet_name(app, sheet)) or {}).get(ref_of(r, c))


def cells(app, sheet=None):
    """{(r, c)} of the sheet's cells that have a script (for drawing their marker)."""
    out = set()
    for ref in all_scripts(app).get(sheet_name(app, sheet)) or {}:
        try:
            out.add(parse_ref(ref))
        except ValueError:
            pass
    return out


def changed(app, name, ref, code):
    """The whole scripts dict with `ref` on sheet `name` set to code (None removes it)."""
    new = {n: dict(d) for n, d in all_scripts(app).items()}
    d = new.setdefault(name, {})
    if code and code.strip():
        d[ref] = code.rstrip() + "\n"
    else:
        d.pop(ref, None)
    if not d:
        new.pop(name, None)
    return new or None


def assign(app, r, c, code, sheet=None):
    """Put a script on a cell (None / "" removes it). One undo step; the app's own scripts are trusted."""
    new = changed(app, sheet_name(app, sheet), ref_of(r, c), code)
    app.apply_cells([("wp", None, 0, "scripts", new)])
    trust(app)
    app.table.redraw()
    return new


def renamed(app, old, new):
    """The scripts dict with sheet `old` called `new`, or None when that sheet has none (App.rename_sheet)."""
    s = all_scripts(app)
    if old not in s:
        return None
    return {(new if k == old else k): v for k, v in s.items()}


# ---------- trust: scripts from a file run only when allowed ----------

def fingerprint(app):
    """The scripts' identity: allowing them once allows exactly this code (any file, Save as keeps it)."""
    return hashlib.sha256(json.dumps(all_scripts(app), sort_keys=True).encode()).hexdigest()


def trusted(app):
    return fingerprint(app) in (settings.get(TRUST_KEY) or []) or getattr(app, "_scripts_allowed", None) == \
        fingerprint(app)


def trust(app, remember=True):
    fp = fingerprint(app)
    app._scripts_allowed = fp
    if remember:
        known = [x for x in settings.get(TRUST_KEY) or [] if x != fp][-199:]
        settings.set(TRUST_KEY, known + [fp])


def ask_trust(app):
    """'Run the scripts in this file?' -> True (allowed and remembered for this file and these scripts)."""
    import os
    from tkinter import messagebox
    n = sum(len(d) for d in all_scripts(app).values())
    name = os.path.basename(app.path) if app.path else "this workbook"
    ok = messagebox.askyesno(
        "Cell scripts",
        f"{name} has {n} cell script{'s' if n != 1 else ''} (clicking a cell runs Python on this computer).\n\n"
        "Only allow them if you trust where the file came from.\n\nAllow this file's scripts?",
        icon="warning", parent=app)
    if ok:
        trust(app)
    return ok


# ---------- running ----------

def run(app, r, c, sheet=None, ask=True):
    """Run the script of cell (r, c). -> (ok, message), or None when there is none / it wasn't allowed."""
    code = get(app, r, c, sheet)
    if not code:
        return None
    if not trusted(app) and not (ask and ask_trust(app)):
        return None
    import scripting
    out = []
    ok, msg = scripting.run(app, code, out=out.append, filename=f"<cell {ref_of(r, c)}>",
                            extra={"clicked": ref_of(r, c)})
    printed = "".join(out).strip()
    if hasattr(app, "status"):                       # what it printed (or did) shows at the bottom right
        app.status["text"] = f"▶ {ref_of(r, c)}: {printed.splitlines()[-1] if printed else msg.splitlines()[0]}"
    if not ok:
        from tkinter import messagebox
        messagebox.showerror(f"Script on {ref_of(r, c)}", msg, parent=app)
    return ok, msg


# ---------- the editor ----------

class Editor:
    """Right click → Cell script…: write / change the script of a cell; Test runs it once."""

    def __init__(self, app, r, c):
        import tkinter as tk
        from tkinter import ttk

        from theme import C
        self.app, self.r, self.c = app, r, c
        self.name = sheet_name(app)
        w = self.win = tk.Toplevel(app)
        w.title(f"Script on {ref_of(r, c)}")
        w.transient(app)
        w.configure(background=C["bg"])
        f = ttk.Frame(w, padding=12)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text=f"Runs when {ref_of(r, c)} is clicked. Same Python as Tools → Script "
                          "(sheet, book, selection, print…); `clicked` is the cell's address.",
                  wraplength=int(520 * getattr(app, "scale", 1))).pack(anchor="w")
        self.text = tk.Text(f, width=72, height=16, font=("Consolas", 10), undo=True, wrap="none",
                            background=C["field_bg"], foreground=C["fg"], insertbackground=C["fg"],
                            relief="flat", highlightthickness=1, highlightbackground=C["header_line"],
                            highlightcolor=C["accent"], padx=6, pady=4)
        self.text.pack(fill="both", expand=True, pady=(8, 8))
        self.text.insert("1.0", get(app, r, c) or EXAMPLE)
        import edit_menu
        edit_menu.attach(self.text)
        b = ttk.Frame(f)
        b.pack(fill="x")
        ttk.Button(b, text="Remove script", command=self.remove).pack(side="left")
        ttk.Button(b, text="Cancel", command=w.destroy).pack(side="right")
        ttk.Button(b, text="Save", command=self.save, default="active").pack(side="right", padx=6)
        ttk.Button(b, text="▶ Test", command=self.test).pack(side="right")
        w.bind("<Escape>", lambda e: w.destroy())
        from ai_assist import center
        center(w, app)
        self.text.focus_set()

    def code(self):
        return self.text.get("1.0", "end-1c")

    def save(self):
        assign(self.app, self.r, self.c, self.code())
        self.win.destroy()
        return True

    def remove(self):
        assign(self.app, self.r, self.c, None)
        self.win.destroy()
        return True

    def test(self):
        import scripting
        ok, msg = scripting.run(self.app, self.code(), filename=f"<cell {ref_of(self.r, self.c)}>",
                                extra={"clicked": ref_of(self.r, self.c)})
        from tkinter import messagebox
        (messagebox.showinfo if ok else messagebox.showerror)("Test", msg, parent=self.win)
        return ok


EXAMPLE = """# Runs when this cell is clicked. Example: jump the calendar to this month.
import datetime
today = datetime.date.today()
# sheet["B2"] = today.year
# sheet["D2"] = today.strftime("%B")
print("Clicked", clicked)
"""


def open_editor(app, r=None, c=None):
    if r is None:
        r, c = app.table.cur
    app.table.commit()
    return Editor(app, r, c)
