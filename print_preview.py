"""Print preview window, PDF of this sheet / all sheets / the selection, and page breaks (View → Page breaks
draws where pages end; Sheet → Insert page break adds manual ones, saved in .xlsx as Excel page breaks)."""
import tkinter as tk
from tkinter import filedialog, ttk

import pdf_export
from formatting import DEFAULT_SIZE, parse_number, render
from theme import C

WHAT = ["This sheet", "All sheets", "Selection"]


# ---------- page breaks ----------

def breaks(app):
    """The current sheet's page starts: (rows, cols) where a new page begins (cached until the sheet changes)."""
    key = (id(app.current), app.hpos, id(app.history[app.hpos - 1]) if app.hpos else 0, app.sheet)
    cache = getattr(app, "_break_cache", None)
    if cache and cache[0] == key:
        return cache[1]
    p = pdf_export.plan(app)
    manual = app.current.props.get("breaks") or {}
    out = (sorted(set(p["row_starts"]) | set(manual.get("rows", ()))),
           sorted(set(p["col_starts"]) | set(manual.get("cols", ()))), manual)
    app._break_cache = (key, out)
    return out


def draw_breaks(grid, cv, ys, cols, W, H):
    """Dashed blue lines on the grid where printed pages end (solid for manual breaks)."""
    rows_at, cols_at, manual = breaks(grid.app)
    for r in rows_at:
        if r in ys:
            y = ys[r][0]
            cv.create_line(0, y, W, y, fill=C["accent"], width=2,
                           dash=() if r in manual.get("rows", ()) else (6, 4), tags="pagebreak")
    xs = {c: x for c, x, _w in cols}
    for c in cols_at:
        if c in xs:
            cv.create_line(xs[c], 0, xs[c], H, fill=C["accent"], width=2,
                           dash=() if c in manual.get("cols", ()) else (6, 4), tags="pagebreak")


def toggle(app, on=None):
    app.page_breaks = (not getattr(app, "page_breaks", False)) if on is None else on
    if hasattr(app, "page_breaks_var"):
        app.page_breaks_var.set(app.page_breaks)
    app.table.redraw()
    return app.page_breaks


def insert_break(app):
    """A manual page break above the cursor's row (and left of its column when it isn't column A)."""
    r, c = app.table.cur
    old = app.current.props.get("breaks") or {}
    new = {"rows": sorted(set(old.get("rows", ())) | ({r} if r else set())),
           "cols": sorted(set(old.get("cols", ())) | ({c} if c else set()))}
    if new != {"rows": list(old.get("rows", ())), "cols": list(old.get("cols", ()))}:
        app.apply_mixed([("sp", 0, "breaks", new if new["rows"] or new["cols"] else None)])
    return new


def remove_breaks(app):
    if app.current.props.get("breaks"):
        app.apply_mixed([("sp", 0, "breaks", None)])


def read(ws):
    rows = sorted(b.id for b in ws.row_breaks.brk if b.id)
    cols = sorted(b.id for b in ws.col_breaks.brk if b.id)
    return {"rows": rows, "cols": cols} if rows or cols else None


def write(ws, val):
    from openpyxl.worksheet.pagebreak import Break, ColBreak, RowBreak
    ws.row_breaks, ws.col_breaks = RowBreak(), ColBreak()
    for r in (val or {}).get("rows", ()):
        ws.row_breaks.append(Break(id=r))
    for c in (val or {}).get("cols", ()):
        ws.col_breaks.append(Break(id=c))


# ---------- PDF of a choice ----------

def selection_area(app):
    r1, c1, r2, c2 = app.table.sel_cells()
    return r1, c1, r2, c2


def pdf_args(app, what):
    if what == "All sheets":
        return {"sheets": [n for n, s in app.sheets.items() if not s.hidden]}
    if what == "Selection":
        return {"area": selection_area(app)}
    return {}


def plans(app, what):
    if what == "All sheets":
        out = [pdf_export.plan(app, n) for n, s in app.sheets.items() if not s.hidden]
        return [p for p in out if p["rows"]] or out[:1]
    return [pdf_export.plan(app, None, selection_area(app) if what == "Selection" else None)]


# ---------- the preview window ----------

class Preview(tk.Toplevel):
    def __init__(self, app, what="This sheet"):
        super().__init__(app)
        self.app = app
        self.title("Print preview")
        self.transient(app)
        self.geometry(f"{int(760 * app.scale)}x{int(900 * app.scale)}")
        bar = ttk.Frame(self, padding=6)
        bar.pack(fill="x")
        self.what = ttk.Combobox(bar, values=WHAT, state="readonly", width=12)
        self.what.set(what)
        self.what.pack(side="left")
        self.what.bind("<<ComboboxSelected>>", lambda e: self.load())
        ttk.Button(bar, text="◀", width=3, command=lambda: self.go(-1)).pack(side="left", padx=(10, 0))
        self.label = ttk.Label(bar, text="", width=14, anchor="center")
        self.label.pack(side="left")
        ttk.Button(bar, text="▶", width=3, command=lambda: self.go(1)).pack(side="left")
        ttk.Button(bar, text="Close", command=self.destroy).pack(side="right")
        ttk.Button(bar, text="Print…", command=self.print_).pack(side="right", padx=4)
        ttk.Button(bar, text="Export PDF…", command=self.export).pack(side="right")
        self.cv = tk.Canvas(self, background=C["header_bg"], highlightthickness=0)
        self.cv.pack(fill="both", expand=True)
        self.cv.bind("<Configure>", lambda e: self.draw())
        for key, step in (("<Prior>", -1), ("<Next>", 1), ("<Left>", -1), ("<Right>", 1)):
            self.bind(key, lambda e, s=step: self.go(s))
        self.bind("<Escape>", lambda e: self.destroy())
        self.load()
        self.focus_set()

    def load(self):
        self.pages = [(p, i) for p in plans(self.app, self.what.get()) for i in range(len(p["pages"]))]
        self.i = 0
        self.draw()

    def go(self, step):
        self.i = max(0, min(self.i + step, len(self.pages) - 1))
        self.draw()

    def draw(self):
        cv = self.cv
        cv.delete("all")
        self.label["text"] = f"Page {self.i + 1} of {len(self.pages)}"
        if not self.pages:
            return
        p, i = self.pages[self.i]
        W, H = max(cv.winfo_width(), 100), max(cv.winfo_height(), 100)
        k = min((W - 40) / p["page_w"], (H - 40) / p["page_h"])        # points -> screen pixels
        ox, oy = (W - p["page_w"] * k) / 2, (H - p["page_h"] * k) / 2
        cv.create_rectangle(ox + 4, oy + 4, ox + p["page_w"] * k + 4, oy + p["page_h"] * k + 4, fill="#888888",
                            outline="")
        cv.create_rectangle(ox, oy, ox + p["page_w"] * k, oy + p["page_h"] * k, fill="#ffffff", outline="#666666")
        draw_page(cv, p, i, ox + p["origin"][0] * k, oy + p["origin"][1] * k, k * p["scale"])
        if len({q["name"] for q, _ in self.pages}) > 1:         # several sheets: say whose page this is
            cv.create_text(ox + 6, oy + 6, text=p["name"], anchor="nw", fill="#999999")

    def export(self):
        named = self.app.path or "Untitled"
        import os
        stem = os.path.splitext(os.path.basename(named))[0]
        dst = filedialog.asksaveasfilename(parent=self, filetypes=[("PDF", "*.pdf")], defaultextension=".pdf",
                                           initialfile=f"{stem}.pdf")
        if dst:
            return self.app.export_pdf(dst, self.what.get())
        return None

    def print_(self):
        return self.app.print_sheet(self.what.get())


def draw_page(cv, p, i, x0, y0, k):
    """One page's cells on a canvas: fills, gridlines, borders, text (approximate, for the preview)."""
    band_r, band_c = p["pages"][i]
    xs, x = {}, 0.0
    for c in band_c:
        xs[c] = x
        x += p["widths"][c]
    ys, y = {}, 0.0
    for r in band_r:
        ys[r] = y
        y += p["heights"][r]
    for r in band_r:
        for c in band_c:
            bx, by = x0 + xs[c] * k, y0 + ys[r] * k
            bw, bh = p["widths"][c] * k, p["heights"][r] * k
            st = p["style"](r, c) or {}
            if st.get("bg") or p["gridlines"]:
                cv.create_rectangle(bx, by, bx + bw, by + bh, fill=st.get("bg") or "",
                                    outline="#d9d9d9" if p["gridlines"] else "")
            for side, (kind, color) in (st.get("border") or {}).items():
                x1, y1, x2, y2 = {"top": (bx, by, bx + bw, by), "bottom": (bx, by + bh, bx + bw, by + bh),
                                  "left": (bx, by, bx, by + bh), "right": (bx + bw, by, bx + bw, by + bh)}[side]
                cv.create_line(x1, y1, x2, y2, fill=color or "#000000", width=2 if "medium" in kind or kind == "thick" else 1)
            t = p["value"](r, c)
            if not t:
                continue
            disp = render(t, st.get("numfmt")) if st.get("numfmt") else t
            size = max(int(float(st.get("size") or DEFAULT_SIZE) * k), 1)
            font = ("Helvetica", -size, ("bold " if st.get("bold") else "") + ("italic" if st.get("italic") else ""))
            align = st.get("align") or ("right" if parse_number(t) is not None else "left")
            anchor, tx = {"right": ("e", bx + bw - 2 * k), "center": ("center", bx + bw / 2)}.get(align, ("w", bx + 2 * k))
            if size >= 3:
                cv.create_text(tx, by + bh / 2, text=disp.split("\n")[0], anchor=anchor, font=font,
                               fill=st.get("fg") or "#000000")


def open_preview(app, what="This sheet"):
    app.table.commit()
    return Preview(app, what)
