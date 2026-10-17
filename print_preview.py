"""Print preview window, PDF of this sheet / all sheets / the selection, and page breaks (View → Page breaks
draws where pages end; Sheet → Insert page break adds manual ones, saved in .xlsx as Excel page breaks)."""
import tkinter as tk
from tkinter import filedialog, ttk
from tkinter import font as tkfont

from simplecore.theme import C

import pdf_export

WHAT = ["This sheet", "All sheets", "Selection"]
ORIENTATIONS = {"Automatic": None, "Portrait": "portrait", "Landscape": "landscape"}
FITS = {"Automatic": None, "Actual size": "actual", "Fit all columns on one page": "width",
        "Fit the sheet on one page": "page"}


def print_options():
    """The orientation / scaling chosen last time in the print dialog: {"orientation", "fit"}."""
    from simplecore import settings
    return {"orientation": settings.get("print_orientation"), "fit": settings.get("print_fit")}


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


def pdf_args(app, what, options=None):
    """Arguments for pdf_export.export: which sheets / cells, plus orientation and scaling."""
    opts = dict(print_options() if options is None else options)
    if what == "All sheets":
        opts["sheets"] = [n for n, s in app.sheets.items() if not s.hidden]
    elif what == "Selection":
        opts["area"] = selection_area(app)
    return opts


def plans(app, what, options=None):
    o = print_options() if options is None else options
    kw = {"orientation": o.get("orientation"), "fit": o.get("fit")}
    if what == "All sheets":
        out = [pdf_export.plan(app, n, **kw) for n, s in app.sheets.items() if not s.hidden]
        return [p for p in out if p["rows"]] or out[:1]
    return [pdf_export.plan(app, None, selection_area(app) if what == "Selection" else None, **kw)]


# ---------- the preview window ----------

class Preview(tk.Toplevel):
    def __init__(self, app, what="This sheet"):
        super().__init__(app)
        self.app = app
        self.title("Print")
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
        self.print_btn = ttk.Button(bar, text="Print…", command=self.print_)
        self.print_btn.pack(side="right", padx=4)
        ttk.Button(bar, text="Export PDF…", command=self.export).pack(side="right")
        opts = ttk.Frame(self, padding=(6, 0, 6, 6))           # page setup: orientation, scaling
        opts.pack(fill="x")
        o = print_options()
        ttk.Label(opts, text="Orientation").pack(side="left")
        self.orientation = ttk.Combobox(opts, values=list(ORIENTATIONS), state="readonly", width=11)
        self.orientation.set(next((k for k, v in ORIENTATIONS.items() if v == o["orientation"]), "Automatic"))
        self.orientation.pack(side="left", padx=(4, 14))
        ttk.Label(opts, text="Scaling").pack(side="left")
        self.fit = ttk.Combobox(opts, values=list(FITS), state="readonly", width=26)
        self.fit.set(next((k for k, v in FITS.items() if v == o["fit"]), "Automatic"))
        self.fit.pack(side="left", padx=4)
        for box in (self.orientation, self.fit):
            box.bind("<<ComboboxSelected>>", lambda e: self.options_changed())
        self.cv = tk.Canvas(self, background=C["header_bg"], highlightthickness=0)
        self.cv.pack(fill="both", expand=True)
        self.cv.bind("<Configure>", lambda e: self.draw())
        for key, step in (("<Prior>", -1), ("<Next>", 1), ("<Left>", -1), ("<Right>", 1)):
            self.bind(key, lambda e, s=step: self.go(s))
        self.bind("<Escape>", lambda e: self.destroy())
        self.load()
        self.focus_set()

    def options(self):
        return {"orientation": ORIENTATIONS[self.orientation.get()], "fit": FITS[self.fit.get()]}

    def options_changed(self):
        from simplecore import settings
        o = self.options()
        settings.set("print_orientation", o["orientation"])       # remembered for the next time
        settings.set("print_fit", o["fit"])
        self.load()

    def load(self):
        self.pages = [(p, i) for p in plans(self.app, self.what.get(), self.options())
                      for i in range(len(p["pages"]))]
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
            return self.app.export_pdf(dst, self.what.get(), self.options())
        return None

    def print_(self):
        return self.app.print_sheet(self.what.get(), self.options())


def draw_page(cv, p, i, x0, y0, k):
    """One page on a canvas, from the same layout as the PDF (pdf_export.page_cells / cell_text): fills,
    gridlines, borders, wrapped and clipped text. k = screen pixels per point."""
    band_r, band_c = p["pages"][i]
    style, measure = p["style"], p["measure"]
    cells = pdf_export.page_cells(band_r, band_c, p["widths"], p["heights"], p["merge_at"])
    for _r, _c, (x, top, w, h) in cells:                     # gridlines under the fills, like the PDF
        if p["gridlines"]:
            cv.create_rectangle(x0 + x * k, y0 + top * k, x0 + (x + w) * k, y0 + (top + h) * k, outline="#d9d9d9")
    for r, c, (x, top, w, h) in cells:
        bg = (style(r, c) or {}).get("bg")
        if bg:
            cv.create_rectangle(x0 + x * k, y0 + top * k, x0 + (x + w) * k, y0 + (top + h) * k, fill=bg,
                                outline="#d9d9d9" if p["gridlines"] else bg)
    for r, c, box in cells:
        for x1, y1, x2, y2, lw, dash, color in pdf_export.border_lines(style(r, c), box):
            cv.create_line(x0 + x1 * k, y0 + y1 * k, x0 + x2 * k, y0 + y2 * k, fill=color,
                           width=max(lw * k, 1), dash=dash or None)
    fonts = {}
    for r, c, box in cells:
        t = pdf_export.cell_text(r, c, box, band_c, p["widths"], p["value"], style, p["merge_at"], measure)
        if not t:
            continue
        px = round(t["size"] * k)
        if px < 3:                                            # too small to read: the page shape is enough
            continue
        key = (px, t["bold"], t["italic"])
        f = fonts.get(key)
        if f is None:
            f = fonts[key] = tkfont.Font(root=cv, family="Arial", size=-px, weight="bold" if t["bold"] else "normal",
                                         slant="italic" if t["italic"] else "roman")
        ascent = f.metrics("ascent")
        cx, ctop, room, ch = t["clip"]
        for lx, ly, line in t["lines"]:
            if ly - t["size"] * 0.75 > ctop + ch:             # below the cell: cut off, like the PDF
                break
            width = lambda s: measure(s, t["size"], t["bold"], t["italic"])   # noqa: E731
            if lx < cx:                                       # starts left of the cell: drop what's outside
                right = lx + width(line)
                while line and width(line) > right - cx:
                    line = line[1:]
                lx = right - width(line)
            line = pdf_export.clip(line, cx + room - lx, t["size"], t["bold"], t["italic"], measure, False)
            if line:
                cv.create_text(x0 + lx * k, y0 + ly * k - ascent, text=line, anchor="nw", font=f, fill=t["fg"])


def open_preview(app, what="This sheet"):
    """The print dialog (Ctrl+P, Print…, Export to PDF…): pages shown first, with orientation and scaling."""
    app.table.commit()
    return Preview(app, what)
