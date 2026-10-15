"""Charts: the file's charts are drawn over the grid (column, bar, line, pie, area, scatter as lines);
Insert → Chart makes a new one from the selection; right-click a chart to delete it. Saved through
openpyxl: the file's own charts are kept as openpyxl read them, new ones are built from these dicts.

Chart dict: {'id': index in the file (absent = made here), 'kind': 'col'|'bar'|'line'|'pie'|'area'|'scatter',
'title', 'cell': (r, c) top-left, 'end': (r, c) bottom-right or None, 'series': [{'name', 'values', 'categories'}]}
"""
import tkinter as tk
from tkinter import messagebox, ttk

from formula import FormulaError, Parser, Range, col_letters, quote_sheet

OFFICE = ["#4472c4", "#ed7d31", "#a5a5a5", "#ffc000", "#5b9bd5", "#70ad47", "#264478", "#9e480e"]
COLOUR_BLIND = ["#0072b2", "#e69f00", "#009e73", "#cc79a7", "#56b4e9", "#d55e00", "#f0e442", "#000000"]  # Okabe-Ito
PALETTE = list(OFFICE)


def use_palette(name):
    """'office' or 'colour-blind' (Settings): the series colours of charts drawn from now on."""
    PALETTE[:] = COLOUR_BLIND if name == "colour-blind" else OFFICE
KINDS = [("Column", "col"), ("Bar", "bar"), ("Line", "line"), ("Pie", "pie"), ("Area", "area")]


def from_file(info):
    """extras.read() chart info -> chart dicts."""
    out = []
    for i, ch in enumerate(info or []):
        kind = ch["kind"]
        kind = {"bar": "bar" if ch.get("horizontal") else "col", "bar3d": "col", "line3d": "line",
                "pie3d": "pie", "doughnut": "pie", "area3d": "area"}.get(kind, kind)
        out.append({"id": i, "kind": kind if kind in ("col", "bar", "line", "pie", "area", "scatter") else "col",
                    "title": ch.get("title") or "", "cell": ch["cell"], "end": ch.get("end"),
                    "series": [dict(s) for s in ch["series"]]})
    return out


def changed(sheet):
    return [id(c) for c in sheet.charts] != getattr(sheet, "charts_disk", [])


def remember(sheet):
    sheet.charts_disk = [id(c) for c in sheet.charts]


# ---------- values ----------

def resolve(app, ref, sheet_name):
    """'Sheet1!$B$2:$B$5' -> [values] (row by row)."""
    if not ref:
        return []
    try:
        v = app.engine.ev(Parser(ref.lstrip("=")).parse(), sheet_name)
    except (FormulaError, KeyError):
        return []
    if isinstance(v, Range):
        return list(v.flat())
    return [v]


def numbers(vals):
    out = []
    for v in vals:
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            out.append(float(v))
        else:
            try:
                out.append(float(v))
            except (TypeError, ValueError):
                out.append(0.0)
    return out


def label(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


# ---------- drawing ----------

def box_of(g, ch):
    """Chart position on the canvas (x, y, w, h), even when its top-left is scrolled away."""
    r, c = ch["cell"]
    x, y = g.col_x(c), g.row_y(r)
    if ch.get("end"):
        er, ec = ch["end"]
        w, h = g.col_x(ec) - x, g.row_y(er) - y
    else:
        w, h = g.px(480), g.px(288)
    return x, y, max(w, g.px(120)), max(h, g.px(80))


def nice_scale(lo, hi, ticks=5):
    """Axis from lo to hi on round steps (1, 2, 2.5, 5 x 10^n): -> (axis lo, axis hi, step)."""
    import math
    if hi == lo:
        hi = lo + 1
    raw = (hi - lo) / max(ticks - 1, 1)
    mag = 10 ** math.floor(math.log10(raw))
    step = next(m * mag for m in (1, 2, 2.5, 5, 10) if m * mag >= raw)
    return math.floor(lo / step) * step, math.ceil(hi / step) * step, step


def axis_label(v):
    """54800 -> '54.8k', 1500000 -> '1.5M', 0.25 -> '0.25'."""
    a = abs(v)
    for size, unit in ((1e9, "B"), (1e6, "M"), (1e3, "k")):
        if a >= size:
            return f"{v / size:.3g}{unit}"
    return f"{v:.3g}"


def draw_all(g, C):
    g.chart_boxes = []
    sheet = g.sheet
    if not sheet.charts:
        return
    name = g.app.sheet
    W, H = g.cv.winfo_width(), g.cv.winfo_height()
    for ch in sheet.charts:
        x, y, w, h = box_of(g, ch)
        if x > W or y > H or x + w < 0 or y + h < 0:
            continue
        draw(g.cv, x, y, w, h, ch, g.app, name, g.base_font.actual("family"), C)
        g.chart_boxes.append(((x, y, x + w, y + h), ch))
        if getattr(g, "selected_chart", None) is ch:     # picked: dashed frame + resize handle
            g.cv.create_rectangle(x - 2, y - 2, x + w + 2, y + h + 2, outline=C["accent"], width=2, dash=(4, 3),
                                  tags="chart")
            g.cv.create_rectangle(x + w - 5, y + h - 5, x + w + 5, y + h + 5, fill=C["accent"],
                                  outline=C["cell_bg"], tags="chart")


def draw(cv, x, y, w, h, ch, app, sheet_name, family, C):
    tag = "chart"
    cv.create_rectangle(x, y, x + w, y + h, fill=C["cell_bg"], outline=C["header_line"], tags=tag)
    top = y + 8
    if ch.get("title"):
        cv.create_text(x + w / 2, y + 14, text=ch["title"], font=(family, 11, "bold"), fill=C["cell_fg"], tags=tag)
        top = y + 28
    series = [(s, numbers(resolve(app, s.get("values"), sheet_name))) for s in ch["series"]]
    series = [(s, v) for s, v in series if v]
    if not series:
        cv.create_text(x + w / 2, y + h / 2, text="(no data)", fill=C["muted"], tags=tag)
        return
    cats = [label(v) for v in resolve(app, series[0][0].get("categories"), sheet_name)] or \
        [str(i + 1) for i in range(len(series[0][1]))]
    names = []
    for i, (s, _) in enumerate(series):
        n = resolve(app, s.get("name"), sheet_name) if s.get("name") and "!" in str(s.get("name")) else [s.get("name")]
        names.append(label(n[0]) if n and n[0] not in (None, "") else f"Series {i + 1}")
    left, right, bottom = x + 44, x + w - 12, y + h - 30
    legend_h = 18 if len(series) > 1 or ch["kind"] == "pie" else 0
    bottom -= legend_h
    if ch["kind"] == "pie":
        vals = series[0][1]
        total = sum(v for v in vals if v > 0) or 1
        size = min(right - left, bottom - top) - 4
        cx, cy = (left + right) / 2 - 20, (top + bottom) / 2
        start = 90.0
        for i, v in enumerate(vals):
            ext = -360 * max(v, 0) / total
            cv.create_arc(cx - size / 2, cy - size / 2, cx + size / 2, cy + size / 2, start=start, extent=ext,
                          fill=PALETTE[i % len(PALETTE)], outline=C["cell_bg"], tags=tag)
            start += ext
        names = cats
    else:
        allv = [v for _, vals in series for v in vals]
        lo, hi, step = nice_scale(min(0.0, min(allv)), max(0.0, max(allv)))
        n = max(len(v) for _, v in series)
        horizontal = ch["kind"] == "bar"

        def vy(v):
            return bottom - (v - lo) / (hi - lo) * (bottom - top)

        def vx(v):
            return left + (v - lo) / (hi - lo) * (right - left)
        for k in range(round((hi - lo) / step) + 1):   # gridlines + axis numbers at round values
            v = lo + step * k
            if horizontal:
                cv.create_line(vx(v), top, vx(v), bottom, fill=C["grid_line"], tags=tag)
                cv.create_text(vx(v), bottom + 8, text=axis_label(v), font=(family, 7), fill=C["muted"], tags=tag)
            else:
                cv.create_line(left, vy(v), right, vy(v), fill=C["grid_line"], tags=tag)
                cv.create_text(left - 4, vy(v), text=axis_label(v), anchor="e", font=(family, 7), fill=C["muted"],
                               tags=tag)
        slot = ((bottom - top) if horizontal else (right - left)) / max(n, 1)
        for si, (_s, vals) in enumerate(series):
            color = PALETTE[si % len(PALETTE)]
            if ch["kind"] in ("col", "bar"):
                bw = slot * 0.8 / len(series)
                for i, v in enumerate(vals):
                    p = (top if horizontal else left) + i * slot + slot * 0.1 + si * bw
                    if horizontal:
                        cv.create_rectangle(vx(0), p, vx(v), p + bw, fill=color, outline="", tags=tag)
                    else:
                        cv.create_rectangle(p, vy(v), p + bw, vy(0), fill=color, outline="", tags=tag)
            else:
                pts = []
                for i, v in enumerate(vals):
                    pts += [left + (i + 0.5) * slot, vy(v)]
                if ch["kind"] == "area" and len(pts) >= 4:
                    cv.create_polygon(*pts, pts[-2], vy(0), pts[0], vy(0), fill=color, outline="", stipple="gray50", tags=tag)
                if len(pts) >= 4:
                    cv.create_line(*pts, fill=color, width=2, tags=tag)
                for i in range(0, len(pts), 2):
                    cv.create_oval(pts[i] - 2, pts[i + 1] - 2, pts[i] + 2, pts[i + 1] + 2, fill=color, outline="", tags=tag)
        for i, cname in enumerate(cats[:n]):   # category labels
            if horizontal:
                cv.create_text(left - 4, top + (i + 0.5) * slot, text=cname[:10], anchor="e", font=(family, 7),
                               fill=C["cell_fg"], tags=tag)
            elif n <= 24:
                cv.create_text(left + (i + 0.5) * slot, bottom + 8, text=cname[:10], font=(family, 7),
                               fill=C["cell_fg"], tags=tag)
    if legend_h:
        lx = left
        for i, nm in enumerate(names[:8]):
            cv.create_rectangle(lx, y + h - 16, lx + 8, y + h - 8, fill=PALETTE[i % len(PALETTE)], outline="", tags=tag)
            cv.create_text(lx + 11, y + h - 12, text=nm[:14], anchor="w", font=(family, 7), fill=C["cell_fg"], tags=tag)
            lx += 18 + min(len(nm), 14) * 6


def hit(g, x, y):
    """The chart under (x, y) - its resize handle at the corner counts too."""
    import images
    for box, ch in reversed(getattr(g, "chart_boxes", [])):
        x1, y1, x2, y2 = box
        if x1 <= x <= x2 and y1 <= y <= y2 or images.on_corner(box, x, y):
            return ch
    return None


def delete(app, ch):
    s = app.current
    app.apply_mixed([("ch", 0, 0, [c for c in s.charts if c is not ch])])


# ---------- creating ----------

def from_selection(app, kind="col", title=""):
    """Chart dict from the selected block: first column = categories when it's text, first row = names."""
    g = app.table
    r1, c1, r2, c2 = g.sel_data() if g.multi() else g.region(*g.cur)
    if r2 < r1 or c2 < c1:
        return None
    sheet_ref = quote_sheet(app.sheet) + "!"
    from formatting import parse_number
    first_col_text = all(parse_number(g.value(r, c1)) is None for r in range(r1 + 1, r2 + 1) if g.value(r, c1))
    header = any(g.value(r1, c) and parse_number(g.value(r1, c)) is None for c in range(c1, c2 + 1))
    data_r1 = r1 + 1 if header else r1
    cat_col = c1 if first_col_text and c2 > c1 else None
    ref = lambda c, a, b: f"{sheet_ref}${col_letters(c)}${a + 1}:${col_letters(c)}${b + 1}"   # noqa: E731
    series = []
    for c in range(c1 + (1 if cat_col is not None else 0), c2 + 1):
        series.append({"name": f"{sheet_ref}${col_letters(c)}${r1 + 1}" if header else None,
                       "values": ref(c, data_r1, r2),
                       "categories": ref(cat_col, data_r1, r2) if cat_col is not None else None})
    if not series:
        return None
    return {"kind": kind, "title": title, "cell": (r1, c2 + 2), "end": (r1 + 15, c2 + 9), "series": series}


def insert(app, kind="col", title=""):
    ch = from_selection(app, kind, title)
    if ch is None:
        messagebox.showinfo("Chart", "Select the numbers to chart first.", parent=app)
        return None
    app.apply_mixed([("ch", 0, 0, list(app.current.charts) + [ch])])
    return ch


def insert_dialog(app):
    d = tk.Toplevel(app)
    d.title("Insert chart")
    d.transient(app)
    f = ttk.Frame(d, padding=12)
    f.pack(fill="both", expand=True)
    ttk.Label(f, text="Type").grid(row=0, column=0, sticky="w")
    kind = ttk.Combobox(f, values=[k for k, _ in KINDS], state="readonly", width=14)
    kind.current(0)
    kind.grid(row=0, column=1, sticky="w")
    ttk.Label(f, text="Title").grid(row=1, column=0, sticky="w")
    title = tk.StringVar(master=d)
    ttk.Entry(f, textvariable=title, width=28).grid(row=1, column=1, sticky="w")

    def ok():
        chosen, text = dict(KINDS)[kind.get()], title.get().strip()   # read before the window is gone
        d.destroy()
        return insert(app, chosen, text)
    ttk.Button(f, text="Insert", command=ok).grid(row=2, column=1, sticky="e", pady=(10, 0))
    d.ok, d.kind, d.title_var = ok, kind, title
    return d


# ---------- saving ----------

def build(ch):
    from openpyxl.chart import AreaChart, BarChart, LineChart, PieChart, Reference, Series  # noqa: F401
    from openpyxl.chart.data_source import AxDataSource, StrRef
    from openpyxl.chart.series import SeriesLabel
    cls = {"col": BarChart, "bar": BarChart, "line": LineChart, "pie": PieChart, "area": AreaChart}.get(ch["kind"], BarChart)
    chart = cls()
    if ch["kind"] in ("col", "bar"):
        chart.type = "bar" if ch["kind"] == "bar" else "col"
    if ch.get("title"):
        chart.title = ch["title"]
    for s in ch["series"]:
        ser = Series(s["values"].lstrip("="))
        if s.get("name"):
            ser.tx = SeriesLabel(strRef=StrRef(f=s["name"].lstrip("=")))
        if s.get("categories"):
            ser.cat = AxDataSource(strRef=StrRef(f=s["categories"].lstrip("=")))
        chart.series.append(ser)
    if ch.get("end"):
        (r1, c1), (r2, c2) = ch["cell"], ch["end"]
        chart.width, chart.height = max((c2 - c1) * 1.7, 6), max((r2 - r1) * 0.5, 4)
    return chart


def _set_anchor(chart, ch):
    from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
    (r1, c1), end = ch["cell"], ch.get("end") or (ch["cell"][0] + 15, ch["cell"][1] + 8)
    chart.anchor = TwoCellAnchor(_from=AnchorMarker(col=c1, row=r1), to=AnchorMarker(col=end[1], row=end[0]))


def _anchor(ch):
    r, c = ch["cell"]
    return f"{col_letters(c)}{r + 1}"


def save_fn(name, sheet):
    if not changed(sheet):
        return None
    charts = list(sheet.charts)

    def fn(wb):
        ws = wb[name]
        old = list(getattr(ws, "_charts", []))
        ws._charts = []
        for ch in charts:
            if "id" in ch and ch["id"] < len(old):
                chart = old[ch["id"]]
                if ch.get("moved"):                    # moved / resized here: the file's chart keeps the rest
                    _set_anchor(chart, ch)
                ws._charts.append(chart)
            else:
                ws.add_chart(build(ch), _anchor(ch))
    return fn


def write_all(ws, sheet):
    for ch in sheet.charts:
        if "id" not in ch:
            ws.add_chart(build(ch), _anchor(ch))


def after_save(sheet):
    for i, ch in enumerate(sheet.charts):
        ch["id"] = i
    remember(sheet)


# ---------- moving, resizing, editing ----------

def replace(app, old, new):
    s = app.current
    app.apply_mixed([("ch", 0, 0, [new if c is old else c for c in s.charts])])
    app.table.selected_chart = new
    app.table.redraw()
    return new


def place(app, ch, cell, end):
    """Move / resize a chart to cells cell .. end (top-left, bottom-right)."""
    r1, c1 = max(cell[0], 0), max(cell[1], 0)
    r2, c2 = max(end[0], r1 + 2), max(end[1], c1 + 2)
    return replace(app, ch, {**ch, "cell": (r1, c1), "end": (r2, c2), "moved": True})


def size_of(ch):
    (r1, c1), end = ch["cell"], ch.get("end") or (ch["cell"][0] + 15, ch["cell"][1] + 8)
    return end[0] - r1, end[1] - c1


def nudge(app, ch, dr, dc, resize=False):
    """Keyboard: arrows move the selected chart one cell, Shift+arrows make it bigger / smaller."""
    (r1, c1) = ch["cell"]
    h, w = size_of(ch)
    if resize:
        return place(app, ch, (r1, c1), (r1 + max(h + dr, 2), c1 + max(w + dc, 2)))
    return place(app, ch, (r1 + dr, c1 + dc), (r1 + dr + h, c1 + dc + w))


def drop(g, ch, mode, x0, y0, x1, y1):
    """End of a mouse drag on a chart: mode 'move' or 'resize', from (x0, y0) to (x1, y1)."""
    r_from, c_from = g.hit(x0, y0)
    r_to, c_to = g.hit(max(x1, g.hdr_w() + 1), max(y1, g.hdr_h() + 1))
    if min(r_from, c_from, r_to, c_to) < 0:
        return None
    dr, dc = r_to - r_from, c_to - c_from
    if not dr and not dc:
        return None
    return nudge(g.app, ch, dr, dc, resize=mode == "resize")


class EditDialog(tk.Toplevel):
    """Change a chart's type, title and data (the cells it shows)."""

    def __init__(self, app, ch):
        super().__init__(app)
        self.app, self.ch = app, ch
        self.title("Edit chart")
        self.transient(app)
        f = ttk.Frame(self, padding=12)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text="Type").grid(row=0, column=0, sticky="w")
        self.kind = ttk.Combobox(f, values=[k for k, _ in KINDS], state="readonly", width=14)
        self.kind.set(next((k for k, v in KINDS if v == ch["kind"]), "Column"))
        self.kind.grid(row=0, column=1, sticky="w")
        ttk.Label(f, text="Title").grid(row=1, column=0, sticky="w")
        self.title_var = tk.StringVar(master=self, value=ch.get("title") or "")
        ttk.Entry(f, textvariable=self.title_var, width=30).grid(row=1, column=1, sticky="w")
        ttk.Label(f, text="Data (cells)").grid(row=2, column=0, sticky="w")
        self.data = tk.StringVar(master=self, value=data_range(ch))
        ttk.Entry(f, textvariable=self.data, width=30).grid(row=2, column=1, sticky="w")
        ttk.Label(f, text="first column = labels when it's text, first row = series names",
                  foreground="gray").grid(row=3, column=1, sticky="w")
        ttk.Button(f, text="OK", command=self.ok).grid(row=4, column=1, sticky="e", pady=(10, 0))
        self.bind("<Escape>", lambda e: self.destroy())

    def ok(self):
        from rules_dialog import parse_ranges
        rect = parse_ranges(self.data.get().split("!")[-1].replace("$", ""))
        new = {k: v for k, v in self.ch.items() if k not in ("id", "moved")}   # rebuilt from these settings
        new["kind"], new["title"] = dict(KINDS)[self.kind.get()], self.title_var.get().strip()
        if rect and data_range(self.ch) != self.data.get().strip():
            g = self.app.table
            g.select(*rect[0])
            made = from_selection(self.app, new["kind"], new["title"])
            if made:
                new["series"] = made["series"]
        replace(self.app, self.ch, new)
        self.destroy()
        return new


def data_range(ch):
    """'A1:B5' covering a chart's series (names, values, categories)."""
    from formula import parse_ref
    cells = []
    for s in ch["series"]:
        for k in ("name", "values", "categories"):
            ref = s.get(k)
            if ref and "!" in ref:
                try:
                    _, r1, c1, r2, c2 = parse_ref("ref", ref.split("!")[-1])
                except (AttributeError, ValueError):
                    continue
                cells += [(r1, c1), (r2, c2)]
    if not cells:
        return ""
    r1, c1 = min(r for r, _ in cells), min(c for _, c in cells)
    r2, c2 = max(r for r, _ in cells), max(c for _, c in cells)
    return f"{col_letters(c1)}{r1 + 1}:{col_letters(c2)}{r2 + 1}"
