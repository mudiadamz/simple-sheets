"""Export a sheet to PDF (and print it through the PDF). A small PDF writer of its own: cells with their
fills, borders, fonts (as Helvetica), colours, alignment and wrapping; pages split down and across;
too-wide sheets turn landscape and shrink (down to 50 %) to fit fewer pages."""
import os
from tkinter import font as tkfont

from formatting import DEFAULT_SIZE, parse_number, render

A4 = (595.0, 842.0)
MARGIN = 36.0
MIN_SCALE = 0.5
FONTS = {(False, False): "F1", (True, False): "F2", (False, True): "F3", (True, True): "F4"}
BASE = {"F1": "Helvetica", "F2": "Helvetica-Bold", "F3": "Helvetica-Oblique", "F4": "Helvetica-BoldOblique"}


class Measure:
    """Text widths in points, from Arial (same metrics as Helvetica) via Tk."""

    def __init__(self, master):
        self.fonts = {}
        self.master = master

    def __call__(self, text, size, bold=False, italic=False):
        key = (bold, italic)
        f = self.fonts.get(key)
        if f is None:
            f = self.fonts[key] = tkfont.Font(root=self.master, family="Arial", size=-100,
                                              weight="bold" if bold else "normal",
                                              slant="italic" if italic else "roman")
        return f.measure(text) * size / 100


def rgb(hex6):
    return tuple(int(hex6[i:i + 2], 16) / 255 for i in (1, 3, 5))


def esc(text):
    data = text.encode("cp1252", errors="replace")
    return data.replace(b"\\", b"\\\\").replace(b"(", b"\\(").replace(b")", b"\\)").replace(b"\r", b"")


def layout(sheet, value_of, style_of, measure=None):
    """Visible used area -> (rows, cols, col widths pt, row heights pt). Rows without a height set by hand grow
    for big fonts and for wrapped text (like the grid and Excel), up to Excel's 409 pt."""
    rows = [r for r in range(len(sheet.rows))
            if r not in sheet.hidden_rows and r not in sheet.filtered_rows]
    last_r = max((r for r in rows if any(sheet.rows[r]) or any((r, c) in sheet.styles for c in range(5))),
                 default=-1)
    rows = [r for r in rows if r <= last_r]
    ncols = max((max((c for c, t in enumerate(sheet.rows[r]) if t), default=-1) for r in rows), default=-1) + 1
    for (r, c), st in sheet.styles.items():         # filled / bordered cells count as used too
        if st and (st.get("bg") or st.get("border")) and r <= last_r:
            ncols = max(ncols, c + 1)
    for r1, _c1, _r2, c2 in sheet.merges:
        if r1 <= last_r:
            ncols = max(ncols, c2 + 1)
    cols = [c for c in range(ncols) if c not in sheet.hidden_cols]
    widths = {c: (sheet.widths.get(c, 9.43) * 7 + 5) * 0.75 for c in cols}   # Excel chars -> px -> pt
    heights = {}
    for r in rows:
        if r in sheet.heights:
            heights[r] = sheet.heights[r]
        else:
            big = max(((style_of(r, c) or {}).get("size") or DEFAULT_SIZE for c in cols), default=DEFAULT_SIZE)
            heights[r] = max(15.0, big * 1.35)
            if measure is not None:
                heights[r] = max(heights[r], _wrapped_height(sheet, r, cols, widths, value_of, style_of, measure))
    return rows, cols, widths, heights


def _wrapped_height(sheet, r, cols, widths, value_of, style_of, measure):
    """Height a row needs for its wrapped cells (merged blocks excepted)."""
    merged = {(i, j) for r1, c1, r2, c2 in sheet.merges for i in range(r1, r2 + 1) for j in range(c1, c2 + 1)}
    h = 0.0
    for c in cols:
        st = style_of(r, c) or {}
        if not st.get("wrap") or (r, c) in merged:
            continue
        t = value_of(r, c) if c < len(sheet.rows[r]) else ""
        if not t:
            continue
        disp = render(t, st.get("numfmt")) if st.get("numfmt") else t
        size = float(st.get("size") or DEFAULT_SIZE)
        lines = wrap(disp, widths[c] - 4, size, bool(st.get("bold")), bool(st.get("italic")), measure)
        h = max(h, len(lines) * size * 1.2 + 6)
    return min(h, 409.0)


def split(items, size, room, breaks=()):
    """Group consecutive items into pages of at most `room` (always at least one item per page);
    a manual page break before item i (i in breaks) starts a new page there."""
    pages, cur, used = [], [], 0.0
    for i in items:
        if cur and (used + size[i] > room or i in breaks):
            pages.append(cur)
            cur, used = [], 0.0
        cur.append(i)
        used += size[i]
    if cur:
        pages.append(cur)
    return pages


def export(app, dst, sheet_name=None, gridlines=True, sheets=None, area=None, orientation=None, fit=None):
    """Write a sheet (or several: sheets=[names]; area=(r1, c1, r2, c2) = only those cells) to `dst` as a PDF.
    orientation: None (the file's / automatic), "portrait", "landscape"; fit: None (automatic), "actual",
    "width" (all columns on one page wide), "page" (the whole sheet on one page). Returns the number of pages."""
    out = []
    for name in sheets or [sheet_name or app.sheet]:
        p = plan(app, name, area, gridlines, orientation, fit)
        if sheets and len(sheets) > 1 and not p["rows"]:
            continue                                 # an empty sheet among several: no blank page
        out += [(page_stream_of(p, i), p["page_w"], p["page_h"]) for i in range(len(p["pages"]))]
    if not out:
        p = plan(app, app.sheet, None, gridlines, orientation, fit)
        out = [(page_stream_of(p, 0), p["page_w"], p["page_h"])]
    write_pdf(dst, out)
    return len(out)


def page_stream_of(p, i):
    band_r, band_c = p["pages"][i]
    return page_stream(p["sheet"], band_r, band_c, p["widths"], p["heights"], p["value"], p["style"], p["merge_at"],
                       p["measure"], p["scale"], p["page_h"], p["gridlines"], p["origin"])


def plan(app, name=None, area=None, gridlines=True, orientation=None, fit=None):
    """How a sheet is laid out on paper: page size, scale, and each page's rows and columns.
    Used by the PDF, the print preview and the page-break lines on the grid."""
    name = name or app.sheet
    sheet = app.sheets[name]
    eng = app.engine

    def value(r, c):
        t = sheet.rows[r][c] if c < len(sheet.rows[r]) else ""
        return eng.display(sheet, r, c) if len(t) > 1 and t[0] == "=" else t

    def style(r, c):
        st = sheet.styles.get((r, c))
        cf = app.rules.conditional(sheet, r, c)
        if cf:
            cf.pop("bar", None)
            st = {**(st or {}), **cf}
        return st

    measure = Measure(app)
    rows, cols, widths, heights = layout(sheet, value, style, measure)
    page = sheet.props.get("page")                 # the file's page setup (sheet_props.read_page)
    if page is not None:
        gridlines = bool(page.get("gridlines"))    # Excel prints gridlines only when asked to
    area = area or (page or {}).get("area")
    if area:                                       # print area / selection: only those cells
        r1, c1, r2, c2 = area
        r2 = min(r2, max(len(sheet.rows) - 1, r1))   # a whole-column selection stops at the data
        c2 = min(c2, max(max(cols, default=c1), c1))
        for r in range(r1, r2 + 1):
            if r not in heights and r < len(sheet.rows) and r not in sheet.hidden_rows:
                heights[r] = sheet.heights.get(r, 15.0)
        for c in range(c1, c2 + 1):
            if c not in widths and c not in sheet.hidden_cols:
                widths[c] = (sheet.widths.get(c, 9.43) * 7 + 5) * 0.75
        rows = [r for r in sorted(heights) if r1 <= r <= r2]
        cols = [c for c in sorted(widths) if c1 <= c <= c2]
    page = page or {}
    if page.get("margins"):
        left, right, top, bottom = (m * 72 for m in page["margins"])
    else:
        left = right = top = bottom = MARGIN
    titles_r = [r for r in rows if page.get("title_rows") and page["title_rows"][0] <= r <= page["title_rows"][1]]
    titles_c = [c for c in cols if page.get("title_cols") and page["title_cols"][0] <= c <= page["title_cols"][1]]
    total_w = sum(widths[c] for c in cols) or 1.0
    total_h = sum(heights[r] for r in rows) or 1.0
    page_w, page_h = A4
    orient = orientation or page.get("orientation")   # chosen in the print dialog, else the file's page setup
    if orient == "landscape" or (not orient and total_w > page_w - left - right):
        page_w, page_h = A4[1], A4[0]              # landscape (asked for, or the sheet is wide)
    usable_w, usable_h = page_w - left - right, page_h - top - bottom
    fit_pages = {"width": (1, 0), "page": (1, 1)}.get(fit) or (None if fit == "actual" else page.get("fit"))
    if fit == "actual":
        scale = 1.0
    elif fit_pages:                                # fit to N pages wide / M tall (0 = as many as needed)
        fw, fh = fit_pages
        scale = min(1.0, (usable_w * fw / total_w) if fw else 1.0, (usable_h * fh / total_h) if fh else 1.0)
    elif page.get("scale"):
        scale = page["scale"]
    else:
        scale = max(MIN_SCALE, min(1.0, usable_w / total_w))
    title_h = sum(heights[r] for r in titles_r)
    title_w = sum(widths[c] for c in titles_c)
    room_w, room_h = usable_w / scale - title_w, usable_h / scale - title_h
    merge_at = {}
    for m in sheet.merges:
        for r in range(m[0], m[2] + 1):
            for c in range(m[1], m[3] + 1):
                merge_at[(r, c)] = m
    body_r = [r for r in rows if r not in titles_r]
    body_c = [c for c in cols if c not in titles_c]
    breaks = sheet.props.get("breaks") or {}
    pages = []
    col_bands = split(body_c, widths, room_w, set(breaks.get("cols", ()))) or [[]]
    row_bands = split(body_r, heights, room_h, set(breaks.get("rows", ()))) or [[]]
    for col_band in col_bands:
        for row_band in row_bands:
            # print titles: repeated on every page (they're part of the first page anyway)
            pages.append(((titles_r if titles_r else []) + row_band, (titles_c if titles_c else []) + col_band))
    return {"sheet": sheet, "name": name, "rows": rows, "cols": cols, "widths": widths, "heights": heights,
            "value": value, "style": style, "merge_at": merge_at, "measure": measure, "scale": scale,
            "page_w": page_w, "page_h": page_h, "origin": (left, top), "gridlines": gridlines, "pages": pages,
            "row_starts": [b[0] for b in row_bands[1:] if b], "col_starts": [b[0] for b in col_bands[1:] if b]}


def page_cells(rows, cols, widths, heights, merge_at):
    """The boxes on one page, in points from its top-left: [(r, c, (x, top, w, h))]; a merged block is one box
    over its part on this page. Shared by the PDF and the print preview, so both look the same."""
    xs, x = {}, 0.0
    for c in cols:
        xs[c] = x
        x += widths[c]
    ys, y = {}, 0.0
    for r in rows:
        ys[r] = y
        y += heights[r]

    cells = []
    for r in rows:
        for c in cols:
            m = merge_at.get((r, c))
            if not m:
                cells.append((r, c, (xs[c], ys[r], widths[c], heights[r])))
                continue
            rr = [i for i in rows if m[0] <= i <= m[2]]
            cc = [j for j in cols if m[1] <= j <= m[3]]
            if (r, c) == (rr[0], cc[0]):        # first visible cell of the block draws it
                cells.append((m[0], m[1], (xs[cc[0]], ys[rr[0]], sum(widths[j] for j in cc),
                                           sum(heights[i] for i in rr))))
    return cells


def border_lines(st, box):
    """A cell's borders: [(x1, y1, x2, y2, line width, dash, colour)] in page points (y grows down)."""
    x, top, w, h = box
    out = []
    for side, (kind, color) in ((st or {}).get("border") or {}).items():
        lw = 1.5 if kind == "thick" else 1.0 if kind.startswith("medium") else 0.5
        dash = (2, 2) if kind in ("dotted", "hair") else (4, 2) if "dash" in kind.lower() else ()
        x1, y1, x2, y2 = {"top": (x, top, x + w, top), "bottom": (x, top + h, x + w, top + h),
                          "left": (x, top, x, top + h), "right": (x + w, top, x + w, top + h)}[side]
        out.append((x1, y1, x2, y2, lw, dash, color or "#000000"))
    return out


def cell_text(r, c, box, cols, widths, value, style, merge_at, measure):
    """How a cell's text sits on the page, or None: {"lines": [(x, baseline y, text)], "size", "bold",
    "italic", "fg", "underline", "strike", "clip": (x, top, w, h)} — wrapped, clipped and aligned like Excel."""
    t = value(r, c)
    if not t:
        return None
    x, top, w, h = box
    st = style(r, c) or {}
    disp = render(t, st.get("numfmt")) if st.get("numfmt") else t
    size = float(st.get("size") or DEFAULT_SIZE)
    bold, italic = bool(st.get("bold")), bool(st.get("italic"))
    is_num = parse_number(t) is not None
    align = st.get("align") or ("right" if is_num else "left")
    pad = 2.0
    lines = wrap(disp, w - 2 * pad, size, bold, italic, measure) if st.get("wrap") else [disp.split("\n")[0]]
    room = w
    if not st.get("wrap"):
        if align == "left" and (r, c) not in merge_at:   # spill into empty cells on the right
            for c2 in cols[cols.index(c) + 1:]:
                if value(r, c2) or (r, c2) in merge_at:
                    break
                room += widths[c2]
        lines = [clip(lines[0], room - 2 * pad, size, bold, italic, measure, is_num)]
    lead = size * 1.2
    valign = st.get("valign") or "bottom"
    block = lead * len(lines)
    if block > h - pad:                         # more lines than the row holds: show the first ones
        valign = "top"
    first = top + (pad + size if valign == "top" else (h - block) / 2 + size if valign == "center"
                   else h - pad - block + size - 1)
    placed = []
    for i, line in enumerate(lines):
        lw = measure(line, size, bold, italic)
        lx = x + (w - lw) / 2 if align == "center" else x + w - pad - lw if align == "right" else x + pad
        placed.append((lx, first + i * lead, line))
    return {"lines": placed, "size": size, "bold": bold, "italic": italic, "fg": st.get("fg") or "#000000",
            "underline": bool(st.get("underline")), "strike": bool(st.get("strike")), "clip": (x, top, room, h)}


def page_stream(sheet, rows, cols, widths, heights, value, style, merge_at, measure, scale, page_h, gridlines,
                origin=(MARGIN, MARGIN)):
    out = [f"q {scale:.4f} 0 0 {scale:.4f} {origin[0]:.2f} {page_h - origin[1]:.2f} cm".encode()]   # y grows down

    def rect(x, top, w, h, op):
        return f"{x:.2f} {-(top + h):.2f} {w:.2f} {h:.2f} re {op}".encode()

    cells = page_cells(rows, cols, widths, heights, merge_at)
    if gridlines:
        out.append(b"0.85 0.85 0.85 RG 0.5 w")
        out += [rect(*b, "S") for _, _, b in cells]
    for r, c, b in cells:                            # fills
        st = style(r, c) or {}
        if st.get("bg"):
            out.append(("%.3f %.3f %.3f rg " % rgb(st["bg"])).encode() + rect(*b, "f"))
            if gridlines:
                out.append(b"0.85 0.85 0.85 RG 0.5 w " + rect(*b, "S"))
    for r, c, box in cells:                         # borders
        for x1, y1, x2, y2, lw, dash, color in border_lines(style(r, c), box):
            d = "[%d %d] 0 d" % dash if dash else "[] 0 d"
            out.append(("%.3f %.3f %.3f RG %.2f w %s %.2f %.2f m %.2f %.2f l S" %
                        (*rgb(color), lw, d, x1, -y1, x2, -y2)).encode())
    out.append(b"[] 0 d")
    for r, c, box in cells:                         # text
        t = cell_text(r, c, box, cols, widths, value, style, merge_at, measure)
        if not t:
            continue
        x, top, room, h = t["clip"]
        size = t["size"]
        # the text stays inside its cell (or the empty cells it spills into), like Excel
        out.append(("q %.2f %.2f %.2f %.2f re W n" % (x, -(top + h), room, h)).encode())
        out.append(("%.3f %.3f %.3f rg" % rgb(t["fg"])).encode())
        font = FONTS[(t["bold"], t["italic"])]
        for lx, ly, line in t["lines"]:
            lw = measure(line, size, t["bold"], t["italic"])
            out.append(b"BT /" + font.encode() + (" %.2f Tf %.2f %.2f Td (" % (size, lx, -ly)).encode()
                       + esc(line) + b") Tj ET")
            if t["underline"]:
                out.append(("%.2f w %.2f %.2f m %.2f %.2f l S" % (size / 18, lx, -ly - 1.5, lx + lw, -ly - 1.5)).encode())
            if t["strike"]:
                out.append(("%.2f w %.2f %.2f m %.2f %.2f l S" % (size / 18, lx, -ly + size * 0.3, lx + lw,
                                                                 -ly + size * 0.3)).encode())
        out.append(b"Q")                            # end of this cell's clip
    out.append(b"Q")
    return b"\n".join(out)


def clip(text, room, size, bold, italic, measure, is_num):
    if measure(text, size, bold, italic) <= room:
        return text
    if is_num:
        return "#" * max(int(room // measure("#", size, bold, italic)), 1)   # like Excel: too narrow
    while text and measure(text, size, bold, italic) > room:
        text = text[:-1]
    return text


def wrap(text, room, size, bold, italic, measure):
    lines = []
    for para in text.split("\n"):
        line = ""
        for word in para.split(" "):
            test = f"{line} {word}" if line else word
            if measure(test, size, bold, italic) <= room or not line:
                line = test
            else:
                lines.append(line)
                line = word
        lines.append(line)
    return lines


def write_pdf(dst, pages):
    """Minimal PDF 1.4: catalog, page tree, the four Helvetica fonts, one content stream per page.
    pages: [(content stream, width pt, height pt)]."""
    objs = []                                        # object bodies; object number = index + 1

    def add(body):
        objs.append(body)
        return len(objs)

    catalog = add(None)
    pages_id = add(None)
    font_ids = {k: add(f"<< /Type /Font /Subtype /Type1 /BaseFont /{v} /Encoding /WinAnsiEncoding >>".encode())
                for k, v in BASE.items()}
    fonts = " ".join(f"/{k} {i} 0 R" for k, i in font_ids.items())
    kids = []
    for s, page_w, page_h in pages:
        content = add(b"<< /Length %d >>\nstream\n" % len(s) + s + b"\nendstream")
        kids.append(add((f"<< /Type /Page /Parent {pages_id} 0 R /MediaBox [0 0 {page_w:.0f} {page_h:.0f}] "
                         f"/Resources << /Font << {fonts} >> >> /Contents {content} 0 R >>").encode()))
    objs[catalog - 1] = f"<< /Type /Catalog /Pages {pages_id} 0 R >>".encode()
    objs[pages_id - 1] = (f"<< /Type /Pages /Kids [{' '.join(f'{k} 0 R' for k in kids)}] "
                          f"/Count {len(kids)} >>").encode()
    data = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
    offsets = []
    for i, body in enumerate(objs, start=1):
        offsets.append(len(data))
        data += b"%d 0 obj\n" % i + body + b"\nendobj\n"
    xref = len(data)
    data += b"xref\n0 %d\n0000000000 65535 f \n" % (len(objs) + 1)
    for off in offsets:
        data += b"%010d 00000 n \n" % off
    data += b"trailer\n<< /Size %d /Root %d 0 R >>\nstartxref\n%d\n%%%%EOF\n" % (len(objs) + 1, catalog, xref)
    tmp = dst + ".tmp"
    with open(tmp, "wb") as f:
        f.write(bytes(data))
    os.replace(tmp, dst)


def print_file(path):
    """Send a PDF to the default printer through the program registered for PDFs; else just open it."""
    import platform_ui
    return platform_ui.print_path(path)

