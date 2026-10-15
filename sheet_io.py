"""Load / save spreadsheet files. A book is {name: Sheet}."""
import csv
import datetime as dt
import os
import re
from itertools import zip_longest

from formatting import DEFAULT_FONT, DEFAULT_SIZE, NUMBER, SIDES, cell_style, parse_date, theme_colors

MAX_ROWS = 50_000  # keep the UI snappy on huge files
FULL_LOAD_BYTES = 15 * 1024 * 1024  # bigger .xlsx: values only (no column widths / merges)


class Sheet:
    def __init__(self, rows, styles=None, widths=None, merges=None, heights=None):
        self.rows = rows               # [[str, ...], ...] - the text shown / edited
        self.styles = styles or {}     # (r, c) -> style dict (see formatting.cell_style), never mutated
        self.widths = widths or {}     # c -> width in Excel character units
        self.heights = heights or {}   # r -> custom row height in points
        self.cache = {}                # (r, c) -> Excel's saved result of a formula cell
        self.links = {}                # (r, c) -> hyperlink address
        self.comments = {}             # (r, c) -> [text, author]
        self.threads = {}              # (r, c) -> [[author, text, time, id], ...] threaded comments
        self.rich = {}                 # (r, c) -> [[text, font dict], ...] for cells with several fonts
        self.merges = merges or []     # [(r1, c1, r2, c2)], 0-based inclusive
        self.px = {}                   # c -> width set by dragging (view only)
        self.view = None               # remembered scroll/selection
        self.disk_name = None          # this sheet's name in the file on disk (None = new sheet)
        self.freeze = (0, 0)           # frozen (rows, columns) at the top / left
        self.filter = None             # {"range": (r1, c1, r2, c2), "criteria": {col: [values shown]}}
        self.filtered_rows = set()     # rows the filter hides (kept apart from rows hidden by hand)
        self.hidden = False            # whole sheet hidden (no tab)
        self.disk_hidden = False       # ... as in the file on disk
        self.hidden_rows = set()       # row indexes hidden
        self.hidden_cols = set()       # column indexes hidden
        self.validations = []          # data validation rules (rules.read_validations), on-disk ranges
        self.cond = []                 # conditional formats (rules.read_conditional), on-disk ranges
        self.props = {}                # sheet settings (sheet_props.py): grid, headings, tab, protect, page
        self.outline_rows = {}         # r -> outline (group) level
        self.outline_cols = {}         # c -> outline (group) level
        self.extra = {}                # read-only extras from the file: tables, images, charts, shapes
        self.cond_disk, self.dv_disk = [], []   # rules as in the file (rules.remember / rules.changed)
        self.charts = []               # charts (the file's + made here, charts.py), saved through openpyxl
        self.charts_disk = []          # ids of the charts as in the file (charts.remember / charts.changed)
        self.images_disk = []          # same for pictures (images.remember / images.changed)
        self.reset_structure()

    # Row/column maps: model index -> index in the file on disk (None = inserted here).
    # Indexes past the map continue from *tail*. Used to rebuild the sheet on save.
    def reset_structure(self):
        self.rowmap, self.rowtail, self.colmap, self.coltail = [], 0, [], 0
        self.styles0 = dict(self.styles)   # styles as on disk, by disk position
        self.disk_merges = list(self.merges)   # merged ranges as on disk

    def structure_changed(self):
        return any(v != i for i, v in enumerate(self.rowmap)) or self.rowtail != len(self.rowmap) or \
            any(v != i for i, v in enumerate(self.colmap)) or self.coltail != len(self.colmap)

    @staticmethod
    def _orig(m, tail, i):
        return m[i] if i < len(m) else tail + i - len(m)

    @staticmethod
    def _inverse(m, tail):
        inv = {v: i for i, v in enumerate(m) if v is not None}
        return lambda o: inv.get(o) if o < tail else o - tail + len(m)

    def orig_pos(self, r, c):
        return self._orig(self.rowmap, self.rowtail, r), self._orig(self.colmap, self.coltail, c)


def fmt(v):
    if v is None:
        return ""
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else f"{v:.15g}"
    if isinstance(v, dt.datetime):
        return v.strftime("%Y-%m-%d") if v.time() == dt.time() else v.strftime("%Y-%m-%d %H:%M:%S")
    return str(v)


def to_value(s):
    """Grid text -> cell value: numbers/ISO dates become typed values, "" becomes empty."""
    if s == "":
        return None
    if NUMBER.fullmatch(s):
        return float(s) if any(ch in s for ch in ".eE") else int(s)
    d = parse_date(s)
    if d is not None:
        return d.date() if len(s) == 10 else d
    return s


# ---------- loaders ----------

def load_xlsx(path, progress=None):
    import openpyxl

    import pillowless
    pillowless.install()   # keeps pictures even without Pillow
    full = os.path.getsize(path) <= FULL_LOAD_BYTES
    # cached values first; fall back to formula text when a formula has no cached value
    wb_v = openpyxl.load_workbook(path, read_only=not full, data_only=True, rich_text=full)
    wb_f = openpyxl.load_workbook(path, read_only=True, data_only=False)
    theme = theme_colors(getattr(wb_v, "loaded_theme", None))
    out = {}
    n_sheets = max(len(wb_v.worksheets), 1)
    for i, (ws_v, ws_f) in enumerate(zip(wb_v.worksheets, wb_f.worksheets)):
        rows, styles, cache, style_cache, links, comments, rich = [], {}, {}, {}, {}, {}, {}
        pairs = zip_longest(ws_v.iter_rows(min_row=1, min_col=1),
                            ws_f.iter_rows(min_row=1, min_col=1, values_only=True), fillvalue=())
        expect = min(ws_f.max_row or MAX_ROWS, MAX_ROWS) or 1
        for r, (cells, forms) in enumerate(pairs):
            if r >= MAX_ROWS:
                break
            if progress is not None and r % 500 == 0:
                progress((i + min(r / expect, 1.0)) / n_sheets)
            row = []
            for c, (cell, f) in enumerate(zip_longest(cells, forms)):
                v = getattr(cell, "value", None)
                f = getattr(f, "text", f)   # array formulas
                if isinstance(f, str) and f.startswith("="):
                    row.append(f)            # keep the formula; Excel's last result is a fallback
                    if v is not None:
                        cache[(r, c)] = fmt(v)
                else:
                    row.append(fmt(f if v is None else v))
                    if type(v).__name__ == "CellRichText":   # several fonts in one cell
                        import richtext
                        rich[(r, c)] = richtext.runs_from(v, theme)
                st = cell_style(cell, theme, style_cache) if cell is not None else None
                if st:
                    styles[(r, c)] = st
                link = getattr(cell, "hyperlink", None)   # read-only cells have none
                if link is not None and (link.target or link.location):
                    links[(r, c)] = link.target or f"#{link.location}"
                note = getattr(cell, "comment", None)
                if note is not None:
                    comments[(r, c)] = [note.text or "", note.author or ""]
            rows.append(row)
        widths, merges, heights = {}, [], {}
        if full:
            heights = {r - 1: d.ht for r, d in ws_v.row_dimensions.items()
                       if d.ht and d.customHeight and r <= MAX_ROWS}
            for dim in ws_v.column_dimensions.values():
                if dim.width and dim.min:
                    for c in range(dim.min - 1, min(dim.max or dim.min, dim.min + 500)):
                        widths[c] = dim.width
            merges = [(m.min_row - 1, m.min_col - 1, m.max_row - 1, m.max_col - 1)
                      for m in ws_v.merged_cells.ranges]
        sheet = out[ws_v.title] = Sheet(rows, styles, widths, merges, heights)
        sheet.cache = cache
        sheet.links = links
        sheet.comments = comments
        sheet.rich = rich
        sheet.hidden =getattr(ws_f, "sheet_state", "visible") != "visible"
        if full:
            m = re.fullmatch(r"\$?([A-Za-z]+)\$?(\d+)", ws_v.freeze_panes or "")
            if m:   # "B3" = 2 rows and 1 column frozen
                from formula import col_index
                sheet.freeze = (int(m.group(2)) - 1, col_index(m.group(1)))
            sheet.hidden_rows = {r - 1 for r, d in ws_v.row_dimensions.items() if d.hidden and r <= MAX_ROWS}
            sheet.filter = read_filter(ws_v)
            import rules
            sheet.validations = rules.read_validations(ws_v)
            sheet.cond = rules.read_conditional(ws_v, theme)
            import sheet_props
            sheet.props = sheet_props.read(ws_v, theme)
            sheet.outline_rows = {r - 1: d.outlineLevel for r, d in ws_v.row_dimensions.items()
                                  if d.outlineLevel and r <= MAX_ROWS}
            for d in ws_v.column_dimensions.values():
                if d.outlineLevel and d.min:
                    for c in range(d.min - 1, min(d.max or d.min, d.min + 500)):
                        sheet.outline_cols[c] = d.outlineLevel
            import extras
            sheet.extra = extras.read(ws_v, path)
            import charts
            sheet.charts = charts.from_file(sheet.extra.get("charts"))
            for d in ws_v.column_dimensions.values():
                if d.hidden and d.min:
                    sheet.hidden_cols.update(range(d.min - 1, min(d.max or d.min, d.min + 500)))
    wb_v.close()
    wb_f.close()
    return out


def load_xlsx_rows(path, sheet_name, start, stop):
    """Rows start..stop-1 of one sheet (for loading more of a long sheet while scrolling).
    -> (rows, styles, cache) with positions counted from the top of the sheet; rows is shorter at the end."""
    import openpyxl

    import pillowless
    pillowless.install()   # keeps pictures even without Pillow
    wb_v = openpyxl.load_workbook(path, read_only=True, data_only=True)
    wb_f = openpyxl.load_workbook(path, read_only=True, data_only=False)
    try:
        theme = theme_colors(getattr(wb_v, "loaded_theme", None))
        ws_v, ws_f = wb_v[sheet_name], wb_f[sheet_name]
        rows, styles, cache, style_cache = [], {}, {}, {}
        pairs = zip_longest(ws_v.iter_rows(min_row=start + 1, max_row=stop),
                            ws_f.iter_rows(min_row=start + 1, max_row=stop, values_only=True), fillvalue=())
        for k, (cells, forms) in enumerate(pairs):
            r = start + k
            row = []
            for c, (cell, f) in enumerate(zip_longest(cells, forms)):
                v = getattr(cell, "value", None)
                f = getattr(f, "text", f)
                if isinstance(f, str) and f.startswith("="):
                    row.append(f)
                    if v is not None:
                        cache[(r, c)] = fmt(v)
                else:
                    row.append(fmt(f if v is None else v))
                st = cell_style(cell, theme, style_cache) if cell is not None else None
                if st:
                    styles[(r, c)] = st
            rows.append(row)
        return rows, styles, cache
    finally:
        wb_v.close()
        wb_f.close()


def load_xls(path):
    import xlrd
    try:
        wb = xlrd.open_workbook(path, formatting_info=True)
    except Exception:          # a few odd files only open without formatting
        wb = xlrd.open_workbook(path)
    out = {}

    def text(cell):
        """xls stores dates as day numbers and booleans / errors as small ints: use the cell type."""
        if cell.ctype == xlrd.XL_CELL_DATE:
            try:
                return fmt(xlrd.xldate.xldate_as_datetime(cell.value, wb.datemode))
            except (ValueError, OverflowError, xlrd.xldate.XLDateError):
                return fmt(cell.value)
        if cell.ctype == xlrd.XL_CELL_BOOLEAN:
            return "TRUE" if cell.value else "FALSE"
        if cell.ctype == xlrd.XL_CELL_ERROR:
            return xlrd.error_text_from_code.get(cell.value, "#VALUE!")
        return fmt(cell.value)

    styles_of = xls_styles(wb) if wb.formatting_info else None
    for sh in wb.sheets():
        n = min(sh.nrows, MAX_ROWS)
        sheet = out[sh.name] = Sheet([[text(c) for c in sh.row(r)] for r in range(n)])
        if styles_of is not None:
            xls_layout(sh, sheet, n, styles_of)
    try:
        import xls_formulas
        xls_formulas.apply(out, path)     # formulas back where they reproduce Excel's saved results
    except Exception:
        pass                              # values are already there; formulas are a bonus
    return out


XLS_BORDERS = {1: "thin", 2: "medium", 3: "dashed", 4: "dotted", 5: "thick", 6: "double", 7: "hair",
               8: "mediumDashed", 9: "dashDot", 10: "mediumDashDot", 11: "dashDotDot", 12: "mediumDashDotDot",
               13: "slantDashDot"}


def xls_styles(wb):
    """xf index -> style dict (same keys as formatting.cell_style), cached."""
    cache = {}
    base = wb.font_list[0] if wb.font_list else None    # the workbook's normal font: shown as ours

    def color(idx):
        rgb = wb.colour_map.get(idx)
        return "#%02x%02x%02x" % rgb if rgb else None

    def style(xf_index):
        if xf_index in cache:
            return cache[xf_index]
        xf = wb.xf_list[xf_index]
        st = {}
        font = wb.font_list[xf.font_index] if xf.font_index < len(wb.font_list) else None
        if font is not None:
            if font.bold:
                st["bold"] = True
            if font.italic:
                st["italic"] = True
            if font.underline_type:
                st["underline"] = True
            if font.struck_out:
                st["strike"] = True
            fg = color(font.colour_index)
            if fg and fg != "#000000":
                st["fg"] = fg
            size = font.height / 20
            if size and size != DEFAULT_SIZE and not (base and font.height == base.height):
                st["size"] = size
            if font.name and font.name != DEFAULT_FONT and not (base and font.name == base.name):
                st["font"] = font.name
        bg = xf.background
        if bg.fill_pattern == 1:
            c = color(bg.pattern_colour_index)
            if c:
                st["bg"] = c
        al = xf.alignment
        h = {1: "left", 2: "center", 3: "right", 5: "justify", 6: "center"}.get(al.hor_align)
        if h:
            st["align"] = h
        if al.vert_align in (0, 1):
            st["valign"] = "top" if al.vert_align == 0 else "center"
        if al.text_wrapped:
            st["wrap"] = True
        f = wb.format_map.get(xf.format_key)
        if f is not None and f.format_str and f.format_str != "General":
            st["numfmt"] = f.format_str
        b = xf.border
        border = {}
        for side in ("top", "bottom", "left", "right"):
            kind = XLS_BORDERS.get(getattr(b, f"{side}_line_style"))
            if kind:
                border[side] = [kind, color(getattr(b, f"{side}_colour_index")) or "#000000"]
        if border:
            st["border"] = border
        cache[xf_index] = st or None
        return cache[xf_index]
    return style


def xls_layout(sh, sheet, n, style):
    """Formatting, widths, heights, merges, hidden rows/columns and frozen panes of an .xls sheet."""
    for r in range(n):
        for c in range(sh.row_len(r)):
            st = style(sh.cell_xf_index(r, c))
            if st:
                sheet.styles[(r, c)] = st
    for c, info in sh.colinfo_map.items():
        if info.hidden:
            sheet.hidden_cols.add(c)
        elif info.width:
            sheet.widths[c] = round(info.width / 256, 2)
    for r, info in sh.rowinfo_map.items():
        if r >= n:
            continue
        if info.hidden:
            sheet.hidden_rows.add(r)
        elif info.has_default_height == 0 and info.height:
            sheet.heights[r] = info.height / 20
    sheet.merges = [(rlo, clo, rhi - 1, chi - 1) for rlo, rhi, clo, chi in sh.merged_cells if rlo < n]
    sheet.hidden = sh.visibility != 0
    if sh.panes_are_frozen:
        sheet.freeze = (sh.horz_split_pos or 0, sh.vert_split_pos or 0)
    sheet.reset_structure()


def load_ods(path):
    """stdlib only: an .ods is a zip holding content.xml (+ styles.xml for named styles)."""
    import xml.etree.ElementTree as ET
    import zipfile

    import ods
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("content.xml"))
        try:
            styles_root = ET.fromstring(z.read("styles.xml"))
        except KeyError:
            styles_root = None
    T = ods.T
    look = ods.Styles(root, styles_root)
    cell_tags = (T + "table-cell", T + "covered-table-cell")
    out = {}
    for table in root.iter(T + "table"):
        rows, empty_rows, styles, merges, heights, hidden_rows = [], 0, {}, [], {}, set()
        widths, hidden_cols = look.columns(table)
        for row in table.iter(T + "table-row"):
            cells, empty_cells, row_styles, spans = [], 0, {}, []
            rep_r = int(row.get(T + "number-rows-repeated", 1))
            for c in row:
                if c.tag not in cell_tags:
                    continue
                rep = int(c.get(T + "number-columns-repeated", 1))
                val = ods.cell_text(c)
                at = len(cells) + empty_cells
                st = look.cell(c.get(T + "style-name"))
                if st and (val or rep <= 100):            # styled blanks only when not repeated far
                    for k in range(rep):
                        row_styles[at + k] = st
                sr, sc = int(c.get(T + "number-rows-spanned", 1)), int(c.get(T + "number-columns-spanned", 1))
                if (sr, sc) != (1, 1):
                    spans.append((at, sr, sc))
                if not val:  # defer empties so huge repeated blanks never expand
                    empty_cells += rep
                    continue
                cells += [""] * empty_cells + [val] * rep
                empty_cells = 0
            if not cells and (not row_styles or rep_r > 100):
                empty_rows += rep_r
                continue
            props = look.row(row.get(T + "style-name"))
            hidden = row.get(T + "visibility") in ("collapse", "filter")
            rows += [[] for _ in range(empty_rows)]
            empty_rows = 0
            for _ in range(min(rep_r, MAX_ROWS - len(rows))):
                r = len(rows)
                rows.append(list(cells))
                for c, st in row_styles.items():
                    styles[(r, c)] = st
                if props.get("height"):
                    heights[r] = props["height"]
                if hidden:
                    hidden_rows.add(r)
                merges += [(r, c, r + sr - 1, c + sc - 1) for c, sr, sc in spans]
            if len(rows) >= MAX_ROWS:
                break
        sheet = Sheet(rows, styles, widths, merges, heights)
        sheet.hidden_rows, sheet.hidden_cols = hidden_rows, hidden_cols
        sheet.hidden = look.table_hidden(table)
        sheet.reset_structure()
        out[table.get(T + "name", f"Sheet{len(out) + 1}")] = sheet
    return out


def sniff_csv(path):
    """-> (encoding, dialect) of a csv/tsv file."""
    with open(path, "rb") as f:
        enc = "utf-8-sig" if f.read(3) == b"\xef\xbb\xbf" else "utf-8"
    with open(path, newline="", encoding=enc, errors="replace") as f:
        sample = f.read(4096)
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
    except csv.Error:
        dialect = csv.excel_tab if path.lower().endswith(".tsv") else csv.excel
    return enc, dialect


def load_csv(path):
    enc, dialect = sniff_csv(path)
    rows = []
    with open(path, newline="", encoding=enc, errors="replace") as f:
        for row in csv.reader(f, dialect):
            rows.append(row)
            if len(rows) >= MAX_ROWS:
                break
    return {os.path.basename(path): Sheet(rows)}


def load_xlsb(path):
    import xlsb
    return xlsb.load(path)


LOADERS = {
    ".xlsx": load_xlsx, ".xlsm": load_xlsx, ".xltx": load_xlsx, ".xltm": load_xlsx,
    ".xls": load_xls,
    ".xlsb": load_xlsb,
    ".ods": load_ods,
    ".csv": load_csv, ".tsv": load_csv, ".txt": load_csv,
}
CSV_EXTS = (".csv", ".tsv", ".txt")
XLSX_EXTS = (".xlsx", ".xlsm", ".xltx", ".xltm")
SAVABLE = XLSX_EXTS + CSV_EXTS + (".ods", ".xlsb")  # others must "Save As"


def load_names(path):
    """Workbook-level defined names that point at cells: {name: "Sheet1!$A$1:$B$5"} (.xlsx/.xlsm)."""
    if os.path.splitext(path)[1].lower() not in XLSX_EXTS:
        return {}
    import openpyxl

    import pillowless
    pillowless.install()   # keeps pictures even without Pillow

    from formula import lex
    wb = openpyxl.load_workbook(path, read_only=True)
    out = {}
    try:
        for name, dn in wb.defined_names.items():
            text = (dn.attr_text or "").strip()
            toks = [k for k, _ in lex(text) if k != "ws"]
            if not name.startswith("_xlnm") and toks and all(k in ("ref", "cols", "rows") for k in toks) \
                    and len(toks) == 1:
                out[name] = text
    finally:
        wb.close()
    return out


def apply_names(wb, names):
    """Workbook-level names = `names` ({name: reference text}); built-in names (print areas…) untouched."""
    from openpyxl.workbook.defined_name import DefinedName
    for name in list(wb.defined_names):
        if not name.startswith("_xlnm") and name not in names:
            del wb.defined_names[name]
    for name, text in names.items():
        wb.defined_names[name] = DefinedName(name, attr_text=text)


def load(path, progress=None):
    """path -> {sheet name: Sheet}. progress(fraction) is called now and then while an .xlsx loads
    (it may raise to cancel)."""
    ext = os.path.splitext(path)[1].lower()
    if ext not in LOADERS:
        raise ValueError(f"Can't open '{ext}' files.")
    if progress is not None and LOADERS[ext] is load_xlsx:
        return load_xlsx(path, progress)
    return LOADERS[ext](path)


# ---------- savers ----------

def trim(rows):
    rows = [list(r) for r in rows]
    for r in rows:
        while r and r[-1] == "":
            r.pop()
    while rows and not rows[-1]:
        rows.pop()
    return rows


def apply_style(cell, st):
    """Style dict -> openpyxl cell (font, fill, alignment, number format, borders).
    Keeps the font name and diagonal borders."""
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    st = st or {}
    old, al = cell.font, cell.alignment
    fg = st.get("fg")
    name = st.get("font") or old.name
    cell.font = Font(name=name, family=old.family if name == old.name else None,
                     scheme=old.scheme if name == old.name else None,
                     sz=st.get("size", DEFAULT_SIZE), b=st.get("bold", False), i=st.get("italic", False),
                     u="single" if st.get("underline") else None, strike=st.get("strike", False),
                     color="FF" + fg[1:].upper() if fg else None)
    if not st.get("gradient"):              # a gradient from the file stays as it is
        cell.fill = PatternFill("solid", fgColor="FF" + st["bg"][1:].upper()) if st.get("bg") else PatternFill()
    rot = st.get("rotate") or 0
    cell.alignment = Alignment(horizontal=st.get("align"), vertical=st.get("valign"), wrap_text=bool(st.get("wrap")),
                               indent=st.get("indent", 0), text_rotation=90 - rot if rot < 0 else rot,
                               shrink_to_fit=al.shrink_to_fit)
    cell.number_format = st.get("numfmt", "General")
    old_b = cell.border
    sides = {}
    for name in SIDES:
        line = (st.get("border") or {}).get(name)
        sides[name] = Side(style=line[0], color="FF" + line[1][1:].upper()) if line else Side()
    cell.border = Border(**sides, diagonal=old_b.diagonal, diagonalUp=old_b.diagonalUp,
                         diagonalDown=old_b.diagonalDown)   # diagonals aren't edited here: keep them
    from openpyxl.styles import Protection
    if st.get("unlocked") or not cell.protection.locked:   # editable on a protected sheet
        cell.protection = Protection(locked=not st.get("unlocked"), hidden=cell.protection.hidden)


def rebuild_sheet(ws, sheet):
    """Rows/columns were inserted or deleted: lay the sheet out again from the model, moving each
    original cell's full style (borders included), row heights and column widths to its new place."""
    from copy import copy

    from openpyxl.comments import Comment
    from openpyxl.utils import get_column_letter
    old, rich = {}, {}
    for row in ws.iter_rows():
        for cell in row:
            if type(cell).__name__ != "MergedCell" and (cell.has_style or cell.comment or cell.hyperlink):
                old[(cell.row - 1, cell.column - 1)] = (copy(cell._style), cell.comment, cell.hyperlink)
            if type(cell.value).__name__ == "CellRichText":   # several fonts inside the cell
                rich[(cell.row - 1, cell.column - 1)] = cell.value
    rdims = {r - 1: (d.ht, d.hidden) for r, d in ws.row_dimensions.items() if d.ht or d.hidden}
    cdims = {}
    for d in ws.column_dimensions.values():
        if d.min and (d.width or d.hidden):
            for c in range(d.min - 1, min(d.max or d.min, d.min + 500)):
                cdims[c] = (d.width, d.hidden)
    for m in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(m))
    ws._cells.clear()
    ws.row_dimensions.clear()
    ws.column_dimensions.clear()
    new_r = Sheet._inverse(sheet.rowmap, sheet.rowtail)
    new_c = Sheet._inverse(sheet.colmap, sheet.coltail)

    def moved(R, C):
        r, c = new_r(R), new_c(C)
        return None if r is None or c is None else (r, c)

    for (R, C), (style, comment, link) in old.items():
        pos = moved(R, C)
        if pos:
            cell = ws.cell(row=pos[0] + 1, column=pos[1] + 1)
            cell._style = style
            if comment:
                cell.comment = Comment(comment.text, comment.author)
            text = sheet.rows[pos[0]][pos[1]] if pos[0] < len(sheet.rows) and pos[1] < len(sheet.rows[pos[0]]) else ""
            if link is not None and link.target and text != "":   # emptied cells drop their link
                cell.hyperlink = link.target
    for r, row in enumerate(sheet.rows):
        for c, t in enumerate(row):
            if t != "":
                value = rich.get(sheet.orig_pos(r, c))
                # unchanged text keeps its rich formatting; edited text is written plain
                ws.cell(row=r + 1, column=c + 1).value = value if value is not None and str(value) == t \
                    else to_value(t)
    # formatting edited in the app (compared with the original cell it came from)
    keys = set(sheet.styles) | {p for p in (moved(*k) for k in sheet.styles0) if p}
    for r, c in keys:
        R, C = sheet.orig_pos(r, c)
        before = sheet.styles0.get((R, C)) if R is not None and C is not None else None
        if sheet.styles.get((r, c)) != before:
            apply_style(ws.cell(row=r + 1, column=c + 1), sheet.styles.get((r, c)))
    for r1, c1, r2, c2 in sheet.merges:
        ws.merge_cells(start_row=r1 + 1, start_column=c1 + 1, end_row=r2 + 1, end_column=c2 + 1)
    for R, (ht, hidden) in rdims.items():
        r = new_r(R)
        if r is not None:
            ws.row_dimensions[r + 1].ht, ws.row_dimensions[r + 1].hidden = ht, hidden
    for C, (width, hidden) in cdims.items():
        c = new_c(C)
        if c is not None:
            d = ws.column_dimensions[get_column_letter(c + 1)]
            d.width, d.hidden = width, hidden
    apply_hidden(ws, sheet)
    move_ranges(ws, sheet, new_r, new_c)


def _span(lo, hi, new):
    """Original rows/cols lo..hi -> (first, last) of what's left after the move, or None if all deleted."""
    i = lo
    while i <= hi and new(i) is None:
        i += 1
    if i > hi:
        return None
    j = hi
    while new(j) is None:
        j -= 1
    return new(i), new(j)


def move_ranges(ws, sheet, new_r, new_c):
    """After rows/columns moved, update everything in the file that names cell ranges on this sheet:
    data validation, conditional formatting, freeze panes, autofilter, Excel tables, print titles /
    area and defined names. A range whose cells were all deleted is dropped."""
    from formula import col_index, col_letters, remap, remap_ref

    span_r = lambda lo, hi: _span(lo, hi, new_r)
    span_c = lambda lo, hi: _span(lo, hi, new_c)
    title = ws.title

    def cells(text):
        """A space-separated list of ranges (sqref)."""
        out = []
        for part in str(text).split():
            kind = "cols" if re.fullmatch(r"\$?[A-Za-z]+:\$?[A-Za-z]+", part) else \
                "rows" if re.fullmatch(r"\$?\d+:\$?\d+", part) else "ref"
            new = remap_ref(kind, part, span_r, span_c)
            if new != "#REF!":
                out.append(new)
        return " ".join(out)

    def formula(text):
        return remap(text, title, title, span_r, span_c) if text else text

    for dv in list(ws.data_validations.dataValidation):
        new = cells(dv.sqref)
        if not new:
            ws.data_validations.dataValidation.remove(dv)
            continue
        dv.sqref = new
        dv.formula1, dv.formula2 = formula(dv.formula1), formula(dv.formula2)

    from openpyxl.formatting.formatting import ConditionalFormattingList
    moved_cf = ConditionalFormattingList()
    for cf in ws.conditional_formatting:
        new = cells(cf.sqref)
        if not new:
            continue
        for rule in cf.rules:
            rule.formula = [formula(f) for f in (rule.formula or [])]
            moved_cf.add(new, rule)
    ws.conditional_formatting = moved_cf

    if ws.freeze_panes:
        # the first unfrozen cell moves with rows / columns inserted or deleted above / left of it
        m = re.fullmatch(r"\$?([A-Za-z]+)\$?(\d+)", ws.freeze_panes)
        if m:
            r, c = int(m.group(2)) - 1, col_index(m.group(1))
            nr, nc = _span(r, r + MAX_ROWS, new_r), _span(c, c + 16384, new_c)
            r, c = (nr[0] if nr else r), (nc[0] if nc else c)
            ws.freeze_panes = None if (r, c) == (0, 0) else f"{col_letters(c)}{r + 1}"

    if ws.auto_filter.ref:
        ws.auto_filter.ref = cells(ws.auto_filter.ref) or None

    for name in list(ws.tables):   # (ws.tables.items() gives range text, not the Table)
        table = ws.tables[name]
        new = cells(table.ref)
        if not new:
            del ws.tables[name]
            continue
        table.ref = new
        if table.autoFilter is not None:
            table.autoFilter.ref = new
        _table_columns(ws, table)

    if ws.print_title_rows:
        ws.print_title_rows = cells(ws.print_title_rows).replace("$", "") or None
    if ws.print_title_cols:
        ws.print_title_cols = cells(ws.print_title_cols).replace("$", "") or None
    if ws.print_area:
        areas = ws.print_area if isinstance(ws.print_area, (list, tuple)) else str(ws.print_area).split(",")
        moved_areas = [a for a in (cells(a.split("!")[-1]) for a in areas) if a]
        ws.print_area = [a.replace("$", "") for a in moved_areas] or None

    names = list(ws.parent.defined_names.values()) + list(ws.defined_names.values())
    for dn in names:
        if dn.attr_text and title.lower() in dn.attr_text.lower():
            dn.attr_text = remap(dn.attr_text, None, title, span_r, span_c)


def _table_columns(ws, table):
    """An Excel table needs one column entry per column, named like its header cell."""
    from openpyxl.utils import range_boundaries
    from openpyxl.worksheet.table import TableColumn
    c1, r1, c2, _ = range_boundaries(table.ref)
    names = {}
    for i, c in enumerate(range(c1, c2 + 1), start=1):
        name = str(ws.cell(row=r1, column=c).value or f"Column{i}")
        while name in names:          # header names must be unique
            name += "_"
        names[name] = TableColumn(id=i, name=name)
    table.tableColumns = list(names.values())


def sync_sheets(wb, book):
    """Make the workbook's sheets match the app: delete removed, rename, add new, same order.
    Each model Sheet knows its name in the file (disk_name, None if new)."""
    keep = {s.disk_name for s in book.values() if s.disk_name}
    for ws in list(wb.worksheets):
        if ws.title not in keep:
            wb.remove(ws)
    by_name = {}
    for i, (name, s) in enumerate(book.items()):   # temp names first: renames may swap names
        ws = wb[s.disk_name] if s.disk_name else wb.create_sheet()
        ws.title = f"__simple_sheets_{i}"
        by_name[name] = ws
    for name, ws in by_name.items():
        ws.title = name
        set_state(ws, book[name])
    wb._sheets = [by_name[n] for n in book]
    wb.active = next(i for i, s in enumerate(book.values()) if not s.hidden)


def set_state(ws, sheet):
    """Hidden sheet -> 'hidden' (a 'veryHidden' sheet stays very hidden)."""
    if not sheet.hidden:
        ws.sheet_state = "visible"
    elif ws.sheet_state == "visible":
        ws.sheet_state = "hidden"


def column_dim(ws, c):
    """The dimension of column c (0-based) on its own: a shared range like E:G is split first,
    so changing one column never changes (or overlaps) its neighbours."""
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.dimensions import ColumnDimension
    dims = ws.column_dimensions
    for key, d in list(dims.items()):
        lo = d.min or 0
        hi = d.max or lo
        if lo and lo <= c + 1 <= hi and hi > lo:
            del dims[key]
            for cc in range(lo, min(hi, lo + 2000) + 1):
                letter = get_column_letter(cc)
                dims[letter] = ColumnDimension(ws, index=letter, width=d.width, hidden=d.hidden,
                                               outlineLevel=d.outlineLevel)
            break
    return dims[get_column_letter(c + 1)]


def apply_widths(wb, widths):
    """widths: {(sheet, 0, col): width in characters, or None for the default width}."""
    from openpyxl.utils import get_column_letter
    for (sheet, _, c), w in widths.items():
        ws = wb[sheet]
        d = column_dim(ws, c)
        if w is not None:
            d.width = w
        elif d.hidden or d.outlineLevel:
            d.width = ws.sheet_format.defaultColWidth or 8.43   # keep the other settings
        else:
            del ws.column_dimensions[get_column_letter(c + 1)]   # back to the sheet's default


FILTER_BLANK = "(Blanks)"


def read_filter(ws):
    """The sheet's autofilter -> sheet.filter dict (or None)."""
    from openpyxl.utils import range_boundaries
    af = ws.auto_filter
    if not af or not af.ref:
        return None
    try:
        c1, r1, c2, r2 = range_boundaries(af.ref)
    except (ValueError, TypeError):
        return None
    criteria = {}
    back = {"equal": "=", "notEqual": "<>", "greaterThan": ">", "greaterThanOrEqual": ">=", "lessThan": "<",
            "lessThanOrEqual": "<="}
    for fc in af.filterColumn or []:
        if fc.customFilters is not None and fc.customFilters.customFilter:
            cf = fc.customFilters.customFilter[0]
            op, val = back.get(cf.operator or "equal", "="), str(cf.val or "")
            if val.startswith("*") and val.endswith("*") and len(val) > 1:
                op, val = ("!contains" if op == "<>" else "contains"), val[1:-1]
            elif val.endswith("*"):
                op, val = "begins", val[:-1]
            elif val.startswith("*"):
                op, val = "ends", val[1:]
            criteria[c1 - 1 + fc.colId] = {"cond": [op, val]}
            continue
        if fc.top10 is not None:
            criteria[c1 - 1 + fc.colId] = {"cond": ["top" if fc.top10.top is not False else "bottom",
                                                     str(int(fc.top10.val or 10))]}
            continue
        if fc.filters is not None:
            keep = [str(v) for v in (fc.filters.filter or [])]
            if fc.filters.blank:
                keep.append(FILTER_BLANK)
            criteria[c1 - 1 + fc.colId] = keep
    return {"range": (r1 - 1, c1 - 1, r2 - 1, c2 - 1), "criteria": criteria}


def write_filter(ws, f):
    """sheet.filter dict -> the sheet's autofilter (None removes it)."""
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.filters import AutoFilter
    tables = list(ws.tables.values())
    for t in tables:                       # an Excel table's own filter: criteria reset, buttons kept
        if t.autoFilter is not None:
            t.autoFilter.filterColumn = []
    ws.auto_filter = AutoFilter()
    if not f:
        return
    r1, c1, r2, c2 = f["range"]
    target = ws.auto_filter
    if f.get("table"):                     # the filter belongs to a table: write it into the table
        t = next((t for t in tables if (t.displayName or t.name) == f["table"]), None)
        if t is not None:
            if t.autoFilter is None:
                t.autoFilter = AutoFilter(ref=t.ref)
            target = t.autoFilter
    target.ref = f"{get_column_letter(c1 + 1)}{r1 + 1}:{get_column_letter(c2 + 1)}{r2 + 1}"
    from openpyxl.worksheet.filters import CustomFilter, CustomFilters, FilterColumn, Top10
    ops = {"=": "equal", "<>": "notEqual", ">": "greaterThan", ">=": "greaterThanOrEqual", "<": "lessThan",
           "<=": "lessThanOrEqual"}
    for c, keep in sorted(f["criteria"].items()):
        if isinstance(keep, list):
            target.add_filter_column(c - c1, [v for v in keep if v != FILTER_BLANK], blank=FILTER_BLANK in keep)
        elif keep.get("cond"):                 # a condition: written as Excel custom / top-10 filters
            op, want = keep["cond"]
            if op in ("top", "bottom"):
                fc = FilterColumn(colId=c - c1, top10=Top10(top=op == "top", val=float(want or 10)))
            else:
                text = {"contains": f"*{want}*", "!contains": f"*{want}*", "begins": f"{want}*",
                        "ends": f"*{want}"}.get(op, want)
                xop = "notEqual" if op == "!contains" else ops.get(op, "equal")
                fc = FilterColumn(colId=c - c1, customFilters=CustomFilters(customFilter=[CustomFilter(xop, text)]))
            target.filterColumn.append(fc)
        # a colour filter has no plain-file form here: its hidden rows are saved as hidden rows


def apply_filters(wb, filters):
    """filters: {(sheet, 0, 0): filter dict or None}."""
    for (sheet, _, _), f in filters.items():
        write_filter(wb[sheet], f)


def apply_comments(wb, comments):
    """comments: {(sheet, r, c): [text, author] or None}."""
    from openpyxl.comments import Comment
    for (sheet, r, c), note in comments.items():
        cell = wb[sheet].cell(row=r + 1, column=c + 1)
        if type(cell).__name__ == "MergedCell":
            continue
        cell.comment = Comment(note[0], note[1] or "") if note else None


def apply_freeze(wb, freeze):
    """freeze: {(sheet, 0, 0): (rows, cols)} -> the sheet's freeze panes."""
    from openpyxl.utils import get_column_letter
    for (sheet, _, _), (rows, cols) in freeze.items():
        wb[sheet].freeze_panes = None if (rows, cols) == (0, 0) else f"{get_column_letter(cols + 1)}{rows + 1}"


def apply_links(wb, links):
    """links: {(sheet, r, c): address or None}."""
    for (sheet, r, c), url in links.items():
        cell = wb[sheet].cell(row=r + 1, column=c + 1)
        if type(cell).__name__ == "MergedCell":
            continue
        if url and url.startswith("#"):
            cell.hyperlink = None
            from openpyxl.worksheet.hyperlink import Hyperlink
            cell.hyperlink = Hyperlink(ref=cell.coordinate, location=url[1:])
        else:
            cell.hyperlink = url or None


def apply_heights(wb, heights):
    """heights: {(sheet, r, 0): points or None (automatic)}."""
    for (sheet, r, _), pt in heights.items():
        wb[sheet].row_dimensions[r + 1].height = pt


def patch_xlsx(src, dst, values, styles, book=None, widths=None, freeze=None, filters=None,
               comments=None, heights=None, links=None, names=None, settings=None, outline=None,
               book_props=None, extra=None):
    """Write only changed cells into a copy of src: keeps everything else (formulas, borders...).
    values: {(sheet, row, col): text}; styles: {(sheet, row, col): style dict or None}.
    Sheets in `book` whose rows/columns moved are rebuilt instead."""
    import openpyxl

    import pillowless
    pillowless.install()   # keeps pictures even without Pillow
    keep_vba = src.lower().endswith((".xlsm", ".xltm")) and dst.lower().endswith((".xlsm", ".xltm"))
    # rich_text: cells with several fonts are kept as they are (default loading flattens them)
    wb = openpyxl.load_workbook(src, keep_vba=keep_vba, rich_text=True)
    wb.template = dst.lower().endswith((".xltx", ".xltm"))   # a template, or a normal workbook made from one
    with open(src, "rb") as f:   # the original as it is now (saving in place overwrites it)
        original = f.read()
    keep = {"src": original, "names": {s.disk_name: n for n, s in (book or {}).items() if s.disk_name}, "maps": {}}
    for n, s in (book or {}).items():
        if s.disk_name and s.structure_changed():   # shapes move with inserted / deleted rows and columns
            keep["maps"][n] = (s._inverse(s.rowmap, s.rowtail), s._inverse(s.colmap, s.coltail))
    if book is not None:
        sync_sheets(wb, book)
    new = {n for n, s in (book or {}).items() if not s.disk_name}
    for name in new:   # not in the file yet: write it whole
        fill_sheet(wb[name], book[name])
    rebuilt = {n for n, s in (book or {}).items() if s.disk_name and s.structure_changed()}
    rebuilt |= new
    for name in rebuilt:
        rebuild_sheet(wb[name], book[name])
    for name, s in (book or {}).items():
        if name not in rebuilt and s.merges != s.disk_merges:
            apply_merges(wb[name], s)
    for changes, is_style in ((values, False), (styles, True)):
        for (sheet, r, c), v in changes.items():
            if sheet in rebuilt:
                continue
            cell = wb[sheet].cell(row=r + 1, column=c + 1)
            if type(cell).__name__ == "MergedCell":  # merged-away cells are read-only
                continue
            if is_style:
                apply_style(cell, v)
            else:
                cell.value = to_value(v)
                if v == "":
                    # an emptied cell loses its link too; otherwise openpyxl puts the link
                    # address back into the empty cell when the file is opened again
                    cell.hyperlink = None
    for name, s in (book or {}).items():
        if name not in rebuilt:
            apply_hidden(wb[name], s)
    apply_widths(wb, widths or {})   # positions are current ones, also right for rebuilt sheets
    apply_freeze(wb, freeze or {})
    apply_filters(wb, filters or {})
    apply_comments(wb, comments or {})
    apply_heights(wb, heights or {})
    apply_links(wb, links or {})
    if names is not None:   # only when the names were changed in the app
        apply_names(wb, names)
    import sheet_props
    sheet_props.apply_all(wb, settings or {})
    apply_outline(wb, outline or {})
    if book_props is not None:
        sheet_props.apply_book(wb, book_props)
    for fn in (extra or []):   # new rules, charts, images… added in the app (see App.save_extras)
        fn(wb)
    wb.save(dst)
    try:   # shapes, text boxes and sparklines openpyxl doesn't keep: copied back from the original
        import keep_parts
        keep_parts.restore(keep["src"], dst, keep["names"], keep["maps"])
    except Exception:
        pass


def apply_outline(wb, outline):
    """{(sheet, r, 0): level} for rows and {('c', sheet, 0, c): level} for columns."""
    from openpyxl.utils import get_column_letter
    for key, level in outline.items():
        if key[0] == "c" and len(key) == 4:
            _, sheet, _, c = key
            wb[sheet].column_dimensions[get_column_letter(c + 1)].outlineLevel = level or 0
        else:
            sheet, r, _ = key
            wb[sheet].row_dimensions[r + 1].outlineLevel = level or 0


EXCEL_ERRORS = {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"}


def _sheet_files(files):
    """{sheet name: path of its XML inside the xlsx zip}."""
    from html import unescape
    book = files["xl/workbook.xml"].decode("utf-8")
    rels = files["xl/_rels/workbook.xml.rels"].decode("utf-8")
    targets = {}
    for rel in re.findall(r"<Relationship\b[^>]*>", rels):
        rid, target = re.search(r'Id="([^"]+)"', rel), re.search(r'Target="([^"]+)"', rel)
        if rid and target:
            t = target.group(1).lstrip("/")
            targets[rid.group(1)] = t if t.startswith("xl/") else "xl/" + t
    out = {}
    for tag in re.findall(r"<sheet\b[^>]*>", book):
        name, rid = re.search(r'\bname="([^"]*)"', tag), re.search(r'\br:id="([^"]+)"', tag)
        if name and rid and rid.group(1) in targets:
            out[unescape(name.group(1))] = targets[rid.group(1)]
    return out


def write_cached_values(path, values):
    """Store formula results in a saved .xlsx/.xlsm, so programs that don't calculate formulas
    (previews, phone viewers, some importers) show the values instead of blanks.
    values: {sheet name: {(r, c): number / text / bool / error text}}.
    openpyxl writes formula cells with an empty <v></v>; this fills it in afterwards."""
    import zipfile
    from xml.sax.saxutils import escape
    with zipfile.ZipFile(path) as z:
        infos = z.infolist()
        files = {i.filename: z.read(i.filename) for i in infos}
    where = _sheet_files(files)
    from formula import col_index
    # a formula cell with an empty value: <c r="B2" s="1"><f>A2*2</f><v></v></c>  (or <v/>)
    cell_re = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*)>(<f\b[^>]*/>|<f\b[^>]*>.*?</f>)(?:<v\s*/>|<v></v>)</c>',
                         re.S)
    changed = False
    for sheet, cells in values.items():
        name = where.get(sheet)
        if not name or not cells:
            continue

        def fill(m):
            col, row, attrs, f = m.group(1, 2, 3, 4)
            v = cells.get((int(row) - 1, col_index(col)))
            attrs = re.sub(r'\s+t="[^"]*"', "", attrs)
            if v is None:
                return m.group(0)
            if isinstance(v, bool):
                kind, text = ' t="b"', "1" if v else "0"
            elif isinstance(v, (int, float)):
                if v != v or v in (float("inf"), float("-inf")):
                    return m.group(0)
                kind, text = "", repr(float(v)) if not float(v).is_integer() else str(int(v))
            elif str(v) in EXCEL_ERRORS:
                kind, text = ' t="e"', str(v)
            else:
                kind, text = ' t="str"', escape(str(v))
            return f'<c r="{col}{row}"{attrs}{kind}>{f}<v>{text}</v></c>'

        xml = files[name].decode("utf-8")
        new = cell_re.sub(fill, xml)
        if new != xml:
            files[name] = new.encode("utf-8")
            changed = True
    if not changed:
        return
    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for info in infos:
            z.writestr(info, files[info.filename])
    os.replace(tmp, path)


def fill_sheet(ws, sheet):
    """Write a whole model sheet into an empty worksheet: values, formats, merges, sizes."""
    from openpyxl.utils import get_column_letter
    for r, row in enumerate(sheet.rows):
        for c, t in enumerate(row):
            if t != "":
                ws.cell(row=r + 1, column=c + 1).value = to_value(t)
    for (r, c), st in sheet.styles.items():
        if st:
            apply_style(ws.cell(row=r + 1, column=c + 1), st)
    if sheet.rich:
        import richtext
        for (r, c), runs in sheet.rich.items():
            if r < len(sheet.rows) and c < len(sheet.rows[r]) and richtext.text_of(runs) == sheet.rows[r][c]:
                ws.cell(row=r + 1, column=c + 1).value = richtext.to_openpyxl(runs)
    for r1, c1, r2, c2 in sheet.merges:
        ws.merge_cells(start_row=r1 + 1, start_column=c1 + 1, end_row=r2 + 1, end_column=c2 + 1)
    for c, w in sheet.widths.items():
        ws.column_dimensions[get_column_letter(c + 1)].width = w
    for r, h in sheet.heights.items():
        ws.row_dimensions[r + 1].height = h
    apply_hidden(ws, sheet)
    if sheet.freeze != (0, 0):
        apply_freeze(ws.parent, {(ws.title, 0, 0): sheet.freeze})
    if sheet.filter:
        write_filter(ws, sheet.filter)
    if sheet.comments:
        apply_comments(ws.parent, {(ws.title, r, c): note for (r, c), note in sheet.comments.items()})
    if sheet.links:
        apply_links(ws.parent, {(ws.title, r, c): url for (r, c), url in sheet.links.items()})
    import sheet_props
    for key, val in sheet.props.items():
        sheet_props.apply(ws, key, val)
    for r, level in sheet.outline_rows.items():
        ws.row_dimensions[r + 1].outlineLevel = level
    for c, level in sheet.outline_cols.items():
        ws.column_dimensions[get_column_letter(c + 1)].outlineLevel = level


def apply_merges(ws, sheet):
    """Merged ranges in the file = the app's (sheet.merges)."""
    want = {(r1 + 1, c1 + 1, r2 + 1, c2 + 1) for r1, c1, r2, c2 in sheet.merges}
    for m in list(ws.merged_cells.ranges):
        if (m.min_row, m.min_col, m.max_row, m.max_col) not in want:
            ws.unmerge_cells(str(m))
    have = {(m.min_row, m.min_col, m.max_row, m.max_col) for m in ws.merged_cells.ranges}
    for r1, c1, r2, c2 in want - have:
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)


def apply_hidden(ws, sheet):
    """Hidden rows/columns in the file = the app's (sheet.hidden_rows / hidden_cols)."""
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.dimensions import ColumnDimension
    hidden = sheet.hidden_rows | sheet.filtered_rows   # Excel stores filtered rows as hidden too
    for r, d in list(ws.row_dimensions.items()):
        d.hidden = (r - 1) in hidden
    for r in hidden:
        ws.row_dimensions[r + 1].hidden = True
    dims = ws.column_dimensions
    for key, d in list(dims.items()):    # a dimension can cover a range of columns (min..max)
        lo = d.min or 0
        hi = d.max or lo
        if not lo:
            continue
        want = [(c - 1) in sheet.hidden_cols for c in range(lo, min(hi, lo + 2000) + 1)]
        if all(w == want[0] for w in want):
            d.hidden = want[0]
            continue
        del dims[key]                    # mixed: split the range into single columns
        for c in range(lo, hi + 1):
            letter = get_column_letter(c)
            dims[letter] = ColumnDimension(ws, index=letter, width=d.width, customWidth=d.customWidth,
                                           hidden=(c - 1) in sheet.hidden_cols, outlineLevel=d.outlineLevel)
    covered = [(d.min, d.max or d.min) for d in dims.values() if d.min]
    for c in sheet.hidden_cols:
        if not any(lo <= c + 1 <= hi for lo, hi in covered):
            dims[get_column_letter(c + 1)].hidden = True


def write_xlsx(dst, sheets, names=None, book_props=None):
    import openpyxl

    import pillowless
    pillowless.install()   # keeps pictures even without Pillow

    import charts
    import rules
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, sheet in sheets.items():
        ws = wb.create_sheet(name[:31])
        fill_sheet(ws, sheet)
        set_state(ws, sheet)
        rules.write_all(ws, sheet)
        charts.write_all(ws, sheet)
        import images
        images.write_all(ws, sheet)
    wb.active = next(i for i, s in enumerate(sheets.values()) if not s.hidden)
    if names:
        apply_names(wb, names)
    if book_props:
        import sheet_props
        sheet_props.apply_book(wb, book_props)
    wb.template = dst.lower().endswith((".xltx", ".xltm"))
    wb.save(dst)


def write_csv(dst, rows, like=None):
    """like: an existing csv whose encoding/delimiter to keep."""
    enc, dialect = sniff_csv(like) if like and os.path.exists(like) else ("utf-8", csv.excel)
    if dst.lower().endswith(".tsv") and not like:
        dialect = csv.excel_tab
    with open(dst, "w", newline="", encoding=enc) as f:
        csv.writer(f, dialect).writerows(trim(rows))
