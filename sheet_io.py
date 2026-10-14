"""Load / save spreadsheet files. A book is {name: Sheet}."""
import csv
import datetime as dt
import os
from itertools import zip_longest

from formatting import DEFAULT_SIZE, NUMBER, cell_style, parse_date, theme_colors

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
        self.merges = merges or []     # [(r1, c1, r2, c2)], 0-based inclusive
        self.px = {}                   # c -> width set by dragging (view only)
        self.view = None               # remembered scroll/selection
        self.disk_name = None          # this sheet's name in the file on disk (None = new sheet)
        self.hidden = False            # whole sheet hidden (no tab)
        self.disk_hidden = False       # ... as in the file on disk
        self.hidden_rows = set()       # row indexes hidden
        self.hidden_cols = set()       # column indexes hidden
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

def load_xlsx(path):
    import openpyxl
    full = os.path.getsize(path) <= FULL_LOAD_BYTES
    # cached values first; fall back to formula text when a formula has no cached value
    wb_v = openpyxl.load_workbook(path, read_only=not full, data_only=True)
    wb_f = openpyxl.load_workbook(path, read_only=True, data_only=False)
    theme = theme_colors(getattr(wb_v, "loaded_theme", None))
    out = {}
    for ws_v, ws_f in zip(wb_v.worksheets, wb_f.worksheets):
        rows, styles, cache, style_cache, links = [], {}, {}, {}, {}
        pairs = zip_longest(ws_v.iter_rows(min_row=1, min_col=1),
                            ws_f.iter_rows(min_row=1, min_col=1, values_only=True), fillvalue=())
        for r, (cells, forms) in enumerate(pairs):
            if r >= MAX_ROWS:
                break
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
                st = cell_style(cell, theme, style_cache) if cell is not None else None
                if st:
                    styles[(r, c)] = st
                link = getattr(cell, "hyperlink", None)   # read-only cells have none
                if link is not None and (link.target or link.location):
                    links[(r, c)] = link.target or f"#{link.location}"
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
        sheet.hidden =getattr(ws_f, "sheet_state", "visible") != "visible"
        if full:
            sheet.hidden_rows = {r - 1 for r, d in ws_v.row_dimensions.items() if d.hidden and r <= MAX_ROWS}
            for d in ws_v.column_dimensions.values():
                if d.hidden and d.min:
                    sheet.hidden_cols.update(range(d.min - 1, min(d.max or d.min, d.min + 500)))
    wb_v.close()
    wb_f.close()
    return out


def load_xls(path):
    import xlrd
    wb = xlrd.open_workbook(path)
    out = {}
    for sh in wb.sheets():
        n = min(sh.nrows, MAX_ROWS)
        out[sh.name] = Sheet([[fmt(v) for v in sh.row_values(r)] for r in range(n)])
    return out


def load_ods(path):
    # stdlib only: an .ods is a zip holding content.xml
    import xml.etree.ElementTree as ET
    import zipfile
    T = "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}"
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("content.xml"))
    cell_tags = (T + "table-cell", T + "covered-table-cell")
    out = {}
    for table in root.iter(T + "table"):
        rows, empty_rows = [], 0
        for row in table.iter(T + "table-row"):
            cells, empty_cells = [], 0
            for c in row:
                if c.tag not in cell_tags:
                    continue
                rep = int(c.get(T + "number-columns-repeated", 1))
                val = "\n".join("".join(p.itertext()) for p in c if p.tag.endswith("}p"))
                if not val:  # defer empties so huge repeated blanks never expand
                    empty_cells += rep
                    continue
                cells += [""] * empty_cells + [val] * rep
                empty_cells = 0
            rep = int(row.get(T + "number-rows-repeated", 1))
            if not cells:
                empty_rows += rep
                continue
            rows += [[] for _ in range(empty_rows)] + [list(cells) for _ in range(rep)]
            empty_rows = 0
            if len(rows) >= MAX_ROWS:
                rows = rows[:MAX_ROWS]
                break
        out[table.get(T + "name", f"Sheet{len(out) + 1}")] = Sheet(rows)
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


LOADERS = {
    ".xlsx": load_xlsx, ".xlsm": load_xlsx,
    ".xls": load_xls,
    ".ods": load_ods,
    ".csv": load_csv, ".tsv": load_csv, ".txt": load_csv,
}
CSV_EXTS = (".csv", ".tsv", ".txt")
XLSX_EXTS = (".xlsx", ".xlsm")
SAVABLE = XLSX_EXTS + CSV_EXTS  # others must "Save As"


def load(path):
    ext = os.path.splitext(path)[1].lower()
    if ext not in LOADERS:
        raise ValueError(f"Can't open '{ext}' files.")
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
    """Style dict -> openpyxl cell. Keeps font name, vertical alignment, wrap, borders."""
    from openpyxl.styles import Alignment, Font, PatternFill
    st = st or {}
    old, al = cell.font, cell.alignment
    fg = st.get("fg")
    cell.font = Font(name=old.name, family=old.family, scheme=old.scheme,
                     sz=st.get("size", DEFAULT_SIZE), b=st.get("bold", False), i=st.get("italic", False),
                     u="single" if st.get("underline") else None, strike=st.get("strike", False),
                     color="FF" + fg[1:].upper() if fg else None)
    cell.fill = PatternFill("solid", fgColor="FF" + st["bg"][1:].upper()) if st.get("bg") else PatternFill()
    cell.alignment = Alignment(horizontal=st.get("align"), vertical=st.get("valign"), wrap_text=bool(st.get("wrap")),
                               indent=al.indent, text_rotation=al.text_rotation, shrink_to_fit=al.shrink_to_fit)
    cell.number_format = st.get("numfmt", "General")


def rebuild_sheet(ws, sheet):
    """Rows/columns were inserted or deleted: lay the sheet out again from the model, moving each
    original cell's full style (borders included), row heights and column widths to its new place."""
    from copy import copy

    from openpyxl.comments import Comment
    from openpyxl.utils import get_column_letter
    old = {}
    for row in ws.iter_rows():
        for cell in row:
            if type(cell).__name__ != "MergedCell" and (cell.has_style or cell.comment or cell.hyperlink):
                old[(cell.row - 1, cell.column - 1)] = (copy(cell._style), cell.comment, cell.hyperlink)
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
                ws.cell(row=r + 1, column=c + 1).value = to_value(t)
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


def patch_xlsx(src, dst, values, styles, book=None, widths=None):
    """Write only changed cells into a copy of src: keeps everything else (formulas, borders...).
    values: {(sheet, row, col): text}; styles: {(sheet, row, col): style dict or None}.
    Sheets in `book` whose rows/columns moved are rebuilt instead."""
    import openpyxl
    keep_vba = src.lower().endswith(".xlsm") and dst.lower().endswith(".xlsm")
    wb = openpyxl.load_workbook(src, keep_vba=keep_vba)
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
    wb.save(dst)


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
    for r1, c1, r2, c2 in sheet.merges:
        ws.merge_cells(start_row=r1 + 1, start_column=c1 + 1, end_row=r2 + 1, end_column=c2 + 1)
    for c, w in sheet.widths.items():
        ws.column_dimensions[get_column_letter(c + 1)].width = w
    for r, h in sheet.heights.items():
        ws.row_dimensions[r + 1].height = h
    apply_hidden(ws, sheet)


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
    for r, d in list(ws.row_dimensions.items()):
        d.hidden = (r - 1) in sheet.hidden_rows
    for r in sheet.hidden_rows:
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


def write_xlsx(dst, sheets):
    import openpyxl
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, sheet in sheets.items():
        ws = wb.create_sheet(name[:31])
        fill_sheet(ws, sheet)
        set_state(ws, sheet)
    wb.active = next(i for i, s in enumerate(sheets.values()) if not s.hidden)
    wb.save(dst)


def write_csv(dst, rows, like=None):
    """like: an existing csv whose encoding/delimiter to keep."""
    enc, dialect = sniff_csv(like) if like and os.path.exists(like) else ("utf-8", csv.excel)
    if dst.lower().endswith(".tsv") and not like:
        dialect = csv.excel_tab
    with open(dst, "w", newline="", encoding=enc) as f:
        csv.writer(f, dialect).writerows(trim(rows))
