"""Links to other workbooks in formulas: '[Prices.xlsx]Sheet1'!B2, or [1]Sheet1!B2 as stored in .xlsx files.
The other file is read (and calculated) when it can be found next to this one or at its saved path;
otherwise the values Excel stored with the link are used."""
import os
import re
import xml.etree.ElementTree as ET
import zipfile

from formula import NA, REF, Array, FormulaError, col_index, literal

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
REF_RE = re.compile(r"^'?\[([^\]]+)\]([^'!]*)'?!(\$?[A-Za-z]+\$?\d+)(?::(\$?[A-Za-z]+\$?\d+))?$")


def cell_pos(ref):
    m = re.fullmatch(r"\$?([A-Za-z]+)\$?(\d+)", ref)
    return int(m.group(2)) - 1, col_index(m.group(1))


def stored_links(path):
    """[{'target': file path, 'sheets': [names], 'cache': {sheet: {(r, c): value}}}] in link order."""
    out = []
    try:
        z = zipfile.ZipFile(path)
    except (OSError, zipfile.BadZipFile):
        return out
    with z:
        names = set(z.namelist())
        i = 1
        while f"xl/externalLinks/externalLink{i}.xml" in names:
            root = ET.fromstring(z.read(f"xl/externalLinks/externalLink{i}.xml"))
            target = ""
            rels = f"xl/externalLinks/_rels/externalLink{i}.xml.rels"
            if rels in names:
                r = ET.fromstring(z.read(rels))
                rel = next(iter(r), None)
                if rel is not None:
                    target = rel.get("Target", "")
            sheets = [s.get("val") for s in root.iter(NS + "sheetName")]
            cache = {}
            for sd in root.iter(NS + "sheetData"):
                sid = int(sd.get("sheetId", 0))
                name = sheets[sid] if sid < len(sheets) else str(sid)
                cells = cache.setdefault(name, {})
                for c in sd.iter(NS + "cell"):
                    v = c.find(NS + "v")
                    if v is not None and c.get("r"):
                        t = v.text or ""
                        cells[cell_pos(c.get("r"))] = t if c.get("t") == "str" else literal(t)
            out.append({"target": target, "sheets": sheets, "cache": cache})
            i += 1
    return out


class Links:
    def __init__(self, app, path):
        self.app, self.path = app, path
        self.stored = None
        self.books = {}          # file path -> (Engine, {sheet name: Sheet}) or None if unreadable

    def find(self, name):
        """The other workbook's file: as given, or next to this file."""
        name = name.replace("file:///", "").replace("/", os.sep)
        here = os.path.dirname(self.path) if self.path else os.getcwd()
        for p in (name, os.path.join(here, name), os.path.join(here, os.path.basename(name))):
            if p and os.path.isfile(p):
                return os.path.abspath(p)
        return None

    def book(self, file):
        if file not in self.books:
            try:
                import sheet_io
                from formula import Engine
                sheets = sheet_io.load(file)
                eng = Engine()
                eng.set_book(sheets)
                self.books[file] = (eng, sheets)
            except Exception:
                self.books[file] = None
        return self.books[file]

    def value(self, tok):
        m = REF_RE.match(tok)
        if not m:
            raise FormulaError(REF)
        book, sheet, a, b = m.groups()
        cached = None
        if book.isdigit():                       # [1]Sheet1!A1: the file's first external link
            if self.stored is None:
                self.stored = stored_links(self.path) if self.path else []
            k = int(book) - 1
            if k >= len(self.stored):
                raise FormulaError(REF)
            link = self.stored[k]
            book, cached = link["target"], link["cache"]
        (r1, c1), (r2, c2) = cell_pos(a), cell_pos(b or a)
        file = self.find(book)
        loaded = self.book(file) if file else None
        if loaded:
            eng, sheets = loaded
            real = next((n for n in sheets if n.lower() == sheet.lower()), None) or (next(iter(sheets)) if not sheet else None)
            if real is None:
                raise FormulaError(REF)
            get = lambda r, c: eng.cell(real, r, c)   # noqa: E731
        else:
            if cached is None:
                if self.stored is None:
                    self.stored = stored_links(self.path) if self.path else []
                want = os.path.basename(book).lower()
                link = next((s for s in self.stored if os.path.basename(s["target"]).lower() == want), None)
                cached = link["cache"] if link else None
            if cached is None:
                raise FormulaError(REF)
            values = next((v for n, v in cached.items() if n.lower() == sheet.lower()), {})
            get = lambda r, c: values.get((r, c), NA if not values else None)   # noqa: E731
        if (r1, c1) == (r2, c2):
            return get(r1, c1)
        return Array([[get(r, c) for c in range(c1, c2 + 1)] for r in range(r1, r2 + 1)])

    def refresh(self):
        """Read the other workbooks again (Data > Update links)."""
        self.books.clear()
        self.stored = None
