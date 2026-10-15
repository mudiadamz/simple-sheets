"""Data → Import: a text/CSV wizard (encoding, delimiter, column types, preview) and tables from a web page."""
import csv
import io
import re
import tkinter as tk
from html.parser import HTMLParser
from tkinter import filedialog, messagebox, ttk

from formatting import parse_number

ENCODINGS = ["utf-8", "utf-8-sig", "utf-16", "cp1252", "latin-1", "cp1250", "cp1251", "shift_jis", "gbk"]
DELIMITERS = {"Comma": ",", "Semicolon": ";", "Tab": "\t", "Pipe": "|", "Space": " "}
TYPES = ["General", "Text", "Date (D/M/Y)", "Date (M/D/Y)", "Date (Y/M/D)", "Skip"]


def guess_encoding(raw):
    if raw.startswith(b"\xef\xbb\xbf"):
        return "utf-8-sig"
    if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        return "utf-16"
    try:
        raw.decode("utf-8")
        return "utf-8"
    except UnicodeDecodeError:
        return "cp1252"


def guess_delimiter(text):
    try:
        return csv.Sniffer().sniff(text[:4096], delimiters=",;\t|").delimiter
    except csv.Error:
        return "\t" if "\t" in text[:4096] else ","


def to_date(text, order):
    """'31/12/2024' with order 'DMY' -> '2024-12-31' (unchanged when it isn't a date)."""
    parts = re.split(r"[/.\-]", text.strip())
    if len(parts) != 3 or not all(p.isdigit() for p in parts):
        return text
    d = dict(zip(order, (int(p) for p in parts)))
    y = d["Y"] + (2000 if d["Y"] < 50 else 1900 if d["Y"] < 100 else 0)
    if not (1 <= d["M"] <= 12 and 1 <= d["D"] <= 31):
        return text
    return f"{y:04d}-{d['M']:02d}-{d['D']:02d}"


def parse(raw, encoding, delimiter, skip=0, types=None):
    """bytes -> rows of text, with column types applied ('Skip' columns dropped). Returns (rows, text_cols):
    text_cols are output columns typed 'Text' (kept as typed: leading zeros, long numbers)."""
    text = raw.decode(encoding, errors="replace")
    rows = list(csv.reader(io.StringIO(text, newline=""), delimiter=delimiter))[skip:]
    types = types or {}
    keep = [c for c in range(max((len(r) for r in rows), default=0)) if types.get(c, "General") != "Skip"]
    out, text_cols = [], [i for i, c in enumerate(keep) if types.get(c) == "Text"]
    for row in rows:
        new = []
        for c in keep:
            v = row[c] if c < len(row) else ""
            t = types.get(c, "General")
            if t.startswith("Date"):
                v = to_date(v, re.sub(r"[^DMY]", "", t[5:]))
            new.append(v)
        out.append(new)
    return out, text_cols


def place(app, rows, text_cols=(), new_sheet=False, name="Import"):
    """Put imported rows at the cursor or on a new sheet."""
    if not rows:
        return None
    if new_sheet:
        import data_more
        styles = {(r, c): {"numfmt": "@"} for r in range(len(rows)) for c in text_cols}
        for c in range(len(rows[0])):
            styles[(0, c)] = {**styles.get((0, c), {}), "bold": True}
        sheet = data_more.add_table_sheet(app, name, rows, styles)
        app.show(sheet)
        return sheet
    r0, c0 = app.table.cur
    changes = [("v", r0 + i, c0 + j, v) for i, row in enumerate(rows) for j, v in enumerate(row)]
    changes += [("s", r0 + i, c0 + c, {**(app.current.styles.get((r0 + i, c0 + c)) or {}), "numfmt": "@"})
                for i in range(len(rows)) for c in text_cols]
    app.apply_mixed(changes)                       # values and Text formats: one undo step
    return app.sheet


class CsvWizard(tk.Toplevel):
    def __init__(self, app, path):
        super().__init__(app)
        self.app, self.path = app, path
        with open(path, "rb") as f:
            self.raw = f.read()
        self.title("Import text / CSV")
        self.transient(app)
        f = ttk.Frame(self, padding=10)
        f.pack(fill="both", expand=True)
        enc = guess_encoding(self.raw)
        delim = guess_delimiter(self.raw.decode(enc, errors="replace"))
        top = ttk.Frame(f)
        top.pack(fill="x")
        ttk.Label(top, text="Encoding").pack(side="left")
        self.enc = ttk.Combobox(top, values=ENCODINGS, width=10)
        self.enc.set(enc)
        self.enc.pack(side="left", padx=(4, 12))
        ttk.Label(top, text="Delimiter").pack(side="left")
        self.delim = ttk.Combobox(top, values=list(DELIMITERS), width=10)
        self.delim.set(next((k for k, v in DELIMITERS.items() if v == delim), delim))
        self.delim.pack(side="left", padx=(4, 12))
        ttk.Label(top, text="Skip rows").pack(side="left")
        self.skip = ttk.Spinbox(top, from_=0, to=1000, width=5)
        self.skip.set(0)
        self.skip.pack(side="left", padx=4)
        col = ttk.Frame(f)
        col.pack(fill="x", pady=6)
        ttk.Label(col, text="Column").pack(side="left")
        self.col = ttk.Combobox(col, state="readonly", width=16)
        self.col.pack(side="left", padx=4)
        ttk.Label(col, text="type").pack(side="left")
        self.type = ttk.Combobox(col, values=TYPES, state="readonly", width=14)
        self.type.set("General")
        self.type.pack(side="left", padx=4)
        self.types = {}
        self.tree = ttk.Treeview(f, show="headings", height=12)
        self.tree.pack(fill="both", expand=True)
        bottom = ttk.Frame(f)
        bottom.pack(fill="x", pady=(8, 0))
        self.new_sheet = tk.BooleanVar(master=self, value=True)
        ttk.Checkbutton(bottom, text="Put on a new sheet (else at the selected cell)",
                        variable=self.new_sheet).pack(side="left")
        ttk.Button(bottom, text="Cancel", command=self.destroy).pack(side="right")
        ttk.Button(bottom, text="Import", command=self.ok).pack(side="right", padx=4)
        for w in (self.enc, self.delim):
            w.bind("<<ComboboxSelected>>", lambda e: self.preview())
            w.bind("<KeyRelease>", lambda e: self.preview())
        self.skip.configure(command=self.preview)
        self.col.bind("<<ComboboxSelected>>", lambda e: self.type.set(self.types.get(self.col.current(), "General")))
        self.type.bind("<<ComboboxSelected>>", lambda e: self.set_type())
        self.preview()

    def options(self):
        d = self.delim.get()
        try:
            "".encode(self.enc.get())
            enc = self.enc.get()
        except LookupError:
            enc = "utf-8"
        try:
            skip = max(int(self.skip.get()), 0)
        except ValueError:
            skip = 0
        return enc, DELIMITERS.get(d, d[:1] or ","), skip

    def set_type(self):
        if self.col.current() >= 0:
            self.types[self.col.current()] = self.type.get()
            self.preview()

    def preview(self):
        enc, delim, skip = self.options()
        all_rows, _ = parse(self.raw, enc, delim, skip)
        n = max((len(r) for r in all_rows), default=0)
        names = [f"{i + 1}: {all_rows[0][i] if all_rows and i < len(all_rows[0]) else ''}"[:24] for i in range(n)]
        self.col["values"] = names
        if names and self.col.current() < 0:
            self.col.current(0)
        rows, _ = parse(self.raw, enc, delim, skip, self.types)
        cols = [f"c{i}" for i in range(max((len(r) for r in rows), default=0))]
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = cols
        shown = [i for i in range(n) if self.types.get(i, "General") != "Skip"]
        for k, cid in enumerate(cols):
            src = shown[k] if k < len(shown) else k
            self.tree.heading(cid, text=f"{src + 1} · {self.types.get(src, 'General')}")
            self.tree.column(cid, width=110, stretch=False)
        for row in rows[:50]:
            self.tree.insert("", "end", values=row)
        return rows

    def ok(self):
        enc, delim, skip = self.options()
        rows, text_cols = parse(self.raw, enc, delim, skip, self.types)
        new_sheet = self.new_sheet.get()          # read before the window is gone
        self.destroy()
        import os
        name = re.sub(r"[\[\]:*?/\\]", "", os.path.splitext(os.path.basename(self.path))[0])[:28] or "Import"
        return place(self.app, rows, text_cols, new_sheet, name)


def import_csv(app, path=None):
    path = path or filedialog.askopenfilename(parent=app, title="Import text / CSV",
                                              filetypes=[("Text files", "*.csv *.tsv *.txt"), ("All files", "*.*")])
    return CsvWizard(app, path) if path else None


# ---------- tables from a web page ----------

class TableParser(HTMLParser):
    """Collects every <table> as rows of cell text (colspan repeated as blanks)."""
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.tables, self.stack = [], []
        self.cell = None
        self.caption = None

    def handle_starttag(self, tag, attrs):
        if tag == "table":
            self.stack.append({"rows": [], "caption": ""})
        elif not self.stack:
            return
        elif tag == "tr":
            self.stack[-1]["rows"].append([])
        elif tag in ("td", "th"):
            if not self.stack[-1]["rows"]:
                self.stack[-1]["rows"].append([])
            span = dict(attrs).get("colspan") or "1"
            self.cell = {"text": [], "span": int(span) if span.isdigit() else 1}
        elif tag == "br" and self.cell is not None:
            self.cell["text"].append("\n")
        elif tag == "caption":
            self.caption = []

    def handle_endtag(self, tag):
        if not self.stack:
            return
        if tag in ("td", "th") and self.cell is not None:
            text = re.sub(r"[ \t\r\f\v]+", " ", "".join(self.cell["text"])).strip()
            text = "\n".join(line.strip() for line in text.split("\n"))
            self.stack[-1]["rows"][-1] += [text] + [""] * (self.cell["span"] - 1)
            self.cell = None
        elif tag == "caption" and self.caption is not None:
            self.stack[-1]["caption"] = " ".join("".join(self.caption).split())
            self.caption = None
        elif tag == "table":
            t = self.stack.pop()
            t["rows"] = [r for r in t["rows"] if any(x for x in r)]
            if t["rows"]:
                self.tables.append(t)

    def handle_data(self, data):
        if self.cell is not None:
            self.cell["text"].append(data)
        elif self.caption is not None:
            self.caption.append(data)


def parse_tables(html):
    p = TableParser()
    p.feed(html)
    p.close()
    return p.tables


def fetch(url, timeout=20):
    import urllib.request
    if not re.match(r"^(https?|file)://", url, re.I):
        url = "https://" + url
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 SimpleSheets"})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        raw = resp.read()
        charset = resp.headers.get_content_charset() if hasattr(resp, "headers") and resp.headers else None
    m = re.search(rb'<meta[^>]+charset=["\']?([\w-]+)', raw[:4096], re.I)
    enc = charset or (m.group(1).decode() if m else "utf-8")
    try:
        return raw.decode(enc, errors="replace")
    except LookupError:
        return raw.decode("utf-8", errors="replace")


class WebDialog(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app)
        self.app, self.tables = app, []
        self.title("Import a table from the web")
        self.transient(app)
        f = ttk.Frame(self, padding=10)
        f.pack(fill="both", expand=True)
        row = ttk.Frame(f)
        row.pack(fill="x")
        ttk.Label(row, text="Address").pack(side="left")
        self.url = ttk.Entry(row, width=50)
        self.url.pack(side="left", padx=4, fill="x", expand=True)
        self.url.bind("<Return>", lambda e: self.load())
        ttk.Button(row, text="Find tables", command=self.load).pack(side="left")
        self.list = tk.Listbox(f, height=10, width=70, exportselection=False)
        self.list.pack(fill="both", expand=True, pady=6)
        self.list.bind("<Double-Button-1>", lambda e: self.ok())
        b = ttk.Frame(f)
        b.pack(fill="x")
        ttk.Button(b, text="Cancel", command=self.destroy).pack(side="right")
        ttk.Button(b, text="Import", command=self.ok).pack(side="right", padx=4)
        self.url.focus_set()

    def load(self, html=None):
        try:
            html = html if html is not None else fetch(self.url.get().strip())
        except Exception as e:
            messagebox.showwarning("Import from web", f"Couldn't open the page:\n{e}", parent=self)
            return []
        self.tables = parse_tables(html)
        self.list.delete(0, "end")
        for i, t in enumerate(self.tables):
            first = " | ".join(x for x in t["rows"][0] if x)[:60]
            label = t["caption"] or first
            self.list.insert("end", f"Table {i + 1} ({len(t['rows'])} rows): {label}")
        if self.tables:
            self.list.selection_set(0)
        else:
            messagebox.showinfo("Import from web", "No tables on that page.", parent=self)
        return self.tables

    def ok(self):
        sel = self.list.curselection()
        if not sel:
            return None
        t = self.tables[sel[0]]
        rows = [[_number(v) for v in r] for r in t["rows"]]
        self.destroy()
        width = max(len(r) for r in rows)
        return place(self.app, [r + [""] * (width - len(r)) for r in rows], new_sheet=True, name="Web")


def _number(v):
    """'1,234' / '$5' from a page -> '1234' / '5' when it is clearly a number; else unchanged."""
    s = v.replace("−", "-").strip()
    if re.fullmatch(r"[-+]?[$€£¥]?\d{1,3}(,\d{3})+(\.\d+)?", s):
        s = re.sub(r"[$€£¥,]", "", s)
        return s if parse_number(s) is not None else v
    return v


def import_web(app):
    return WebDialog(app)
