"""OpenDocument (.ods) reading helpers: cell text and formulas (OpenFormula -> Excel style), cell styles,
number formats, column widths, row heights. Used by sheet_io.load_ods."""
import re

from formatting import DEFAULT_SIZE

NS = {
    "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "style": "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
    "fo": "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    "number": "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",
}
T, OFC, S, FO, TX, N = ("{%s}" % NS[k] for k in ("table", "office", "style", "fo", "text", "number"))
PREFIXES = ("COM.MICROSOFT.", "ORG.OPENOFFICE.", "ORG.LIBREOFFICE.", "_XLFN.")


# ---------- cell values ----------

def para_text(p):
    """Text of a <text:p>, with <text:s/> spaces, tabs and line breaks."""
    out = [p.text or ""]
    for ch in p:
        tag = ch.tag
        if tag == TX + "s":
            out.append(" " * int(ch.get(TX + "c", 1)))
        elif tag == TX + "tab":
            out.append("\t")
        elif tag == TX + "line-break":
            out.append("\n")
        else:
            out.append(para_text(ch))
        out.append(ch.tail or "")
    return "".join(out)


def cell_text(c):
    """What the grid holds for a cell: a formula ('=...'), a plain number / ISO date, or the text."""
    f = c.get(T + "formula")
    if f:
        try:
            return excel_formula(f)
        except ValueError:
            pass
    kind = c.get(OFC + "value-type")
    if kind in ("float", "percentage", "currency"):
        v = c.get(OFC + "value")
        if v is not None:
            n = float(v)
            return str(int(n)) if n.is_integer() and abs(n) < 1e15 else repr(n)
    if kind == "date" and c.get(OFC + "date-value"):
        v = c.get(OFC + "date-value")
        d, _, t = v.partition("T")
        t = t.split(".")[0]
        return d if not t or t == "00:00:00" else f"{d} {t}"
    if kind == "boolean":
        return "TRUE" if c.get(OFC + "boolean-value") in ("true", "1") else "FALSE"
    return "\n".join(para_text(p) for p in c if p.tag == TX + "p")


def _ref(inner):
    """'$Sheet1.$A$1:.B2' -> 'Sheet1!$A$1:B2' (sheet from the first part only)."""
    from formula import quote_sheet
    parts, sheet = [], None
    for i, part in enumerate(inner.split(":")):
        if "." not in part:
            raise ValueError(inner)
        s, _, cell = part.rpartition(".")
        s = s.lstrip("$")
        if s.startswith("'") and s.endswith("'"):
            s = s[1:-1].replace("''", "'")
        if i == 0 and s:
            sheet = s
        parts.append(cell)
    ref = ":".join(parts)
    return f"{quote_sheet(sheet)}!{ref}" if sheet else ref


def excel_formula(f):
    """OpenFormula ('of:=SUM([.A1:.B2];2)') -> Excel formula ('=SUM(A1:B2,2)')."""
    f = re.sub(r"^[a-z]+:", "", f, count=1)          # namespace prefix (of:, msoxl:)
    if not f.startswith("="):
        f = "=" + f
    out, i = [], 0
    while i < len(f):
        ch = f[i]
        if ch == '"':                                   # string: copied as it is
            j = i + 1
            while j < len(f):
                if f[j] == '"':
                    if f[j + 1:j + 2] == '"':
                        j += 2
                        continue
                    break
                j += 1
            out.append(f[i:j + 1])
            i = j + 1
        elif ch == "[":
            j = f.index("]", i)
            out.append(_ref(f[i + 1:j]))
            i = j + 1
        elif ch == ";":
            out.append(",")
            i += 1
        else:
            m = re.match(r"[A-Za-z_][A-Za-z0-9_.]*", f[i:])
            if m:
                word = m.group()
                up = word.upper()
                for p in PREFIXES:
                    if up.startswith(p):
                        word = word[len(p):]
                        break
                out.append(word)
                i += m.end()
            else:
                out.append(ch)
                i += 1
    return "".join(out)


# ---------- styles ----------

def length_pt(v):
    """'2.258cm' / '0.889in' / '64pt' / '15mm' / '20px' -> points."""
    m = re.fullmatch(r"\s*([\d.]+)\s*(cm|mm|in|pt|pc|px)?\s*", v or "")
    if not m:
        return None
    n = float(m.group(1))
    return n * {"cm": 72 / 2.54, "mm": 72 / 25.4, "in": 72, "pt": 1, "pc": 12, "px": 0.75}.get(m.group(2) or "pt")


def border(v):
    """'0.74pt solid #000000' -> ['thin', '#000000'] (None for 'none')."""
    if not v or v.strip() == "none":
        return None
    parts = v.split()
    width = next((length_pt(p) for p in parts if length_pt(p) is not None), 0.75) or 0.75
    kind = next((p for p in parts if p in ("solid", "dashed", "dotted", "double")), "solid")
    color = next((p for p in parts if p.startswith("#")), "#000000")
    if kind == "double":
        name = "double"
    elif kind == "dashed":
        name = "mediumDashed" if width >= 1.5 else "dashed"
    elif kind == "dotted":
        name = "dotted"
    else:
        name = "thick" if width >= 2.5 else "medium" if width >= 1.5 else "thin"
    return [name, color.lower()]


class Styles:
    def __init__(self, content, styles_root):
        self.by_family = {}
        for root, _auto_only in ((styles_root, False), (content, True)):
            if root is None:
                continue
            for el in root.iter():
                if el.tag == S + "style":
                    fam = el.get(S + "family")
                    self.by_family.setdefault(fam, {})[el.get(S + "name")] = el
                elif el.tag.startswith(N) and el.tag.endswith("-style") and el.get(S + "name"):
                    self.by_family.setdefault("data", {})[el.get(S + "name")] = el
        self.cache = {}

    def _chain(self, family, name):
        """The style and its parents, base first (the 'Default' look is left as ours)."""
        out, seen = [], set()
        styles = self.by_family.get(family, {})
        while name and name not in seen and name in styles:
            seen.add(name)
            if name != "Default":
                out.append(styles[name])
            name = styles[name].get(S + "parent-style-name")
        return out[::-1]

    def cell(self, name):
        if not name:
            return None
        if name in self.cache:
            return self.cache[name]
        st = {}
        for el in self._chain("table-cell", name):
            for p in el:
                if p.tag == S + "text-properties":
                    self._text(p, st)
                elif p.tag == S + "table-cell-properties":
                    self._cellprops(p, st)
                elif p.tag == S + "paragraph-properties":
                    a = p.get(FO + "text-align")
                    if a and el.get(S + "name") != "Default":
                        st["align"] = {"start": "left", "end": "right", "left": "left", "right": "right",
                                       "center": "center", "justify": "justify"}.get(a, st.get("align"))
            data = el.get(S + "data-style-name")
            if data:
                code = self.numfmt(data)
                if code:
                    st["numfmt"] = code
        self.cache[name] = st or None
        return self.cache[name]

    @staticmethod
    def _text(p, st):
        if p.get(FO + "font-weight") in ("bold", "600", "700", "800", "900"):
            st["bold"] = True
        if p.get(FO + "font-style") == "italic":
            st["italic"] = True
        if p.get(S + "text-underline-style") not in (None, "none"):
            st["underline"] = True
        if p.get(S + "text-line-through-style") not in (None, "none"):
            st["strike"] = True
        color = p.get(FO + "color")
        if color and color.lower() != "#000000":
            st["fg"] = color.lower()
        size = length_pt(p.get(FO + "font-size"))
        if size and size != DEFAULT_SIZE:
            st["size"] = round(size, 1)
        font = p.get(S + "font-name")
        if font:
            st["font"] = font

    @staticmethod
    def _cellprops(p, st):
        bg = p.get(FO + "background-color")
        if bg and bg != "transparent":
            st["bg"] = bg.lower()
        if p.get(FO + "wrap-option") == "wrap":
            st["wrap"] = True
        v = p.get(S + "vertical-align")
        if v in ("top", "middle"):
            st["valign"] = "top" if v == "top" else "center"
        sides = {}
        every = border(p.get(FO + "border"))
        for side in ("top", "bottom", "left", "right"):
            b = border(p.get(FO + "border-" + side)) or every
            if b:
                sides[side] = b
        if sides:
            st["border"] = sides

    def numfmt(self, name):
        el = self.by_family.get("data", {}).get(name)
        if el is None:
            return None
        kind = el.tag[len(N):]
        if kind in ("number-style", "percentage-style", "currency-style"):
            out = ""
            for ch in el:
                tag = ch.tag[len(N):]
                if tag == "number":
                    dp = int(ch.get(N + "decimal-places", 0))
                    out += ("#,##0" if ch.get(N + "grouping") == "true" else "0") + ("." + "0" * dp if dp else "")
                elif tag == "scientific-number":
                    dp = int(ch.get(N + "decimal-places", 0))
                    out += "0" + ("." + "0" * dp if dp else "") + "E+00"
                elif tag == "currency-symbol":
                    sym = (ch.text or "").strip()
                    out += sym if sym in ("$", "€", "£", "¥") else f'"{sym}"' if sym else ""
                elif tag == "text":
                    t = ch.text or ""
                    out += "%" if t.strip() == "%" else f'"{t}"' if t.strip() else t
            out = out.strip()
            return None if out in ("", "0") and kind == "number-style" else out or None
        if kind in ("date-style", "time-style"):
            out = ""
            for ch in el:
                tag, long_ = ch.tag[len(N):], ch.get(N + "style") == "long"
                if tag == "day":
                    out += "dd" if long_ else "d"
                elif tag == "month":
                    out += ("mmmm" if long_ else "mmm") if ch.get(N + "textual") == "true" else ("mm" if long_ else "m")
                elif tag == "year":
                    out += "yyyy" if long_ else "yy"
                elif tag == "day-of-week":
                    out += "dddd" if long_ else "ddd"
                elif tag == "hours":
                    out += "hh" if long_ else "h"
                elif tag == "minutes":
                    out += "mm"
                elif tag == "seconds":
                    out += "ss"
                elif tag == "am-pm":
                    out += "AM/PM"
                elif tag == "text":
                    out += ch.text or ""
            return out or None
        return None

    def row(self, name):
        out = {}
        for el in self._chain("table-row", name):
            for p in el:
                if p.tag == S + "table-row-properties" and p.get(S + "use-optimal-row-height") == "false":
                    h = length_pt(p.get(S + "row-height"))
                    if h:
                        out["height"] = round(h, 2)
        return out

    def columns(self, table):
        """(widths in Excel characters, hidden column indexes) from <table:table-column>s."""
        widths, hidden, c = {}, set(), 0
        for col in table.iter(T + "table-column"):
            rep = int(col.get(T + "number-columns-repeated", 1))
            w = None
            for el in self._chain("table-column", col.get(T + "style-name")):
                for p in el:
                    if p.tag == S + "table-column-properties":
                        w = length_pt(p.get(S + "column-width")) or w
            for k in range(c, min(c + rep, 500)):
                if col.get(T + "visibility") in ("collapse", "filter"):
                    hidden.add(k)
                elif w:
                    widths[k] = round(max((w / 0.75 - 5) / 7, 0.5), 2)
            c += rep
        return widths, hidden

    def table_hidden(self, table):
        for el in self._chain("table", table.get(T + "style-name")):
            for p in el:
                if p.tag == S + "table-properties" and p.get(T + "display") == "false":
                    return True
        return False
