"""Write .ods (OpenDocument spreadsheet) with the standard library: values, formulas (Excel syntax ->
OpenFormula, with calculated results), number formats, fonts, colours, borders, alignment, merges,
column widths, row heights, hidden rows / columns / sheets."""
import datetime as dt
import zipfile
from xml.sax.saxutils import escape, quoteattr

from formatting import DEFAULT_SIZE, parse_date, parse_number
from formula import XLError, lex, split_sheet

NS = ('xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
      'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" '
      'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" '
      'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" '
      'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" '
      'xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" '
      'xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" office:version="1.2"')
BORDER = {"thin": "0.74pt solid", "hair": "0.5pt solid", "dotted": "0.74pt dotted", "dashed": "0.74pt dashed",
          "medium": "1.75pt solid", "mediumDashed": "1.75pt dashed", "thick": "2.5pt solid", "double": "2.5pt double"}


def _ref(tok):
    sheet, ref = split_sheet(tok)
    parts = ref.split(":")
    s = ""
    if sheet is not None:
        s = "$'" + sheet.replace("'", "''") + "'" if not sheet.replace("_", "").isalnum() else "$" + sheet
    return "[" + s + "." + parts[0] + (":." + parts[1] if len(parts) > 1 else "") + "]"


def of_formula(f):
    """'=SUM(A1:B2,Sheet2!C3)' -> 'of:=SUM([.A1:.B2];[$Sheet2.C3])'."""
    out = []
    for kind, tok in lex(f[1:]):
        if kind in ("ref", "cols", "rows"):
            out.append(_ref(tok))
        elif kind == "op" and tok == ",":
            out.append(";")
        else:
            out.append(tok)
    return "of:=" + "".join(out)


def data_style(code, name):
    """Excel number format -> <number:*-style> (None for General / unknown)."""
    if not code or code == "General":
        return None
    low = code.lower().split(";")[0]
    if any(ch in low for ch in "dmyhs") and not low.startswith('"'):
        parts, i = [], 0
        tokens = [("yyyy", '<number:year number:style="long"/>'), ("yy", "<number:year/>"),
                  ("mmmm", '<number:month number:textual="true" number:style="long"/>'),
                  ("mmm", '<number:month number:textual="true"/>'), ("mm", '<number:month number:style="long"/>'),
                  ("m", "<number:month/>"), ("dddd", '<number:day-of-week number:style="long"/>'),
                  ("ddd", "<number:day-of-week/>"), ("dd", '<number:day number:style="long"/>'), ("d", "<number:day/>"),
                  ("hh", '<number:hours number:style="long"/>'), ("h", "<number:hours/>"),
                  ("ss", '<number:seconds number:style="long"/>'), ("s", "<number:seconds/>")]
        last = None   # last date/time part written: "m" right after hours means minutes
        while i < len(low):
            for t, xml in tokens:
                if low.startswith(t, i):
                    if t in ("mm", "m") and last == "h":
                        xml = '<number:minutes number:style="long"/>'
                    parts.append(xml)
                    last = t[0]
                    i += len(t)
                    break
            else:
                ch = code[i]
                if ch not in '"\\':
                    parts.append(f"<number:text>{escape(ch)}</number:text>")
                i += 1
        return f'<number:date-style style:name="{name}">{"".join(parts)}</number:date-style>'
    dec = sum(1 for ch in low.split(".")[1] if ch in "0#") if "." in low else 0
    group = "true" if "," in low else "false"
    num = f'<number:number number:decimal-places="{dec}" number:min-integer-digits="1" number:grouping="{group}"/>'
    if "%" in low:
        return f'<number:percentage-style style:name="{name}">{num}<number:text>%</number:text></number:percentage-style>'
    sym = next((s for s in ("$", "€", "£", "¥") if s in code), None)
    if sym:
        return (f'<number:currency-style style:name="{name}"><number:currency-symbol>{sym}</number:currency-symbol>'
                f'{num}</number:currency-style>')
    return f'<number:number-style style:name="{name}">{num}</number:number-style>'


def cell_props(st, data_name):
    text, cell, para = [], [], []
    if st.get("bold"):
        text.append('fo:font-weight="bold"')
    if st.get("italic"):
        text.append('fo:font-style="italic"')
    if st.get("underline"):
        text.append('style:text-underline-style="solid" style:text-underline-width="auto" '
                    'style:text-underline-color="font-color"')
    if st.get("strike"):
        text.append('style:text-line-through-style="solid"')
    if st.get("fg"):
        text.append(f'fo:color="{st["fg"]}"')
    if st.get("size") and st["size"] != DEFAULT_SIZE:
        text.append(f'fo:font-size="{st["size"]:g}pt"')
    if st.get("font"):
        text.append(f'style:font-name={quoteattr(st["font"])}')
    if st.get("bg"):
        cell.append(f'fo:background-color="{st["bg"]}"')
    if st.get("wrap"):
        cell.append('fo:wrap-option="wrap"')
    if st.get("valign"):
        cell.append(f'style:vertical-align="{"top" if st["valign"] == "top" else "middle"}"')
    for side, (kind, color) in (st.get("border") or {}).items():
        cell.append(f'fo:border-{side}="{BORDER.get(kind, "0.74pt solid")} {color}"')
    if st.get("align"):
        para.append(f'fo:text-align="{ {"left": "start", "right": "end"}.get(st["align"], st["align"]) }"')
        cell.append('style:text-align-source="fix"')
    out = ""
    if cell:
        out += f"<style:table-cell-properties {' '.join(cell)}/>"
    if para:
        out += f"<style:paragraph-properties {' '.join(para)}/>"
    if text:
        out += f"<style:text-properties {' '.join(text)}/>"
    data = f' style:data-style-name="{data_name}"' if data_name else ""
    return data, out


def value_attrs(v, shown):
    """Typed value -> (attributes, text shown)."""
    if isinstance(v, bool):
        return f'office:value-type="boolean" office:boolean-value="{"true" if v else "false"}"', "TRUE" if v else "FALSE"
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return f'office:value-type="float" office:value="{float(v)!r}"', shown
    if isinstance(v, (dt.date, dt.datetime)):
        iso = v.isoformat() if isinstance(v, dt.datetime) else v.isoformat()
        return f'office:value-type="date" office:date-value="{iso}"', shown
    return 'office:value-type="string"', shown


def cell_value(text):
    """Grid text -> typed value."""
    n = parse_number(text)
    if n is not None:
        return n
    if text.upper() in ("TRUE", "FALSE"):
        return text.upper() == "TRUE"
    d = parse_date(text)
    if d is not None:
        return d
    return text


def write(dst, app, sheets):
    eng = app.engine
    styles, style_xml, data_xml, col_styles, row_styles = {}, [], [], {}, {}

    def style_name(st):
        if not st:
            return None
        key = repr(sorted(st.items()))
        if key not in styles:
            n = len(styles) + 1
            data = None
            if st.get("numfmt"):
                data = f"N{n}"
                d = data_style(st["numfmt"], data)
                if d:
                    data_xml.append(d)
                else:
                    data = None
            attr, props = cell_props(st, data)
            styles[key] = f"ce{n}"
            style_xml.append(f'<style:style style:name="ce{n}" style:family="table-cell" '
                             f'style:parent-style-name="Default"{attr}>{props}</style:style>')
        return styles[key]

    def col_style(width):
        k = round(width, 2)
        if k not in col_styles:
            col_styles[k] = f"co{len(col_styles) + 1}"
            pt = (width * 7 + 5) * 0.75
            style_xml.append(f'<style:style style:name="{col_styles[k]}" style:family="table-column">'
                             f'<style:table-column-properties style:column-width="{pt:.2f}pt"/></style:style>')
        return col_styles[k]

    def row_style(height):
        k = round(height, 2)
        if k not in row_styles:
            row_styles[k] = f"ro{len(row_styles) + 1}"
            style_xml.append(f'<style:style style:name="{row_styles[k]}" style:family="table-row">'
                             f'<style:table-row-properties style:row-height="{height:.2f}pt" '
                             f'style:use-optimal-row-height="false"/></style:style>')
        return row_styles[k]

    style_xml.append('<style:style style:name="taHidden" style:family="table">'
                     '<style:table-properties table:display="false"/></style:style>')
    tables = []
    for name, s in sheets.items():
        rows = s.rows
        ncols = max([len(r) for r in rows] + [c + 1 for _, c in s.styles] + [0])
        covered = {}
        for r1, c1, r2, c2 in s.merges:
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    covered[(r, c)] = (r1, c1, r2, c2)
        parts = [f'<table:table table:name={quoteattr(name)}' + (' table:style-name="taHidden"' if s.hidden else "") + ">"]
        for c in range(ncols):
            attrs = f' table:style-name="{col_style(s.widths[c])}"' if c in s.widths else ""
            if c in s.hidden_cols:
                attrs += ' table:visibility="collapse"'
            parts.append(f"<table:table-column{attrs}/>")
        nrows = max(len(rows), max([r + 1 for r, _ in s.styles] + [0]))
        for r in range(nrows):
            attrs = f' table:style-name="{row_style(s.heights[r])}"' if r in s.heights else ""
            if r in s.hidden_rows or r in s.filtered_rows:
                attrs += ' table:visibility="collapse"'
            cells = []
            row = rows[r] if r < len(rows) else []
            for c in range(ncols):
                m = covered.get((r, c))
                if m and (r, c) != m[:2]:
                    cells.append("<table:covered-table-cell/>")
                    continue
                t = row[c] if c < len(row) else ""
                st = s.styles.get((r, c))
                sn = style_name(st)
                a = f' table:style-name="{sn}"' if sn else ""
                if m:
                    a += f' table:number-rows-spanned="{m[2] - m[0] + 1}" table:number-columns-spanned="{m[3] - m[1] + 1}"'
                if not t:
                    cells.append(f"<table:table-cell{a}/>")
                    continue
                if t.startswith("=") and len(t) > 1:
                    v = eng.cell(name, r, c)
                    shown = eng.display(s, r, c)
                    a += f" table:formula={quoteattr(of_formula(t))}"
                    if isinstance(v, XLError) or v is None:
                        v = shown
                else:
                    v, shown = cell_value(t), t
                va, shown = value_attrs(v, shown)
                body = "".join(f"<text:p>{escape(line)}</text:p>" for line in str(shown).split("\n"))
                cells.append(f"<table:table-cell{a} {va}>{body}</table:table-cell>")
            while cells and cells[-1] == "<table:table-cell/>":
                cells.pop()
            parts.append(f"<table:table-row{attrs}>{''.join(cells) or '<table:table-cell/>'}</table:table-row>")
        parts.append("</table:table>")
        tables.append("".join(parts))
    content = (f'<?xml version="1.0" encoding="UTF-8"?><office:document-content {NS}>'
               f"<office:automatic-styles>{''.join(data_xml)}{''.join(style_xml)}</office:automatic-styles>"
               f"<office:body><office:spreadsheet>{''.join(tables)}</office:spreadsheet></office:body>"
               "</office:document-content>")
    styles_doc = (f'<?xml version="1.0" encoding="UTF-8"?><office:document-styles {NS}><office:styles>'
                  '<style:style style:name="Default" style:family="table-cell"/></office:styles></office:document-styles>')
    manifest = ('<?xml version="1.0" encoding="UTF-8"?>'
                '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">'
                '<manifest:file-entry manifest:full-path="/" '
                'manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>'
                '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>'
                '<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>'
                '</manifest:manifest>')
    tmp = dst + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(zipfile.ZipInfo("mimetype"), "application/vnd.oasis.opendocument.spreadsheet",
                   compress_type=zipfile.ZIP_STORED)
        z.writestr("META-INF/manifest.xml", manifest)
        z.writestr("content.xml", content)
        z.writestr("styles.xml", styles_doc)
    import os
    os.replace(tmp, dst)
    return dst
