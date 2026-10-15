"""Export the current sheet as a stand-alone HTML page (values as shown, with fonts, colours, borders,
alignment, merged cells; hidden rows / columns left out)."""
import html
import os
from tkinter import filedialog, messagebox

from formatting import DEFAULT_SIZE, parse_number, render

BORDER_CSS = {"thin": "1px solid", "hair": "1px dotted", "dotted": "1px dotted", "dashed": "1px dashed",
              "medium": "2px solid", "mediumDashed": "2px dashed", "thick": "3px solid", "double": "3px double"}


def cell_css(st, is_num):
    css = []
    if st.get("bold"):
        css.append("font-weight:bold")
    if st.get("italic"):
        css.append("font-style:italic")
    deco = [d for k, d in (("underline", "underline"), ("strike", "line-through")) if st.get(k)]
    if deco:
        css.append("text-decoration:" + " ".join(deco))
    if st.get("size") and st["size"] != DEFAULT_SIZE:
        css.append(f"font-size:{st['size']:g}pt")
    if st.get("font"):
        css.append(f"font-family:'{st['font']}'")
    if st.get("fg"):
        css.append(f"color:{st['fg']}")
    if st.get("bg"):
        css.append(f"background:{st['bg']}")
    align = st.get("align") or ("right" if is_num else None)
    if align:
        css.append(f"text-align:{align}")
    if st.get("valign"):
        css.append("vertical-align:" + ("top" if st["valign"] == "top" else "middle"))
    if st.get("wrap"):
        css.append("white-space:pre-wrap")
    for side, (kind, color) in (st.get("border") or {}).items():
        css.append(f"border-{side}:{BORDER_CSS.get(kind, '1px solid')} {color}")
    return ";".join(css)


def sheet_html(app, name):
    s = app.sheets[name]
    eng = app.engine
    rows = [r for r in range(len(s.rows)) if r not in s.hidden_rows and r not in s.filtered_rows]
    ncols = max((len(s.rows[r]) for r in rows), default=0)
    cols = [c for c in range(ncols) if c not in s.hidden_cols]
    merge_at = {}
    for m in s.merges:
        for r in range(m[0], m[2] + 1):
            for c in range(m[1], m[3] + 1):
                merge_at[(r, c)] = m
    out = ["<table>\n<colgroup>"]
    out += [f'<col style="width:{(s.widths.get(c, 9.43) * 7 + 5):.0f}px">' for c in cols]
    out.append("</colgroup>")
    for r in rows:
        h = f' style="height:{s.heights[r] * 96 / 72:.0f}px"' if r in s.heights else ""
        out.append(f"<tr{h}>")
        for c in cols:
            m = merge_at.get((r, c))
            if m and (r, c) != m[:2]:
                continue
            t = s.rows[r][c] if c < len(s.rows[r]) else ""
            v = eng.display(s, r, c) if len(t) > 1 and t[0] == "=" else t
            st = s.styles.get((r, c)) or {}
            cf = app.rules.conditional(s, r, c)
            if cf:
                cf.pop("bar", None)
                st = {**st, **cf}
            text = render(v, st.get("numfmt")) if st.get("numfmt") else v
            attrs = ""
            if m:
                span_r = len([x for x in rows if m[0] <= x <= m[2]])
                span_c = len([x for x in cols if m[1] <= x <= m[3]])
                attrs += (f' rowspan="{span_r}"' if span_r > 1 else "") + (f' colspan="{span_c}"' if span_c > 1 else "")
            css = cell_css(st, parse_number(v) is not None)
            if css:
                attrs += f' style="{html.escape(css, quote=True)}"'
            body = html.escape(text).replace("\n", "<br>")
            link = s.links.get((r, c))
            if link and not link.startswith("#"):
                body = f'<a href="{html.escape(link, quote=True)}">{body}</a>'
            out.append(f"<td{attrs}>{body}</td>")
        out.append("</tr>")
    out.append("</table>")
    return "\n".join(out)


PAGE = """<!doctype html>
<html><head><meta charset="utf-8"><title>{title}</title>
<style>
body {{ font-family: Calibri, Arial, sans-serif; font-size: 11pt; margin: 16px; }}
table {{ border-collapse: collapse; table-layout: fixed; }}
td {{ border: 1px solid #e3e3e3; padding: 1px 4px; white-space: nowrap; overflow: hidden; }}
</style></head>
<body>
{body}
</body></html>
"""


def export(app, dst, name=None):
    name = name or app.sheet
    data = PAGE.format(title=html.escape(name), body=sheet_html(app, name))
    with open(dst, "w", encoding="utf-8") as f:
        f.write(data)
    return dst


def export_dialog(app, dst=None):
    if dst is None:
        named = app.path or getattr(app, "recovered_from", None)
        stem = os.path.splitext(os.path.basename(named))[0] if named else "Untitled"
        dst = filedialog.asksaveasfilename(filetypes=[("Web page", "*.html")], defaultextension=".html",
                                           initialfile=f"{stem} - {app.sheet}.html", parent=app)
        if not dst:
            return None
    try:
        export(app, dst)
    except OSError as e:
        messagebox.showerror("Export", str(e), parent=app)
        return None
    app.status["text"] = f"Exported {os.path.basename(dst)}"
    return dst
