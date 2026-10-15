"""Rich text: cells with several fonts (e.g. one bold word). Runs are [[text, props]] with props from
bold / italic / underline / strike / fg / size / font. Shown run by run on the grid; in the cell editor
Ctrl+B / Ctrl+I / Ctrl+U format the selected characters; saved as openpyxl CellRichText."""
from formatting import color_hex

KEYS = ("bold", "italic", "underline", "strike", "fg", "size", "font")


def runs_from(value, theme):
    """openpyxl CellRichText -> runs."""
    out = []
    for block in value:
        if isinstance(block, str):
            out.append([block, {}])
            continue
        f, props = block.font, {}
        if f is not None:
            if f.b:
                props["bold"] = True
            if f.i:
                props["italic"] = True
            if f.u:
                props["underline"] = True
            if f.strike:
                props["strike"] = True
            fg = color_hex(f.color, theme) if f.color is not None else None
            if fg and fg != "#000000":
                props["fg"] = fg
            if f.sz:
                props["size"] = float(f.sz)
            if f.rFont:
                props["font"] = f.rFont
        out.append([block.text, props])
    return merge(out)


def merge(runs):
    """Join neighbouring runs with the same look; drop empty ones."""
    out = []
    for text, props in runs:
        if not text:
            continue
        if out and out[-1][1] == props:
            out[-1][0] += text
        else:
            out.append([text, dict(props)])
    return out


def text_of(runs):
    return "".join(t for t, _ in runs)


def plain(runs):
    return all(not p for _, p in runs)


def to_openpyxl(runs):
    from openpyxl.cell.rich_text import CellRichText, TextBlock
    from openpyxl.cell.text import InlineFont
    parts = []
    for text, p in runs:
        if not p:
            parts.append(text)
            continue
        parts.append(TextBlock(InlineFont(b=p.get("bold") or None, i=p.get("italic") or None,
                                          u="single" if p.get("underline") else None, strike=p.get("strike") or None,
                                          color="FF" + p["fg"][1:].upper() if p.get("fg") else None,
                                          sz=p.get("size"), rFont=p.get("font")), text))
    return CellRichText(parts)


def valid(sheet, r, c, text):
    """Runs for a cell, if it still has exactly that text (edited elsewhere = plain again)."""
    runs = sheet.rich.get((r, c))
    return runs if runs and text_of(runs) == text else None


# ---------- grid ----------

def draw(g, runs, st, x, y, w, h, fg_default, anchor_v):
    """Draw runs one after another on one line, aligned like the cell. Returns True if drawn."""
    pad = g.px(4)
    pieces = []
    for text, p in runs:
        look = {**(st or {}), **p}
        font = g.font(look)
        pieces.append((text.split("\n")[0], font, p.get("fg") or fg_default))
        if "\n" in text:
            break
    total = sum(f.measure(t) for t, f, _ in pieces)
    align = (st or {}).get("align") or "left"
    tx = x + pad if align in ("left", "justify") else x + w - pad - total if align == "right" else x + (w - total) / 2
    vpad = g.px(3)
    ty = y + vpad if anchor_v == "top" else y + h / 2 if anchor_v == "center" else y + h - vpad
    anchor = {"top": "nw", "center": "w", "bottom": "sw"}[anchor_v]
    for text, font, color in pieces:
        if tx > x + w:
            break
        g.cv.create_text(tx, ty, text=text, anchor=anchor, font=font, fill=color, tags="txt")
        tx += font.measure(text)
    return True


# ---------- cell editor ----------

TAG_KEYS = ("bold", "italic", "underline", "strike")


def load_editor(g, e, runs, st):
    """Put runs into the editor with one tag per property."""
    e.delete(0, "end")
    for text, p in runs:
        start = e.index("end")
        e.insert("end", text)
        for k in TAG_KEYS:
            if p.get(k):
                e.tag_add(k, f"1.0+{start}c", f"1.0+{start + len(text)}c")
        for k in ("fg", "size", "font"):
            if p.get(k) is not None:
                e.tag_add(f"{k}={p[k]}", f"1.0+{start}c", f"1.0+{start + len(text)}c")
    e.rich_base = dict(st or {})
    restyle(g, e)


def props_at(e, i):
    p = {}
    for tag in e.tag_names(f"1.0+{i}c"):
        if tag in TAG_KEYS:
            p[tag] = True
        elif "=" in tag:
            k, _, v = tag.partition("=")
            if k in ("fg", "font"):
                p[k] = v
            elif k == "size":
                p[k] = float(v)
    return p


def runs_of_editor(e):
    text = e.get()
    return merge([[ch, props_at(e, i)] for i, ch in enumerate(text)])


def restyle(g, e):
    """Show each character with its own font (Tk tags can't combine fonts, so one tag per combination)."""
    for tag in e.tag_names():
        if tag.startswith("look:"):
            e.tag_delete(tag)
    text = e.get()
    base = getattr(e, "rich_base", {})
    i = 0
    while i < len(text):
        p = props_at(e, i)
        j = i + 1
        while j < len(text) and props_at(e, j) == p:
            j += 1
        if p:
            name = "look:" + repr(sorted(p.items()))
            look = {**base, **p}
            e.tag_configure(name, font=g.font(look), foreground=p.get("fg") or "")
            e.tag_add(name, f"1.0+{i}c", f"1.0+{j}c")
        i = j


def toggle(g, e, key):
    """Ctrl+B / I / U in the editor: format the selected characters (all of them on, or off)."""
    try:
        a, b = e.index("sel.first"), e.index("sel.last")
    except Exception:
        return None
    if a == b:
        return None
    on = all(key in e.tag_names(f"1.0+{i}c") for i in range(a, b))
    (e.tag_remove if on else e.tag_add)(key, f"1.0+{a}c", f"1.0+{b}c")
    e.rich_touched = True
    restyle(g, e)
    return "break"
