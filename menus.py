"""Right-click menu: cut/copy/paste, paste special, insert/delete rows, columns and cells."""
import re
import tkinter as tk
import webbrowser

import theme
from grid import ALL, col_name

ICON = 16


def _icon(master, scale, draw):
    s = max(int(ICON * scale), ICON)
    img = tk.PhotoImage(master=master, width=s, height=s)
    k = s / ICON

    def box(x1, y1, x2, y2, color=None):
        color = color or theme.C["icon"]
        img.put(color, to=(int(x1 * k), int(y1 * k), max(int(x2 * k), int(x1 * k) + 1),
                           max(int(y2 * k), int(y1 * k) + 1)))
    draw(box)
    return img


def make_icons(master, scale):
    def plus(b):
        b(7, 2, 9, 14); b(2, 7, 14, 9)

    def trash(b):
        b(3, 3, 13, 4); b(6, 1, 10, 2)
        b(4, 5, 5, 15); b(11, 5, 12, 15); b(4, 14, 12, 15); b(7, 6, 8, 13); b(9, 6, 10, 13)

    def copy(b):
        b(2, 2, 10, 3); b(2, 2, 3, 11); b(5, 5, 14, 6); b(5, 5, 6, 15); b(13, 5, 14, 15); b(5, 14, 14, 15)

    def paste(b):
        b(3, 3, 13, 4); b(3, 3, 4, 15); b(12, 3, 13, 15); b(3, 14, 13, 15); b(6, 1, 10, 5)

    def cut(b):
        for i in range(8):
            b(4 + i, 2 + i, 5 + i, 3 + i); b(11 - i, 2 + i, 12 - i, 3 + i)
        b(2, 10, 6, 11); b(2, 14, 6, 15); b(2, 10, 3, 15); b(5, 10, 6, 15)
        b(10, 10, 14, 11); b(10, 14, 14, 15); b(10, 10, 11, 15); b(13, 10, 14, 15)

    return {n: _icon(master, scale, f) for n, f in
            (("plus", plus), ("trash", trash), ("copy", copy), ("paste", paste), ("cut", cut))}


URL_TEXT = re.compile(r"(https?://|www\.)\S+|mailto:\S+@\S+", re.I)


def link_at(g, r, c):
    """Address to open for a cell: its hyperlink, or its text when that is a web address."""
    text = g.value(r, c).strip()
    if not text:
        return None          # an emptied cell's old link doesn't count
    link = g.sheet.links.get((r, c))
    if link and not link.startswith("#"):   # "#Sheet!A1" = a place inside the workbook
        return link
    if URL_TEXT.fullmatch(text):
        return text if "://" in text or text.lower().startswith("mailto:") else "http://" + text
    return None


def open_link(url):
    webbrowser.open(url)


def show(app, x_root, y_root):
    g = app.table
    if not g.sheet:
        return
    if not hasattr(app, "menu_icons"):
        app.menu_icons = make_icons(app, app.scale)
    ic = app.menu_icons
    r1, c1, r2, c2 = g.sel()
    whole_cols, whole_rows = r2 >= ALL, c2 >= ALL
    nr = 1 if whole_cols else r2 - r1 + 1
    nc = 1 if whole_rows else c2 - c1 + 1
    rows_txt = "1 row" if nr == 1 else f"{nr} rows"
    cols_txt = "1 column" if nc == 1 else f"{nc} columns"

    m = tk.Menu(app, tearoff=False)
    item = dict(compound="left")
    url = link_at(g, *g.span(*g.cur)[:2])
    if url:
        m.add_command(label="Open link", command=lambda: open_link(url))
    has_link = g.span(*g.cur)[:2] in g.sheet.links
    m.add_command(label="Edit link…" if has_link else "Insert link…", accelerator="Ctrl+K", command=app.edit_link)
    m.add_separator()
    m.add_command(label="Cut", accelerator="Ctrl+X", image=ic["cut"], command=g.cut, **item)
    m.add_command(label="Copy", accelerator="Ctrl+C", image=ic["copy"], command=g.copy, **item)
    m.add_command(label="Paste", accelerator="Ctrl+V", image=ic["paste"], command=g.paste, **item)
    special = tk.Menu(m, tearoff=False)
    special.add_command(label="Values only", accelerator="Ctrl+Shift+V", command=lambda: g.paste("values"))
    special.add_command(label="Format only", command=lambda: g.paste("format"))
    special.add_separator()
    special.add_command(label="Transpose", command=lambda: g.paste(transpose=True))
    special.add_command(label="Skip blanks", command=lambda: g.paste(skip_blanks=True))
    for label, op in (("Add", "add"), ("Subtract", "subtract"), ("Multiply", "multiply"), ("Divide", "divide")):
        special.add_command(label=label, command=lambda op=op: g.paste(op=op))
    m.add_cascade(label="Paste special", menu=special, image=ic["paste"], **item)
    m.add_separator()

    m.add_command(label=f"Insert {rows_txt} above", image=ic["plus"],
                  command=lambda: app.restructure("row", r1, nr), **item)
    m.add_command(label=f"Insert {rows_txt} below", image=ic["plus"],
                  command=lambda: app.restructure("row", r1 + nr, nr), **item)
    m.add_command(label=f"Insert {cols_txt} left", image=ic["plus"],
                  command=lambda: app.restructure("col", c1, nc), **item)
    m.add_command(label=f"Insert {cols_txt} right", image=ic["plus"],
                  command=lambda: app.restructure("col", c1 + nc, nc), **item)
    cells = tk.Menu(m, tearoff=False)
    cells.add_command(label="Shift right", command=lambda: app.shift_cells("right"))
    cells.add_command(label="Shift down", command=lambda: app.shift_cells("down"))
    m.add_cascade(label="Insert cells", menu=cells, image=ic["plus"], **item)
    m.add_separator()

    del_rows = "Delete row" if nr == 1 else f"Delete rows {r1 + 1} - {r2 + 1}"
    del_cols = "Delete column" if nc == 1 else f"Delete columns {col_name(c1)} - {col_name(c2)}"
    m.add_command(label=del_rows, image=ic["trash"], state="disabled" if whole_cols else "normal",
                  command=lambda: app.restructure("row", r1, -nr), **item)
    m.add_command(label=del_cols, image=ic["trash"], state="disabled" if whole_rows else "normal",
                  command=lambda: app.restructure("col", c1, -nc), **item)
    dcells = tk.Menu(m, tearoff=False)
    dcells.add_command(label="Shift left", command=lambda: app.shift_cells("left"))
    dcells.add_command(label="Shift up", command=lambda: app.shift_cells("up"))
    m.add_cascade(label="Delete cells", menu=dcells, image=ic["trash"], **item)
    m.add_separator()

    note_cell = g.span(*g.cur)[:2]
    has_note = note_cell in g.sheet.comments
    m.add_command(label="Edit comment" if has_note else "Insert comment", command=app.edit_comment)
    if has_note:
        m.add_command(label="Delete comment", command=lambda: app.set_comment(*note_cell, ""))
    import cell_scripts  # a script that runs when the cell is clicked
    has_script = bool(cell_scripts.get(app, *note_cell))
    m.add_command(label="Edit cell script…" if has_script else "Run a script on click…",
                  command=lambda: cell_scripts.open_editor(app, *note_cell))
    if has_script:
        m.add_command(label="Run cell script", command=lambda: cell_scripts.run(app, *note_cell))
        m.add_command(label="Remove cell script", command=lambda: cell_scripts.assign(app, *note_cell, None))
    m.add_separator()

    import data_tools
    m.add_command(label="Sort A → Z", command=lambda: data_tools.quick_sort(app, True))
    m.add_command(label="Sort Z → A", command=lambda: data_tools.quick_sort(app, False))
    m.add_command(label="Sort range…", command=lambda: data_tools.SortDialog(app))
    m.add_command(label="Remove filter" if g.sheet.filter else "Create filter",
                  command=lambda: data_tools.toggle_filter(app))
    m.add_separator()

    cr, cc = g.cur
    fz = tk.Menu(m, tearoff=False)
    fz.add_command(label=f"Up to row {cr + 1}" if cr else "Top row", command=lambda: app.set_freeze(max(cr, 1), g.fc))
    fz.add_command(label=f"Up to column {col_name(cc)}" if cc else "First column",
                   command=lambda: app.set_freeze(g.fr, max(cc, 1)))
    fz.add_command(label=f"Up to cell {col_name(cc)}{cr + 1}", state="normal" if cr and cc else "disabled",
                   command=lambda: app.set_freeze(cr, cc))
    fz.add_command(label="Unfreeze", state="normal" if g.fr or g.fc else "disabled",
                   command=lambda: app.set_freeze(0, 0))
    m.add_cascade(label="Freeze", menu=fz)
    m.add_separator()

    multi_cell = g.multi() and not (whole_rows and whole_cols)
    in_merge = any(a <= r2 and b >= r1 and c <= c2 and d >= c1 for a, c, b, d in g.sheet.merges)
    m.add_command(label="Merge cells", state="normal" if multi_cell else "disabled", command=app.merge)
    m.add_command(label="Merge & center", state="normal" if multi_cell else "disabled",
                  command=lambda: app.merge(center=True))
    m.add_command(label="Unmerge cells", state="normal" if in_merge else "disabled", command=app.unmerge)
    m.add_separator()

    hide_rows = "Hide row" if nr == 1 else f"Hide rows {r1 + 1} - {r2 + 1}"
    hide_cols = "Hide column" if nc == 1 else f"Hide columns {col_name(c1)} - {col_name(c2)}"
    any_rows = any(r1 <= i <= r2 for i in g.sheet.hidden_rows)
    any_cols = any(c1 <= i <= c2 for i in g.sheet.hidden_cols)
    m.add_command(label=hide_rows, state="disabled" if whole_cols else "normal",
                  command=lambda: app.hide("row"))
    m.add_command(label=hide_cols, state="disabled" if whole_rows else "normal",
                  command=lambda: app.hide("col"))
    m.add_command(label="Unhide rows", state="normal" if any_rows else "disabled",
                  command=lambda: app.hide("row", False))
    m.add_command(label="Unhide columns", state="normal" if any_cols else "disabled",
                  command=lambda: app.hide("col", False))
    m.add_separator()
    m.add_command(label="Clear contents", accelerator="Delete", command=g.clear)
    import i18n
    i18n.translate_menu(m)
    theme.style_menu(m)
    try:
        m.tk_popup(x_root, y_root)
    finally:
        m.grab_release()
