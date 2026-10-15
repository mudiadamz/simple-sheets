"""Row / column grouping (outline): Data → Group (Alt+Shift+→) / Ungroup (Alt+Shift+←). A small −/+
button in the header after each group hides / shows its rows or columns. Levels are saved in the file."""
from grid import ALL


def groups(levels):
    """{index: level} -> [(first, last)] of contiguous grouped runs (any level > 0)."""
    out, start, prev = [], None, None
    for i in sorted(k for k, v in levels.items() if v > 0):
        if start is None:
            start = prev = i
        elif i == prev + 1:
            prev = i
        else:
            out.append((start, prev))
            start = prev = i
    if start is not None:
        out.append((start, prev))
    return out


def group_selection(app, on=True):
    """Group (or ungroup) the selected rows - or columns when whole columns are selected."""
    g = app.table
    g.commit()
    r1, c1, r2, c2 = g.sel()
    s = g.sheet
    if r2 >= ALL and c2 < ALL:
        axis, lo, hi, levels = "oc", c1, c2, s.outline_cols
    else:
        if r2 >= ALL:
            r2 = max(len(s.rows) - 1, r1)
        axis, lo, hi, levels = "or", r1, r2, s.outline_rows
    changes = []
    for i in range(lo, hi + 1):
        level = levels.get(i, 0)
        new = min(level + 1, 7) if on else max(level - 1, 0)
        if new != level:
            changes.append((axis, i if axis == "or" else 0, i if axis == "oc" else 0, new))
    if not on:   # ungrouped rows that were collapsed come back
        hidden = s.hidden_rows if axis == "or" else s.hidden_cols
        kind = "hr" if axis == "or" else "hc"
        changes += [(kind, i if axis == "or" else 0, i if axis == "oc" else 0, False)
                    for i in range(lo, hi + 1) if i in hidden and levels.get(i, 0) <= 1]
    if changes:
        app.apply_mixed(changes)
    return len(changes)


def toggle(app, axis, first, last):
    """−/+ button: collapse (hide) the group's rows / columns, or expand them again."""
    s = app.table.sheet
    hidden = s.hidden_rows if axis == "row" else s.hidden_cols
    collapse = not all(i in hidden for i in range(first, last + 1))
    kind = "hr" if axis == "row" else "hc"
    app.apply_mixed([(kind, i if axis == "row" else 0, i if axis == "col" else 0, collapse)
                     for i in range(first, last + 1)])
    return collapse


def draw(g, ys, xs, C):
    """Buttons in the row / column headers just after each group (Excel puts the summary below / right)."""
    g.outline_buttons = []
    if not g.headings():
        return
    s = g.sheet
    size = max(g.px(9), 8)
    for axis, levels, pos in (("row", s.outline_rows, ys), ("col", s.outline_cols, xs)):
        hidden = s.hidden_rows if axis == "row" else s.hidden_cols
        for first, last in groups(levels):
            after = last + 1
            if after not in pos:
                continue
            p, _ = pos[after]
            if axis == "row":
                box = (2, p + 2, 2 + size, p + 2 + size)
            else:
                box = (p + 2, 2, p + 2 + size, 2 + size)
            collapsed = all(i in hidden for i in range(first, last + 1))
            g.cv.create_rectangle(*box, fill=C["cell_bg"], outline=C["header_fg"], tags="outline")
            cx, cy = (box[0] + box[2]) / 2, (box[1] + box[3]) / 2
            g.cv.create_line(box[0] + 2, cy, box[2] - 2, cy, fill=C["header_fg"], tags="outline")
            if collapsed:
                g.cv.create_line(cx, box[1] + 2, cx, box[3] - 2, fill=C["header_fg"], tags="outline")
            g.outline_buttons.append((box, axis, first, last))


def click(g, x, y):
    """True if (x, y) hit a −/+ button (then it toggles the group)."""
    for (x1, y1, x2, y2), axis, first, last in getattr(g, "outline_buttons", []):
        if x1 - 1 <= x <= x2 + 1 and y1 - 1 <= y <= y2 + 1:
            toggle(g.app, axis, first, last)
            return True
    return False
