"""Keyboard everywhere: Ctrl+Alt+5 picks the sheet's charts / pictures one by one (Tab / Shift+Tab go on while
one is picked; arrows move it, Shift+arrows resize, Delete removes, Esc lets go); F6 / Shift+F6 move the
focus between the cells, the formula bar, the name box and the sheet tabs."""


def objects(app):
    s = app.current
    return [("chart", ch) for ch in s.charts] + [("image", im) for im in s.extra.get("images") or []]


def picked(g):
    if getattr(g, "selected_chart", None) is not None:
        return ("chart", g.selected_chart)
    if getattr(g, "selected_image", None) is not None:
        return ("image", g.selected_image)
    return None


def next_object(app, step=1):
    """Pick the next (or previous) chart / picture. Returns ('chart' | 'image', obj) or None."""
    g = app.table
    items = objects(app)
    if not items:
        app.status["text"] = "No charts or pictures on this sheet"
        return None
    now = picked(g)
    i = next((k for k, (kind, o) in enumerate(items) if now and kind == now[0] and o is now[1]), None)
    i = 0 if i is None and step > 0 else (len(items) - 1 if i is None else (i + step) % len(items))
    kind, obj = items[i]
    g.selected_chart = obj if kind == "chart" else None
    g.selected_image = obj if kind == "image" else None
    g.see(*obj["cell"])
    g.cv.focus_set()
    g.redraw()
    app.status["text"] = (f"{'Chart' if kind == 'chart' else 'Picture'} {i + 1} of {len(items)}: arrows move it, "
                          "Shift+arrows resize, Delete removes, Esc lets go")
    return kind, obj


def tab(app, step):
    """Tab in the grid: the next object while one is picked, else the next cell."""
    if picked(app.table) is not None:
        next_object(app, step)
        return "break"
    return None


def panes(app):
    out = [app.table.cv, app.value, app.ref]
    tabs = getattr(app, "tabs", None)
    if tabs is not None:
        out.append(tabs)
    return out


def cycle_focus(app, step=1):
    """F6: cells → formula bar → name box → sheet tabs → cells. Returns the widget focused."""
    order = panes(app)
    now = app.focus_get()
    i = next((k for k, w in enumerate(order) if now is not None and str(now).startswith(str(w))), -1)
    target = order[(i + step) % len(order)]
    target.focus_set()
    if target is getattr(app, "tabs", None):
        app.status["text"] = "Sheet tabs: ← / → change sheet, Enter or Esc goes back to the cells"
    return target


def tabs_key(app, key):
    """Arrow keys / Enter on the focused sheet tabs."""
    if key in ("Left", "Right"):
        app.next_sheet(-1 if key == "Left" else 1)
        app.tabs.focus_set()
    else:
        app.table.cv.focus_set()
    return "break"


def install(app):
    g = app.table
    for seq in ("<Control-Alt-Key-5>", "<Control-Alt-KP_5>"):
        app.bind(seq, lambda e: (next_object(app, 1), "break")[-1])
    g.cv.bind("<Tab>", lambda e: tab(app, 1) or g.move(0, 1))
    g.cv.bind("<Shift-Tab>", lambda e: tab(app, -1) or g.move(0, -1))
    app.bind("<F6>", lambda e: (cycle_focus(app, 1), "break")[-1])
    app.bind("<Shift-F6>", lambda e: (cycle_focus(app, -1), "break")[-1])
    tabs = getattr(app, "tabs", None)
    if tabs is not None:
        tabs.configure(takefocus=True)
        for key in ("Left", "Right", "Return", "Escape"):
            tabs.bind(f"<{key}>", lambda e, k=key: tabs_key(app, k))
