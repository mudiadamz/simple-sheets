"""View → Split window: a second view of the same sheet under the main one, scrolled on its own.
Clicking a cell there selects it in the main view (where editing happens)."""
from grid import Grid


class SplitGrid(Grid):
    def __init__(self, master, app, scale):
        super().__init__(master, app, scale)
        self.zoom = app.table.zoom

    def sync(self):
        """Follow the main view's sheet and changes."""
        if self.sheet is not self.app.table.sheet:
            view = (self.top, self.left)
            self.sheet = None
            self.load(self.app.table.sheet)
            self.top, self.left = view if self.sheet is getattr(self, "_last", None) else (0, 0)
            self._last = self.sheet
        else:
            self.data_changed()

    # editing belongs to the main view
    def on_press(self, e, extend=False, add=False):
        r, c = self.hit(e.x, e.y)
        if r >= 0 and c >= 0:
            self.app.table.goto(r, c)
        self.app.table.cv.focus_set()
        return "break"

    def on_double(self, e):
        return "break"

    def on_key(self, e):
        return "break"

    def on_right(self, e):
        return "break"

    def edit(self, initial=None):
        return "break"


def toggle(app):
    """Show / hide the second view."""
    g = getattr(app, "split_grid", None)
    if g is not None:
        g.destroy()
        app.split_grid = None
        return False
    g = SplitGrid(app, app, app.scale)
    g.pack(fill="both", expand=True, after=app.table)
    app.split_grid = g
    g.load(app.table.sheet)
    g.top = app.table.top
    app.table.cv.focus_set()
    return True
