import os
import sys
import time
import tkinter as tk

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

# Tests never touch the real settings folder (%APPDATA%\SimpleSheets: the user's settings, AI key, recent
# files): every test - with or without the app window - gets a throw-away one, set before any app module loads.
import tempfile  # noqa: E402

SAFE_HOME = tempfile.mkdtemp(prefix="ss-test-home-")
os.environ["SIMPLE_SHEETS_HOME"] = SAFE_HOME
import version  # noqa: E402,F401  (tells simplecore who the app is before any test imports it)

os.environ["SIMPLECORE_SYSTEM_THEME"] = "light"   # the app follows the system's light / dark: same on every machine


@pytest.fixture(autouse=True)
def _never_real_settings(monkeypatch, tmp_path):
    """Each test starts with its own empty settings folder (the app fixture may point it elsewhere again)."""
    monkeypatch.setenv("SIMPLE_SHEETS_HOME", str(tmp_path / "settings-home"))


@pytest.fixture(autouse=True)
def _no_network(monkeypatch):
    """The AI assistant never reaches OpenRouter from a test: its default transport refuses."""
    from simplecore import ai_assist as core_ai

    import ai_assist

    def refuse(*a, **k):
        raise ai_assist.AIError("No network in tests.")
    real = ai_assist.__dict__.setdefault("real_http", core_ai.http)   # for tests of http itself (urlopen faked)
    monkeypatch.setattr(core_ai, "http", refuse)
    monkeypatch.setattr(ai_assist, "http", refuse)
    core_ai.use(ai_assist.SheetsProfile())
    assert real is not refuse
    from simplecore import google_drive  # nor Google: tests give it a fake transport

    def refuse_google(*a, **k):
        raise google_drive.GoogleError("No network in tests.")
    monkeypatch.setattr(google_drive, "transport", refuse_google)
    monkeypatch.setitem(google_drive._access, "token", None)
    monkeypatch.setitem(google_drive._access, "until", 0.0)
    import cell_images  # =IMAGE() pictures: tests give their own bytes

    def refuse_image(url, timeout=15):
        raise OSError("No network in tests.")
    monkeypatch.setattr(cell_images, "fetch", refuse_image)
    cell_images.clear()

if sys.platform == "win32":  # same pixel math as the real app
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass


def make_app():
    """A new App window; Tk sometimes fails to start when many windows open in one process."""
    import simple_sheets
    for attempt in range(5):
        try:
            return simple_sheets.App()
        except simple_sheets.tk.TclError:
            if attempt == 4:
                raise
            time.sleep(0.1)


_shared = {"app": None}   # one window reused by the tests (opening a window per test is most of the time)


def _alive(a):
    try:
        return a is not None and bool(a.winfo_exists())
    except Exception:
        return False


def _reset(a):
    """Back to a just-opened window: one new empty workbook, default view, no dialogs, fresh engine."""
    import inspect

    from simplecore import theme

    import recovery
    from formula import Engine
    for obj in (a, a.table, a.tabs):           # methods a test replaced on the object itself (app.dirty = …)
        for k in list(vars(obj)):
            if inspect.isfunction(getattr(type(obj), k, None)):
                del obj.__dict__[k]
    for w in list(a.winfo_children()):        # dialogs, popups, menus left by the previous test
        if isinstance(w, tk.Toplevel) or isinstance(w, tk.Menu):
            try:
                w.destroy()
            except tk.TclError:
                pass
    g = a.table
    g.commit(cancel=True)
    g.close_dropdown()
    if getattr(a, "split_grid", None) is not None:
        a.split_grid.destroy()
        a.split_grid = None
    if a.attributes("-fullscreen"):
        a.attributes("-fullscreen", False)
    a.full_var.set(False)
    a.show_formulas = False
    a.page_breaks = False
    if a.value.expanded:
        a.toggle_formula_bar()
    pane = a.__dict__.pop("ai_pane", None)
    if pane is not None and pane.win.winfo_exists():
        pane.win.destroy()
    __import__("app_menu").refresh_ai_button(a)
    a.page_breaks_var.set(False)
    a.formulas_var.set(False)
    a.painter = None
    a.formula_home = None
    g.drag = g.fill_preview = g.point = None
    g._wheel_acc, g._pad_acc = 0.0, (0.0, 0.0)   # half-rows of scrolling left over from the previous test
    g.extra, g.ref_marks = [], []
    a.books, a.book_index = [None], 0
    a.book_bar.rebuild()
    a.engine = Engine()
    a.rules = type(a.rules)(a)
    a.history, a.hpos, a.base = [], 0, {}
    a.recovery_id, a.autosaved = recovery.new_id(), None
    a.last_error = None
    a._error_dialog_open = False
    for k in ("unlocked_files", "unlocked_password"):
        a.__dict__.pop(k, None)
    a.follow_var.set(True)
    if a.dark_var.get() or theme.C["dark"] or a.contrast_var.get():
        a.contrast_var.set(False)
        a.set_theme(False)
    if g.zoom != 1.0:
        a.zoom_to(1.0)
    fb = a.fmtbar                              # toolbar "last used" choices back to the defaults
    fb.borders.color = "#000000"
    fb.borders.style.set("thin")
    fb.colors = {"fg": "#000000", "bg": "#ffff00"}
    a.new_document()
    a.status["text"] = ""
    # a click from the previous test must not combine with the next one into a double-click
    g.cv.event_generate("<ButtonPress-2>", x=-50, y=-50, when="tail")    # (button 2 does nothing here)
    g.cv.event_generate("<ButtonRelease-2>", x=-50, y=-50, when="tail")
    a.update()


@pytest.fixture
def app(monkeypatch, tmp_path, request):
    """A real (invisible) app window. Dialogs are stubbed so tests never block;
    the recent-files list lives in the test's temp folder. One window is reused between tests and reset
    each time (mark a test @pytest.mark.fresh_app for a brand-new window)."""
    monkeypatch.setenv("SIMPLE_SHEETS_HOME", str(tmp_path / "home"))
    import simple_sheets
    mb = simple_sheets.messagebox
    dialogs = []   # every stubbed dialog, for assertions / debugging: (kind, title, message)

    def stub(kind, answer):
        return lambda *a, **k: (dialogs.append((kind,) + a), answer)[1]
    monkeypatch.setattr(mb, "askyesnocancel", stub("askyesnocancel", False))
    monkeypatch.setattr(mb, "askokcancel", stub("askokcancel", True))
    for name in ("showerror", "showwarning", "showinfo"):
        monkeypatch.setattr(mb, name, stub(name, None))
    a = _shared["app"]
    fresh = request.node.get_closest_marker("fresh_app") is not None
    if fresh and _alive(a):
        a.destroy()
        a = None
    if _alive(a):
        try:
            _reset(a)
        except Exception:
            try:
                a.destroy()
            except Exception:
                pass
            a = None
    if not _alive(a):
        a = make_app()
        a.opened_state = a.state()   # checked by a test; then use a normal window (steadier focus)
        a.state("normal")
        # invisible AND off-screen: an invisible window still catches real mouse clicks
        a.geometry(f"{int(1000 * a.scale)}x{int(450 * a.scale)}-20000-20000")
        a.attributes("-alpha", 0.0)
        a.update()
    a.dialogs = dialogs
    _shared["app"] = None if fresh else a
    yield a
    if fresh or not _alive(a):
        try:
            a.destroy()
        except Exception:
            pass
        _shared["app"] = None


def pytest_sessionfinish(session, exitstatus):
    a = _shared["app"]
    if _alive(a):
        a.destroy()


def open_in(app, path):
    app.open_file(str(path))
    app.update()
    return app.table


def key(app, sequence, widget=None, **kw):
    """Send a key/mouse event to a widget (default: the grid). Key events need real keyboard focus."""
    w = widget or app.table.cv
    for _ in range(50):   # a just-opened window may take a moment to get focus
        w.focus_force()
        app.update()
        if app.focus_get() is w:
            break
        time.sleep(0.02)
    w.event_generate(sequence, when="now", **kw)
    app.update()
