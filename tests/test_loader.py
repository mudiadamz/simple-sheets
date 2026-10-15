"""Big files load on a worker thread with a progress window (and Cancel)."""
import time

import loader
import workbooks


def test_background_load_shows_progress(app, tmp_path, monkeypatch):
    monkeypatch.setattr(loader, "BACKGROUND_BYTES", 0)       # treat every file as big
    path = workbooks.sales(tmp_path / "s.xlsx")
    seen = []
    real = loader.io.load

    def slow(p, progress=None):
        progress(0.5)
        time.sleep(0.2)
        seen.append(app.loading is not None)                  # the progress window is up meanwhile
        return real(p, progress)
    monkeypatch.setattr(loader.io, "load", slow)
    assert app.open_file(str(path))
    assert seen == [True] and app.loading is None
    assert app.path == str(path) and app.table.text(0, 0)


def test_cancel(app, tmp_path, monkeypatch):
    monkeypatch.setattr(loader, "BACKGROUND_BYTES", 0)
    path = workbooks.sales(tmp_path / "s.xlsx")

    def endless(p, progress=None):
        for i in range(500):
            progress(i / 500)
            time.sleep(0.01)
        return {}
    monkeypatch.setattr(loader.io, "load", endless)
    app.after(150, lambda: app.loading.cancel_btn.invoke())
    t = time.time()
    assert app.open_file(str(path)) is False
    assert time.time() - t < 3 and app.path is None and app.status["text"] == "Opening cancelled"


def test_errors_reach_the_user(app, tmp_path, monkeypatch):
    monkeypatch.setattr(loader, "BACKGROUND_BYTES", 0)
    bad = tmp_path / "bad.xlsx"
    bad.write_bytes(b"not a zip")
    assert app.open_file(str(bad)) is False
    assert app.dialogs[-1][:2] == ("showerror", "Error")


def test_progress_reported_by_xlsx_loader(tmp_path):
    import openpyxl
    p = tmp_path / "long.xlsx"
    wb = openpyxl.Workbook()
    for r in range(1, 2001):
        wb.active.append([r])
    wb.save(p)
    got = []
    loader.io.load(str(p), got.append)
    assert got[0] == 0 and 0 < got[-1] <= 1 and got == sorted(got)
