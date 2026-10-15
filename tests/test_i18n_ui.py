"""Whole-UI translation: dialogs (widget tree), tooltips, message boxes, own language files."""
import json
import tkinter as tk
from tkinter import messagebox, ttk

import pytest

import i18n


@pytest.fixture
def indonesian():
    i18n.set_lang("id")
    yield
    i18n.unwrap_dialogs()
    i18n.set_lang("en")


def test_translate_widget_tree(app, indonesian):
    top = tk.Toplevel(app)
    top.title("Solver")
    f = ttk.Frame(top)
    f.pack()
    b = ttk.Button(f, text="Cancel")
    lab = ttk.Label(f, text="Set objective")
    keep = ttk.Label(f, text="B5")
    tree = ttk.Treeview(f, columns=("a",), show="headings")
    tree.heading("a", text="Sheet")
    for w in (b, lab, keep, tree):
        w.pack()
    assert i18n.translate_widgets(top) == 4
    assert b["text"] == "Batal" and lab["text"] == "Sel tujuan" and keep["text"] == "B5"
    assert tree.heading("a", "text") == "Lembar" and top.title() == "Pemecah"
    top.destroy()


def test_new_windows_translated_when_shown(app, indonesian):
    assert i18n.install(app) is True
    import whatif
    d = whatif.scenario_manager(app)
    d.update()
    texts = [w["text"] for w in d.winfo_children()[0].winfo_children() if w.winfo_class() == "TButton"]
    assert "Tutup" in texts and d.title() == "Pengelola skenario"
    d.destroy()


def test_message_boxes(app, monkeypatch, indonesian):     # monkeypatch first: undone last
    seen = []
    monkeypatch.setattr(messagebox, "showinfo", lambda title=None, message=None, **k: seen.append((title, message)))
    i18n.wrap_dialogs()
    messagebox.showinfo("Solver", "No solution found.\nB5 = 3", parent=app)
    assert seen == [("Pemecah", "Tidak ada solusi.\nB5 = 3")]


def test_english_untouched(app):
    i18n.set_lang("en")
    assert i18n.install(app) is False
    b = ttk.Button(app, text="Cancel")
    assert i18n.translate_widgets(b) == 0 and b["text"] == "Cancel"
    b.destroy()


def test_tooltip_translated(app, indonesian):
    from format_bar import Tooltip
    b = ttk.Button(app, text="x")
    b.pack()
    app.update()
    t = Tooltip(b, "Menu")
    t.text = "Undo (Ctrl+Z)"
    t.show()
    assert t.tip.winfo_children()[0]["text"] == "Batalkan (Ctrl+Z)"
    t.hide()
    b.destroy()


def test_own_language_file(tmp_path, monkeypatch):
    monkeypatch.setenv("SIMPLE_SHEETS_HOME", str(tmp_path))
    (tmp_path / "lang_de.json").write_text(json.dumps({"_name": "Deutsch", "Save": "Speichern"}), encoding="utf-8")
    i18n._merge()
    try:
        assert i18n.LANGS["de"] == "Deutsch"
        i18n.set_lang("de")
        assert i18n.tr("Save") == "Speichern" and i18n.tr("Open…") == "Open…"
    finally:
        i18n.set_lang("en")
        i18n.LANGS.pop("de", None)
        i18n.T.pop("de", None)


def test_new_menu_items_have_translations():
    import i18n_strings
    for code in ("id", "es"):
        assert set(i18n_strings.MENUS[code]) == set(i18n_strings.MENUS["id"])
        assert set(i18n_strings.UI[code]) == set(i18n_strings.UI["id"])
