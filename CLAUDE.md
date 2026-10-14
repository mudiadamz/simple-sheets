# Simple Sheets

Light Tkinter spreadsheet viewer/editor for .xlsx/.xlsm/.xls/.ods/.csv, built into a single Windows exe.

## Rules
- **Always write tests.** Every feature or fix ships with pytest tests in `tests/` in the same change — no throwaway scratch scripts as the only check.
  - Build sample files with `tests/workbooks.py`; use the `app` fixture and `open_in` / `key` helpers from `tests/conftest.py`.
  - Run `python -m pytest` from the project root and report the pass count before calling work done.
- Keep the exe light: stdlib + `openpyxl` + `xlrd` only at runtime (no pandas/numpy in app code).

## Commands
- Run: `python simple_sheets.py [file]`
- Test: `python -m pytest`
- Build: `build.bat` → `dist\SimpleSheets.exe` + `dist\SimpleSheets-Setup.exe` (installer with the app inside). Close any running copy first; it locks the exe. From bash: `cmd //c build.bat`.
- Version: bump `APP_VERSION` in `version.py` (used by the app, installer and Settings > Apps).
- Installer (`installer.py`): per-user, no admin. Installs to `%LOCALAPPDATA%\Programs\SimpleSheets`, Start menu shortcut, "Open with" for .xlsx/.xlsm/.xls/.ods/.csv/.tsv (HKCU registry), uninstaller in Settings > Apps (`Uninstall.exe --uninstall`). Tests use the scratch key `HKCU\Software\SimpleSheetsTest`, never the real Classes keys.
- Icon: `tools/make_icon.py` draws `assets/simple_sheets.ico` (Pillow, build time only).

## Layout
- `simple_sheets.py` — app window, files, sheets add/rename/delete, undo/redo history, save
- `app_menu.py` — ☰ menu (new/open/recent/save) + undo/redo buttons
- `recent.py` — recent files list (`%APPDATA%\SimpleSheets\recent.json`; tests set `SIMPLE_SHEETS_HOME`)
- `settings.py` — user preferences (`settings.json`, same folder), e.g. dark mode
- `theme.py` — light/dark palettes (`theme.C` = active colours); new UI colours go here, never hard-coded
- Tk images / variables always get an explicit master (`master=self`) — without it they attach to another Tk instance in tests
- `sheet_tabs.py` — bottom sheet tabs, + button, rename/delete
- `grid.py` — canvas grid: drawing, selection, keyboard/mouse, editor, clipboard, wrapped text
- `cell_editor.py` — multi-line in-cell editor (Text widget with an Entry-like API)
- `format_bar.py` — formatting toolbar, table styles
- `menus.py` — right-click menu
- `structure.py` — insert/delete rows, columns, cells
- `fill.py` — fill-handle series (numbers, dates, day/month names, "Item 7" text)
- `formula.py` — formula engine
- `formatting.py` — number formats, colors, style dicts
- `sheet_io.py` — load/save (`Sheet` model, patch/rebuild xlsx)
