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
- Build: `build.bat` → `dist\SimpleSheets.exe` + `dist\SimpleSheets-Setup.exe` (installer with the app inside) + `.sha256` + `SimpleSheets-portable.zip` (`tools/package.py`). Close any running copy first; it locks the exe. From bash: `cmd //c build.bat`.
- Version: bump `APP_VERSION` in `version.py` (used by the app, installer and Settings > Apps).
- Installer (`installer.py`): per-user, no admin. Installs to `%LOCALAPPDATA%\Programs\SimpleSheets`, Start menu shortcut, "Open with" for .xlsx/.xlsm/.xlsb/.xls/.ods/.csv/.tsv (HKCU registry), uninstaller in Settings > Apps (`Uninstall.exe --uninstall`). Silent: `Setup.exe /S [/D=folder] [/desktop] [/noopenwith] [/relaunch]` (log in `%TEMP%\SimpleSheets-setup.log`). Tests use the scratch key `HKCU\Software\SimpleSheetsTest`, never the real Classes keys.
- Icon: `tools/make_icon.py` draws `assets/simple_sheets.ico` (Pillow, build time only).

## Layout
- `simple_sheets.py` — app window, files, sheets add/rename/delete, undo/redo history, save
- `app_menu.py` — ☰ menu (new/open/recent/save) + undo/redo buttons
- `recent.py` — recent files list (`%APPDATA%\SimpleSheets\recent.json`; tests set `SIMPLE_SHEETS_HOME`)
  - conftest points `SIMPLE_SHEETS_HOME` at a temp folder for EVERY test (autouse): never write the user's real
    settings (a test once wiped the saved AI key); `tests/test_settings_safety.py` guards it
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
- `functions_extra.py` — more functions (finance, statistics, trig, dates, ROW/OFFSET/INDIRECT…), registered into `formula.FUNCS`
- `autocomplete.py` — function / name list and argument hints while typing a formula
- `formatting.py` — number formats, colors, style dicts
- `borders.py` — border picker
- `comments.py` — comment / link dialogs, hover note
- `data_tools.py` — sort and filter
- `find_replace.py` — find & replace dialog
- `name_manager.py` — named ranges dialog
- `rules.py` — data validation (drop-down, reject input) and conditional formatting from the file
- `pdf_export.py` — own PDF writer: export / print a sheet
- `recovery.py` — crash-recovery copies (`<config>\recovery`), offered back after a crash
- `diskwatch.py` — file changed on disk / locked by Excel
- `dropfiles.py` — drag & drop files from Explorer (WM_DROPFILES)
- `loader.py` — big files load on a thread with a progress window
- `ods.py` — .ods formulas and styles; `sheet_io.py` — load/save (`Sheet` model, patch/rebuild xlsx)
- `installer.py` — the setup / uninstall program (removes an older installed version first)
- `functions_array.py` — dynamic-array functions (results spill: `Engine.ensure_spill`)
- `sheet_props.py` — sheet settings (gridlines, headings, tab colour, protection, page setup) and workbook properties
- `workbook_tools.py` — view toggles, tab colours, duplicate / copy sheet, protection, properties, calculation dialogs
- `extras.py` — tables / images / charts read from the file; parts a save would drop (shapes…) for the warning
- `tables.py` — Excel tables (banding, ▼ filter, growing); `charts.py`, `images.py` — drawn over the grid
- `rules_dialog.py` — create conditional formats / validation; `outline.py` — row / column groups
- `data_more.py` — text to columns, remove duplicates, flash fill, pivot, goal seek
- `richtext.py` — several fonts in one cell; `externals.py` — links to other workbooks
- `office_crypto.py` — open password-protected files (OLE reader + AES); `xlsb.py` — .xlsb reader;
  `xls_formulas.py` — .xls formula decoder; `ods_write.py` — .ods writer; `templates.py` — new from template
- `book_tabs.py` — several workbooks in one window; `split_view.py` — split window
- `spell.py` (Windows spell checker via ctypes), `html_export.py`, `help_ui.py` (+ `assets/USER_GUIDE.md`),
  `settings_ui.py` (also interface size, chart palette), `i18n.py` (menus, dialogs, tooltips, messages;
  `i18n_strings.py` texts; own `lang_<code>.json` in the settings folder)
- `updates.py` — update feed; downloads the setup, checks its SHA-256, runs it silently (`/S /relaunch`) at the next start
- `keep_parts.py` — shapes / text boxes / sparklines copied through on save; `pillowless.py` — pictures kept without Pillow
- `crash.py` — error dialog + `errors.log`; `threads.py` — threaded comments; `pivots.py` — real Excel pivot parts
- `function_help.py` — Insert function / Evaluate formula; `functions_lambda.py` — LET, LAMBDA, MAP, XMATCH…
- `import_data.py` — CSV wizard, web tables; `whatif.py` — scenarios, data tables, solver; `compare.py` — compare workbooks
- `xlsb_formula.py` / `xlsb_write.py` — .xlsb formulas both ways, save as .xlsb (cross-checked in tests with pyxlsb)
- `print_preview.py` — preview window, PDF of all sheets / selection, page breaks (view, manual, saved)
- `cloud.py` — OneDrive / Dropbox / Google Drive folders, conflict copies; `platform_ui.py` — Linux / macOS differences
- `keyboard_nav.py` — Ctrl+Alt+5 charts / pictures, F6 panes, sheet tabs by keyboard; `palette.py` — Ctrl+Shift+P, tips
- `scripting.py` — Tools → Script… (Python with a `sheet` / `book` / `selection` API, one undo step, rollback on error),
  Data → Run command on selection (CSV through any program), `--run script.py [file] [--quit]`
- `ai_assist.py` — Tools → AI assistant (OpenRouter, user's key stored with DPAPI, model picker); the model replies
  with a script for `scripting.run`, shown first and run only on ▶ Run. Tests use a fake transport, never the network
- `structure.move` / `App.move_lines` — drag selected column / row headers to move them (`formula.permute` fixes refs)
- Portable mode: a `SimpleSheetsData` folder (or `portable.txt`) next to the exe holds the settings (`recent.portable_dir`)
- Don't write Python containing backslashes or non-ASCII through bash heredocs (they get mangled): use Edit / Write
- Saving: per-sheet things that aren't cell edits are undoable kinds in `App.read/write_cell`
  ("sp", "wp", "cf", "dv", "ch", "tb", "rt", "or"/"oc") and are written by `patch_xlsx(..., extra=App.save_extras())`
- Lint: `python -m ruff check .` (config `ruff.toml`); Linux/macOS build: `build.sh` (→ `build_linux.sh` / `build_macos.sh`,
  untested here); signing: set SIGN_CERT / SIGN_PASS
- CI: `.github/workflows/ci.yml` (tests + build on Windows); feature checklist in `ROADMAP.md`
