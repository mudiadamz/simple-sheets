# Simple Sheets – feature checklist & roadmap

What a spreadsheet app covers (grouped like the parts of the OOXML / OpenDocument spreadsheet formats),
what Simple Sheets already does, and what is open for future work.

- `[x]` done  ·  `[~]` partly done (see note)  ·  `[ ]` not done yet
- **Kept** = not shown or editable in the app, but left intact in the file when saving.
- Priority for open items: **P1** data safety / correctness, **P2** high-value feature, **P3** nice to have.
- One unchecked item ≈ one pull request. Every PR ships with pytest tests (see `CLAUDE.md`).

> Preservation facts in this file were checked by saving a test workbook through the app
> (normal save, and save after inserting a row, which rebuilds the sheet).

---

## 1. File formats

- [x] `.xlsx` read / write (only changed cells are written; the rest of the file is kept)
- [x] `.xlsm` read / write, macros (VBA) kept
- [x] `.csv` / `.tsv` read / write (delimiter and encoding kept; formulas saved as results)
- [x] `.xls` (Excel 97–2003) – values, formulas, formatting, widths, heights, merges, hidden parts, frozen panes; save goes to `.xlsx`
  - [x] **P1** dates, booleans and errors read by cell type (dates used to show as day numbers)
  - [x] **P2** read formatting (xlrd `formatting_info=True`)
  - [x] **P3** read formulas: own BIFF8 formula decoder; a formula is kept only when it reproduces Excel's saved result
- [x] `.ods` (OpenDocument) – read and write values, formulas and formatting
  - [x] **P2** read formulas (OpenFormula → Excel) + formatting, number formats, widths, heights, merges, hidden parts
  - [x] **P3** write `.ods` (values, formulas, number formats, fonts, colours, borders, merges, sizes)
- [~] **P3** `.xlsb` (binary Excel): read values, dates and hidden sheets (formulas as results, no formatting); save goes to `.xlsx`
- [x] **P3** templates (`.xltx` / `.xltm`): open, save as template, ☰ → New from template
- [x] **P3** password-protected / encrypted files: open with the password (Agile + Standard encryption, own AES); saved without a password (Save as)
- [~] large files – over 15 MB: values + formats only (no widths / merges / hidden); rows to load is a setting (50,000), more on request
  - [x] **P2** load in the background with a progress bar (and Cancel)
  - [x] **P3** rows to load is a setting; click the status bar to load every row of a long sheet

## 2. Workbook & sheets

- [x] several sheets, tabs at the bottom
- [x] add, rename (inline), delete, reorder (drag), hide / unhide
- [x] "very hidden" sheets stay very hidden
- [x] formulas follow sheet renames; deleted sheet → `#REF!`
- [x] defined names (named ranges) – used in formulas, follow inserted/deleted rows and sheet renames, saved
  - [x] **P2** evaluate named ranges in formulas; name manager (☰ → Named ranges…)
- [x] sheet tab colours – shown and set
  - [x] **P3** show and set tab colours
- [x] **P3** duplicate sheet; copy a sheet to a new workbook
- [x] document properties (title, author…) – shown and edited
  - [x] **P3** edit properties (Tools → Properties)
- [x] **P3** external links to other workbooks: the other file is read when found, else Excel's stored values; Data → Update links

## 3. Cells & data

- [x] text, numbers, dates (ISO), booleans, errors
- [x] line breaks in a cell (Ctrl+Enter / Alt+Enter)
- [x] hyperlinks: kept, emptied cells drop their link, **Open link** on right-click (also for plain web addresses)
  - [x] **P2** insert / edit / remove a link (Ctrl+K)
- [x] rich text (several fonts inside one cell)
  - [x] **P1** rich text kept on save (also when the sheet is rebuilt); editing such a cell makes it plain
  - [x] **P3** show and edit rich text (Ctrl+B / I / U on selected characters in the cell)
- [x] comments / notes – shown, edited, saved; they move with inserted/deleted rows and columns
  - [x] **P2** show (red corner + hover) and add / edit / delete
- [x] data validation (drop-down lists, number limits) – used, created, kept
  - [x] **P1** ranges (and their formulas) move with inserted/deleted rows and columns
  - [x] **P2** show the drop-down arrow and list (Alt+↓), reject invalid input (stop / warning / information)
  - [x] **P3** create / remove validation rules (Data → Data validation)
- [x] images – shown and kept
  - [x] **P3** show images on the grid (PNG / GIF; other formats as a labelled box)
- [~] **P3** shapes, text boxes, form controls, slicers, sparklines: not kept by openpyxl - the app now warns before a save would drop them
- [x] **P3** "Text to columns", remove duplicates, flash fill (Ctrl+E)

## 4. Formulas

- [x] formula engine, ~170 functions (math, logic, text, dates, lookup, conditional aggregates)
- [x] cross-sheet references, quoted sheet names
- [x] click / shift-click / drag to insert references (also across sheets), coloured reference outlines
- [x] relative / absolute refs move on paste, fill and insert / delete
- [x] circular reference detection (`#CIRC!`), long dependency chains
- [x] unknown function → Excel's last saved result is shown
- [x] **P1** calculated results are written into saved files, so viewers that don't calculate show values
- [x] **P2** more functions: financial (PMT, FV, PV, NPER, RATE, NPV, IRR), statistical (STDEV, VAR, MODE, PERCENTILE, QUARTILE, RANK, COUNTUNIQUE), math / trig, EDATE, EOMONTH, NETWORKDAYS, WORKDAY, DATEDIF, ROW, COLUMN, ADDRESS, INDIRECT, OFFSET, SUMPRODUCT…
  - [x] **P3** SEQUENCE, FILTER, UNIQUE, SORT, SORTBY, TRANSPOSE, VSTACK, HSTACK, TAKE, DROP, TEXTSPLIT, MMULT…
- [x] **P2** formula autocomplete (function and range names, argument hints; in the cell and the formula bar)
- [x] **P3** dynamic arrays: results spill (#SPILL! when blocked), A1# spill references, {1,2;3,4} constants, element-wise math on ranges
- [x] **P3** structured references to Excel tables (`Table1[Col]`, `[#All]`, `[@Col]`…)
- [x] **P3** iterative calculation, manual calculation mode (F9), saved in the file
- [x] **P3** arrow keys (Shift for a range) point at cells while typing a formula

## 5. Formatting

- [x] bold, italic, underline, font size, text style presets
- [~] strikethrough – read / saved, no toolbar button
- [x] font family – shown and editable
  - [x] **P2** font family picker and display
- [x] text colour, fill colour (theme colours + tints read correctly)
- [~] gradient / pattern fills – kept, shown as plain
- [x] horizontal align (dropdown), vertical align (dropdown), wrap text, merge / unmerge / merge & center
- [x] borders: 10 placements, 6 line styles, colour (diagonals kept, not editable)
- [x] number formats displayed like Excel (decimals, thousands, %, currency, dates, times, negatives)
  - [x] **P2** number format picker (General, Number, Currency, Percent, Date, Custom…), more / fewer decimals
- [~] indent, text rotation, shrink-to-fit – kept, not shown
- [x] table styles (applied as cell formatting); real Excel tables in a file are kept and resized with inserted/deleted rows/columns
  - [x] **P3** real Excel tables: style banding, ▼ filter buttons (filter saved in the table), the table grows when typing below it
- [x] conditional formatting – shown, created, kept
  - [x] **P1** ranges (and their formulas) move with inserted/deleted rows and columns
  - [x] **P2** show colour rules (cell value, formula, text, blanks, errors, duplicates, top/bottom, above/below average, data bars, colour scales)
  - [x] **P3** create / delete rules (Data → Conditional formatting)
- [x] clear formatting, paste format only
- [x] **P3** Excel's cell styles (Good, Bad, Neutral, Heading 1-4, Title, Total, Note…) - applied as formatting
- [x] **P3** format painter (🖌)

## 6. Rows & columns

- [x] insert / delete rows and columns (formulas, merges, hidden state, widths and links follow)
- [x] insert / delete cells (shift right / down / left / up)
- [x] hide / unhide rows and columns (header markers, arrows skip them)
- [x] column width: drag, double-click auto-fit, saved
- [x] row height – grows automatically for big fonts and wrapped text; drag to set, double-click to auto-fit, saved
  - [x] **P2** drag to resize rows, double-click to auto-fit
- [x] freeze panes – shown, set from the menu, saved
  - [x] **P2** show frozen rows/columns and set them
  - [x] **P1** frozen position moves with inserted/deleted rows and columns
- [x] **P3** row / column grouping (outline +/−), saved

## 7. Editing

- [x] undo / redo for everything (values, formats, structure, sheets, widths)
- [x] cut / copy / paste with Excel's clipboard; paste values only / format only
- [x] fill handle (numbers, dates, weekdays, months, "Item 7"), Ctrl+D / Ctrl+R
- [x] smart Ctrl+A, Ctrl+arrows, Shift selection, row/column header selection
- [x] in-cell editor + formula bar; the edit survives switching to another program
- [x] **P2** find & replace (Ctrl+F / Ctrl+H), with match case / whole cell / in formulas
- [x] **P2** go to (Ctrl+G / name box: type `B12`, a range or a name)
- [x] **P2** multiple selections (Ctrl+click several ranges: format, clear, status sum)
- [x] **P3** autocomplete from values already in the column
- [x] **P3** drag a selection by its edge to move it
- [x] **P3** spell check (F7, the Windows spell checker)

## 8. Data tools

- [x] **P2** sort (A→Z, Z→A, sort range dialog with header row; keeps rows together)
- [x] autofilter – header buttons, filter by value / condition, saved
  - [x] **P2** filter buttons in the header row, filter by value / condition
- [~] **P3** pivot tables: kept on save by openpyxl (not checked with a real pivot file here); Data → Pivot table makes a summary sheet
- [x] **P3** charts: shown over the grid, Insert → Chart, right-click → Delete chart, saved
- [x] **P3** what-if: goal seek

## 9. View

- [x] zoom (Ctrl +/−, Ctrl+wheel, 50–300%)
- [x] dark mode (remembered)
- [x] status bar: count / sum / average of the selection
- [x] formula reference highlighting while typing
- [x] **P2** freeze panes view (see 6)
- [x] **P3** hide / show gridlines and headings (saved)
- [x] **P3** split window, full screen (F11)
- [x] **P3** show formulas instead of results (View → Formulas, Ctrl+`)

## 10. Protection & security

- [~] sheet protection – enforced, protect / unprotect; workbook-structure protection kept, not enforced
  - [x] **P3** locked cells respected on protected sheets; protect / unprotect with password; lock / unlock cells
- [x] macros kept in `.xlsm` (never run)
- [x] **P3** open password-protected files (see 1)

## 11. Print & export

- [x] **P2** print (Ctrl+P: through a PDF and the PDF program)
  - [x] **P3** the file's page setup is used: orientation, fit to pages, scale, margins, print area, repeated title rows / columns, gridlines
- [x] **P2** export to PDF (own writer: fills, borders, fonts, alignment, wrap, pages down and across)
- [x] **P3** export current sheet as HTML

## 12. App & platform

- [x] ☰ menu, Save button with unsaved state, recent files, new window
- [x] remembers dark mode; recent files in `%APPDATA%\SimpleSheets`
- [x] installer (per user, no admin): Start menu, "Open with" for .xlsx/.xlsm/.xls/.ods/.csv/.tsv, uninstaller
- [x] app icon
- [x] **P2** crash recovery: a copy of unsaved work every minute, offered back after a crash
- [x] **P2** drag & drop a file onto the window to open it
- [x] **P2** warn when the file changed on disk (offer reload; ask before overwriting), show when it's locked by Excel
- [x] **P3** several workbooks in one window (tabs): ☰ → New tab (Ctrl+T) / Open in new tab; Settings → open files in tabs
- [x] **P3** settings dialog (language, recovery interval, rows to load, tabs, update feed)
- [~] **P3** menus in English / Bahasa Indonesia / Español; keyboard: F10 menu, Shift+F10 context menu, Ctrl+PgUp/PgDn sheets. Screen readers: not possible with Tk
- [~] **P3** code signing in build.bat / CI when a certificate is set (none here yet); update check against a release feed (no automatic install)
- [~] **P3** macOS / Linux builds (`build.sh`, CI jobs) - not tried on those systems yet

## 13. Project

- [x] pytest suite (`python -m pytest`), GUI tests run off-screen
- [x] **P2** CI (GitHub Actions on Windows: tests + build artifacts, `.github/workflows/ci.yml`)
- [x] **P3** lint check (ruff, `ruff.toml`), in CI
- [x] **P3** user guide (Help → User guide, `assets/USER_GUIDE.md`) and keyboard shortcuts (F1)

---

## Roadmap 2 (from version 1.2)

Version 1.1.0 finished the first checklist. The next stage is about making what exists solid, fast and pleasant,
then filling the remaining gaps. Same rules: one item ≈ one pull request, every PR ships with tests.
**R1** = do next (quality / trust), **R2** = high value, **R3** = later.

### A. Trust & data safety (R1)

- [x] **R1** round-trip test corpus: real files from Excel, LibreOffice, Google Sheets (pivots, charts, macros,
      `.xlsb`, old `.xls`) opened and saved in CI, compared part by part with the original
- [x] **R1** keep shapes, text boxes and form controls on save (copy their drawing XML through instead of losing it)
- [x] **R1** save password-protected files with their password (AES encryption + OLE writer; reading is done)
- [x] **R1** sparklines and slicers kept (copy the `x14` extension parts through) — sparklines kept; **slicers are not** (still warned before saving: their cache parts need Excel to verify)
- [x] **R1** formulas that point at moved cells: moving a block (drag) updates references to it, like Excel
- [x] **R1** crash report: an unexpected error shows a dialog with "copy details" instead of failing quietly

### B. Speed (R1)

- [x] **R1** startup under 0.3 s: build the menus and icons lazily (test windows now take ~0.75 s to open)
- [x] **R1** dependency tracking in the formula engine: recalculate only the cells an edit affects
      (today every edit clears all results)
- [x] **R1** draw only what changed (scrolling and typing redraw the whole visible grid)
- [x] **R2** large files: read rows on demand while scrolling instead of loading a fixed number
- [x] **R2** test suite back under 1 minute (share one app window between read-only tests)

### C. Everyday editing (R2)

- [x] **R2** edit and reorder existing conditional formatting / validation rules (today: add and delete)
- [x] **R2** chart editing: change type, title, data range; move and resize charts with the mouse
- [x] **R2** insert pictures (and move / resize / delete them)
- [x] **R2** find & replace across all sheets, with "find all" results list
- [x] **R2** paste special: transpose, add / multiply, skip blanks
- [x] **R2** fill handle for formulas across sheets, custom lists (Mon–Sun in other languages)
- [x] **R2** cell comments as threads (replies), notes vs comments like Excel 365
- [x] **R2** strikethrough, indent, text rotation buttons; gradient fills shown
- [x] **R2** formula bar: grows for long formulas, colours the references like the cell editor
- [x] **R3** "Insert function" dialog with descriptions and examples for every function

### D. Data & analysis (R2)

- [x] **R2** real pivot tables: build, refresh and save them as Excel pivot tables (not just a summary sheet) — checked with openpyxl, not yet opened in real Excel
- [x] **R2** advanced filter and custom sort lists; filter by colour
- [x] **R2** data import: CSV wizard (encoding, delimiter, column types preview), import from a web table
- [x] **R3** what-if: scenario manager, data tables; solver for simple linear problems
- [x] **R3** LAMBDA / LET / XMATCH / more 365 functions; formula evaluation step by step

### E. Files & sharing (R2)

- [x] **R2** `.xlsb` formatting and formulas; save as `.xlsb` — `xlsb.py`, `xlsb_formula.py`, `xlsb_write.py`; checked with pyxlsb, not yet opened in real Excel
- [x] **R2** export: PDF of all sheets / the selection, print preview window, page breaks view
- [x] **R3** open from / save to OneDrive or Google Drive folders with conflict detection
- [x] **R3** compare two workbooks (differences by sheet and cell)

### F. Platform & delivery (R2)

- [ ] **R2** code-signing certificate in CI; signed installer — wired in `build.bat` (`SIGN_CERT` / `SIGN_PASS`); needs a bought certificate
- [x] **R2** automatic updates: download, verify the signature, install on next start — SHA-256 checked (the signature check waits for the certificate)
- [x] **R2** try and fix the Linux / macOS builds; a `.deb` / AppImage and a macOS `.dmg` — `build_linux.sh` (.deb), `build_macos.sh` (.app + .dmg), `platform_ui.py`; not run on those systems yet
- [x] **R3** portable mode (settings next to the exe), silent install for companies (`/S`)
- [x] **R3** more languages for the whole UI (dialogs, tooltips, messages — not only menus)

### G. Accessibility & UX (R3)

- [x] **R3** high-contrast theme, bigger UI scale setting, colour-blind safe chart palette
- [x] **R3** every action reachable by keyboard (dialogs, chart and picture handling)
- [x] **R3** screen-reader support (needs a different UI toolkit or a native accessibility layer — study first) — study done: `docs/accessibility-study.md` (next: an announce channel)
- [x] **R3** first-run tips; command palette (Ctrl+Shift+P) to find any action by name

Done in version 2.0.0 (all of the above except the code-signing certificate and keeping slicers).

### Suggested order

1. **1.2 "Solid":** A (round-trip corpus, shapes kept, encrypted save) + B (startup, dependency tracking).
2. **1.3 "Comfortable":** C (rule / chart editing, pictures, find in all sheets, paste special).
3. **1.4 "Analyst":** D (real pivot tables, import wizard) + E (`.xlsb`, print preview).
4. **2.0 "Everywhere":** F (signed, auto-updating, Linux / macOS packages) + G.
