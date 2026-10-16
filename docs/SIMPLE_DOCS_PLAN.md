# Simple Docs — plan for a Word-like editor built on Simple Sheets

**In short:** about half of what's needed already exists in Simple Sheets: the app shell, files, AI, Google,
PDF, installer and tests. The truly new part is a **document model** and a **page layout engine** (the code
that breaks text into lines and pages). That engine is the core of the project and its biggest risk, so it is
built first and shared by the screen, PDF export and printing.

## 1. Key decisions (recommendation first)

| Decision | Recommended | Alternative | Why |
|---|---|---|---|
| Editor surface | **Draw it ourselves on a canvas**, like `grid.py` | Tk's built-in Text widget | The Text widget has no pages, headers/footers, real tables or columns. A Word-like page view needs our own layout anyway. |
| Layout engine | **One engine** for screen, PDF and printing | Separate code for each | The print-preview bug proved that two layouts drift apart. One engine means the same page breaks everywhere. |
| Measuring text | Width tables per font, cached in Python | Ask Tk for every word | Fast, works off the main thread, and gives identical results on screen and in PDF. |
| .docx files | **Edit in place:** keep the original XML and rewrite only what changed | Rebuild the whole file | The same approach as `patch_xlsx`. Anything we can't edit (SmartArt, content controls…) is written back untouched. |
| Repo | **Same repo**, with shared code moved into a `core/` package | A new repo | Sheets keeps working through small shims. One test suite, one place for CI. |
| Product | A separate `SimpleDocs.exe`; the installer can offer both apps | One combined app | Each exe stays light. They share settings, AI key and Google sign-in. |
| Runtime libraries | Standard library only (+ openpyxl if we embed spreadsheets) | python-docx or lxml | Matches the "keep the exe light" rule. python-docx can still be used in tests to cross-check, like pyxlsb is for .xlsb. |

## 2. What we can reuse from Simple Sheets

| Area | Modules | Changes needed |
|---|---|---|
| App shell | `app_menu`, `recent`, `settings`, `theme`, `i18n`, `palette`, `keyboard_nav`, `book_tabs`, `split_view`, `help_ui`, `settings_ui` | Move to `core/`; small app-specific hooks |
| Safety | `crash`, `recovery`, `diskwatch`, `loader` (big files on a thread) | As they are |
| Windows integration | `dropfiles`, `installer`, `updates`, `platform_ui`, `cloud`, portable mode | Installer: register .docx/.odt/.rtf for "Open with" |
| Spelling | `spell.py` (Windows' own spell checker) | Red squiggles while typing; this matters far more in Docs |
| Files | `office_crypto` (password-protected files use the same OLE + AES scheme), `keep_parts` idea, `pillowless`, `cell_images` (WIC/GDI+ picture decoding) | Point them at .docx parts |
| Output | `pdf_export` (PDF writer), `print_preview`, `html_export` | Add fonts embedded in the PDF (fixes non-Latin text and emoji) |
| AI | `ai_assist`: docked pane, OpenRouter, usage/credits, web search, Fix it, fitting to the model, history | A new script API for documents; actions on the selection |
| Google | `google_drive` (sign-in, Drive) | Export / upload as Google Docs (a different file type in the same Drive calls) |
| Scripting | `scripting` (undo, rollback), `cell_scripts` idea | A `doc` API: paragraphs, styles, tables, find/replace |
| Small UI pieces | `edit_menu`, `format_bar` icons, `borders`, `comments`, `threads`, `find_replace` | Adapt to text instead of cells |
| Build & tests | `build.bat`, `tools/check_exe.py`, conftest (no network, safe settings home), `workbooks.py` pattern | A `documents.py` test-file builder |
| Synergy with Sheets | `formula` engine, `charts` | Formulas in tables (=SUM(ABOVE)), embedded charts, mail merge from a sheet |

## 3. New parts to build

- **Document model (`doc_model.py`)**
  - Section → blocks (paragraph / table / page break) → runs (pieces of text sharing one format) → inline
    objects (picture, field, footnote mark).
  - Styles with inheritance (Normal → Heading 1…), plus list definitions (bullets, numbering, multilevel).
  - Stable IDs on every block, to link back to the XML and to do partial re-layout.
  - Undo uses "kinds" like the Sheets app (`"text"`, `"run"`, `"para"`, `"style"`, `"table"`…), grouped into
    steps; typing merges into one step.
- **Layout engine (`layout.py`)**
  - Lines: break at words, tabs and stops; left / center / right / justify; indents, spacing, line spacing;
    inline pictures sized like large letters.
  - Pages: margins, page size and orientation per section, page breaks, keep-with-next for headings,
    widow/orphan control, headers/footers, page numbers.
  - Tables: column widths (fixed or auto-fit), wrapping inside cells, merged cells, rows split across pages,
    header row repeated on each page.
  - Incremental: after an edit, lay out again only from the changed paragraph until line breaks match the old
    ones; results cached per paragraph.
  - **Speed targets:** typing response under 16 ms in a 100-page document; first layout of 300 pages under
    1 s, on a background thread.
- **Page view (`page_view.py`)**
  - Draws pages on a canvas (only the visible ones), with caret, selection, zoom, and draft / page / outline views.
  - Mouse: click, drag, double-click selects a word, triple-click a paragraph, Shift+click extends.
  - Keys: arrows, Ctrl+arrows by word, Home/End, PageUp/PageDown, Ctrl+Backspace.
  - Typing: input methods for Asian languages (limited in Tk — a risk, see section 6).
  - Clipboard: plain text plus the HTML/RTF formats Windows uses when copying from Word or a browser, so pasted
    formatting survives.
- **.docx (`docx_io.py`):** reads `document.xml`, `styles.xml`, `numbering.xml`, headers/footers, footnotes,
  comments, relationships and media.
  - Writes back with the edit-in-place approach; parts it can't handle are listed in a "can't keep" warning
    (like `extras.lost_parts`).
- **Other formats:** `.odt` read and write (like `ods.py`), `.rtf` read and write, `.txt`, `.md` import/export,
  `.html` export, `.doc` (old binary) as text only, through the OLE reader.
- **PDF upgrade:** embed real TrueType fonts (only the characters used), with Unicode text. This also fixes
  the open gap in Sheets' PDF export.

## 4. Phases (small / medium / large effort)

**Phase 0: foundation (large)**
- Move shared code into `core/`, keeping Sheets' tests green with shims.
- Build the document model, the layout engine for paragraphs only, a page view with caret and typing,
  undo/redo, and open/save `.txt` plus basic `.docx` (text, bold/italic, paragraphs).
- Recovery copies, recent files, dark mode, zoom.
- *Finished when:* you can open a real .docx, edit text, and save it so Word opens it with nothing lost.

**Phase 1: formatting (medium)**
- Character: font, size, bold, italic, underline, strike, colour, highlight, superscript/subscript, clear
  formatting, format painter.
- Paragraph: alignment, indents, spacing, line spacing, borders and shading.
- Styles and a style gallery (Normal, Title, Heading 1–6, Quote, Code), a navigation pane built from the
  headings, bullets and numbering.
- Page setup, page breaks, headers/footers, page numbers.
- *Finished when:* a typical report or letter round-trips with Word, looks the same, and paginates the same.

**Phase 2: content (large)**
- Tables: insert, resize, merge, borders, header row.
- Pictures: inline first, then square text wrapping.
- Links, bookmarks, footnotes/endnotes, table of contents from headings, comments.
- Track changes: show / accept / reject what Word recorded, then record our own edits.
- Find & replace with formatting and regex; word count; spell check while typing; autocorrect; symbols.

**Phase 3: output & interop (medium)**
- PDF with embedded fonts, print preview, .odt / .rtf / .html / .md.
- Password-protected .docx; Google Docs open and save; templates (letter, résumé, report, invoice).
- **Mail merge from an .xlsx or a sheet open in Simple Sheets** — something Word users ask for, and our clear edge.

**Phase 4: AI & automation (small–medium, mostly reuse)**
- AI pane actions on the selection: rewrite, shorten, fix grammar, translate, change tone, summarise. Results
  come in as tracked changes you accept or reject.
- "Write section about…", outline to draft, and Docs versions of web search, Fix it and usage.
- A `doc` script API, cell-script-style buttons (clickable text runs a script), and Docs commands in the
  command palette.

## 5. Tests (following the project rules)

- `tests/documents.py` builds sample .docx files with zipfile and XML templates, the way `workbooks.py` builds
  spreadsheets.
- **Layout tests** with fixed character widths, so line and page breaks are exact and repeatable.
- **Round-trip tests:** open, save without changes, then check untouched XML parts are byte-identical and
  edited ones are valid. Cross-check with python-docx in tests only.
- **Screen vs PDF:** same page count and the same first word on every page.
- **Speed tests:** a 200-page document lays out in under 1 s; typing in it takes under 16 ms per key.
- Existing safety nets carry over: no network in tests, a throw-away settings folder, a check that the exe
  imports every module, and screenshots for UI checks.

## 6. Risks and how to handle them

| Risk | How we handle it |
|---|---|
| Layout engine gets complex: tables, floating pictures, columns | Build it in steps: paragraphs → tables → floating pictures last. Keep it inside one module behind a clear interface. |
| Losing .docx content | Edit in place, keep anything we don't understand, warn before saving if something can't be kept, and test with real Word files. |
| Python speed in large documents | Cached character widths, re-layout only what changed, draw only visible pages, background threads. |
| Right-to-left text (Arabic, Hebrew), Thai shaping, input methods | Tk has limited support. Treat as later work and document the limitation. |
| Screen readers | A Tk canvas isn't accessible (the same limitation discussed for Sheets). A future switch to Qt would mainly change the view; the model, layout and file code stay. |
| Breaking Sheets while moving code to `core/` | Move one module at a time with re-export shims, and require Sheets' full test suite to pass after every move. |

## 7. Open decisions (before Phase 0)

- **Name:** "Simple Docs"?
- **Packaging:** a separate exe (recommended), or one combined "Simple Office" app?
- **Must-haves for the first release:** for example, tables and track changes in Phase 1 instead of Phase 2?
- **Shared code:** move it to `core/` now (recommended), or copy what's needed first and unify later?
