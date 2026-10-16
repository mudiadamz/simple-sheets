# Simple Sheets — user guide

Simple Sheets opens and edits spreadsheet files: Excel (.xlsx, .xlsm, .xlsb, .xls, templates .xltx),
OpenDocument (.ods) and text tables (.csv, .tsv). Saving changes only what you changed, so the
rest of an Excel file (charts, pictures, macros, pivot tables) stays as it was.

## Opening and saving

- ☰ menu → Open, Open recent, or drag a file onto the window.
- Save (Ctrl+S) writes back to the same file. Save as can make .xlsx, .ods, .csv or a template (.xltx).
- Password-protected Excel files open after you type the password; they are saved without one.
- New from template makes an untitled workbook from a template; the template itself never changes.
- If Simple Sheets closes unexpectedly, your unsaved work is offered back the next time it starts.
- If another program changes the file while it's open, you're asked whether to reload it.

## Entering data

- Type in a cell; Enter goes down, Tab goes right. F2 edits the cell. Ctrl+Enter makes a new line.
- Typing the start of a word already in the column suggests the rest; Enter keeps it.
- Drag the small square at the selection's corner to fill a series (1, 2, 3… / Mon, Tue…).
- Drag the selection's edge to move the cells somewhere else.
- Data → Flash fill (Ctrl+E) continues a pattern from one or two examples
  (e.g. first names out of full names).
- Cells with a drop-down list show ▼; Alt+↓ opens the list.

## Formulas

- Start with =. A list of matching functions appears; Tab or Enter takes one, and the arguments
  are shown while you type.
- Click cells, or use the arrow keys right after = ( , or an operator, to put references in.
- About 170 functions, including dynamic arrays (SEQUENCE, FILTER, UNIQUE, SORT…) whose results
  spill into the cells below; A1# means the whole spilled block.
- Excel table references (Sales[Qty], [@Price]) and links to other workbooks ([Book.xlsx]Sheet1!A1) work.
- Tools → Calculation: manual calculation (F9 recalculates) and iterative calculation for circular
  references.

## Formatting

- The toolbar has fonts, sizes, bold / italic / underline, colours, alignment, wrap, borders,
  number formats, table styles and Excel's cell styles (Good, Bad, Heading 1…).
- 🖌 Format painter copies the formatting of the selection to the next cells you select.
- Data → Conditional formatting colours cells by their value, text, duplicates, top / bottom,
  colour scales or data bars. Data → Data validation makes drop-down lists and limits.

## Sheets, rows and columns

- Right-click a sheet tab to rename, duplicate, hide, colour, protect or copy it to a new workbook.
- Right-click row / column headers to insert, delete, hide or resize; Data → Group makes
  collapsible groups (−/+ in the header).
- View: gridlines, headings, formulas, full screen (F11), split window, dark mode.
- Move whole columns or rows: select them by their header, then drag the selected header to a new place
  (a line shows where they'll go). Formats, sizes and formulas move with them; Ctrl+Z undoes it.

## Data tools

- Sort and filter (▼ buttons in the header row), find & replace (Ctrl+F / Ctrl+H), go to (Ctrl+G).
- Data → Text to columns, Remove duplicates, Pivot table (saved as a real Excel pivot table), Goal seek.
- Filters: pick values, a condition (greater than, contains, top 10…) or a fill colour; Data → Advanced filter
  uses a criteria range. Sort can use your own order (Low, Medium, High) from Settings → Fill lists.
- Data → Import text / CSV (encoding, delimiter, column types) and Import table from web.
- Data → Scenario manager, Data table and Solver for what-if questions.
- Tools → Compare with another workbook lists every cell that differs.
- Insert → Chart draws a column, bar, line, pie or area chart from the selection; Insert → Picture adds an image.
- Ctrl+Shift+P (command palette) finds any command by typing part of its name.

## Scripts

- Tools → Script… opens a Python editor. `sheet` is the active sheet, `selection` the selected cells,
  `book` the workbook:
  `sheet["B1"] = sheet["A1"] * 2`, `sheet["C1"] = "=SUM(A1:A9)"`, `sheet.format("A1:C1", bold=True)`,
  `for row in selection.rows(): ...`, `print(...)`. Press F5 to run; Examples… has ready-made scripts.
- A whole run is one undo step. If a script fails or you press Stop, nothing is changed.
- Data → Run command on selection… sends the selected cells as CSV to any program (PowerShell, Node, R,
  another Python…) and puts the CSV it prints back.
- Automation: `SimpleSheets.exe --run job.py book.xlsx --quit` opens the file, runs the script (call
  `book.save()` to keep its changes) and closes; the result is written to `job.log`.
- Scripts can do anything Python can on your computer: only run code you trust. Code inside a workbook
  is never run.

## Pictures in cells: =IMAGE()

- `=IMAGE("https://…/logo.png", "Logo")` shows a web picture inside the cell (PNG, GIF, JPEG, BMP…).
  Optional: sizing 0 = fit keeping its shape (default), 1 = fill the cell, 2 = original size,
  3 = your own height and width in pixels: `=IMAGE(A2, "Photo", 3, 60, 80)`.
- The alt text shows while it downloads, when it can't be loaded, and in other formulas, CSV and PDF.
  Without alt text a picture that can't be loaded shows #picture?; select the cell and the bottom bar says why
  (e.g. the link is a web page - use the picture's own address: right-click it in the browser → Copy image address).
- WebP, JPEG, PNG, GIF, BMP, TIFF and HEIC work; SVG drawings don't.
- Saved so Excel 365 shows it too.

## Google Drive and Google Sheets

- ☰ → Google Drive → Google account…: sign in with Google. The first time, make your own sign-in client
  (free): in the Google Cloud console enable the Google Drive API, add yourself as a test user on the OAuth
  consent screen, create an OAuth client ID of type "Desktop app", and paste its Client ID and secret.
- Sign in opens Google's page in your browser; when it says "Signed in", come back to Simple Sheets.
- Open from Google Drive…: your spreadsheets, newest first (type to search). A Google Sheet opens as .xlsx;
  Save (Ctrl+S) puts it back on Drive. Save to Google Drive… uploads a new copy (as a Google Sheet or as the file).
- The secret and your sign-in are stored encrypted for your Windows account; Sign out removes them.

## Cell scripts (buttons in the sheet)

- Right click a cell → Run a script on click…: write Python (same as Tools → Script; `clicked` is the cell's
  address). The cell gets a small blue corner and a hand pointer; clicking it runs the script (one Ctrl+Z).
- Right click again to Edit / Run / Remove it. Scripts can make buttons too: `sheet.on_click("G2", code)`.
- Saved inside the .xlsx (Excel keeps them, but doesn't run them). A file from somewhere else asks once
  before its scripts run; changed scripts ask again.

## AI assistant

- Tools → AI assistant… (Ctrl+Shift+A) opens a pane on the right of the window (Ctrl+Shift+A or ✕ hides it;
  drag its left edge to make it wider). The blue ✦ AI button at the right of the toolbar does the same.
  Press Enter to ask (Shift+Enter for a new line).
- Add things to a question: Attach… (a picture, PDF, text file or another spreadsheet), Capture (dim the
  screen and drag a box around what to send), or Ctrl+V to paste a copied picture or screenshot (Win+Shift+S).
  They show above the question (✕ removes one) and go with the next message. Pictures and PDFs need a model
  that reads them: in ⚙ each model says what it reads (text, image, file).
- It asks an AI model about your sheet or to change it, through
  OpenRouter with your own API key. Open ⚙ to paste the key (get one at openrouter.ai/keys) and pick a
  model from the list (prices per million tokens are shown; some models are free).
- Ask in plain words: "what's the average of column C?", "add a total row", "make the header bold and
  blue", "split the names in column A into first and last name".
- When cells should change, the AI writes a script. It appears under the chat: read it (you can edit it),
  then press ▶ Run. It's one undo step, and nothing changes if it fails. Nothing runs by itself.
- Auto (the ⚡ icon in the pane's title bar, lit while on; off = Ask): in Ask mode a script shows as a card (Code ▾ to read it) and runs
  when you press ▶ Run; in Auto mode it runs straight away. Either way it's one Ctrl+Z, and a failing
  script changes nothing. ⚙ → Auto model lets OpenRouter pick a model for each question.
- When a script fails, its card says "Script failed" with the error: ↻ Fix it sends the error and the script
  back to the AI, which answers with a corrected script. In Auto mode this happens by itself (up to 2 tries).
- Web (the globe icon in the title bar, or + → Search the web): the AI looks things up online before answering
  ("fill column B with each country's capital", "today's EUR to USD rate in C1") and lists the pages it
  used as links under the answer. OpenRouter charges a little extra for each search; click Web again to turn it off.
- Chats are kept on this computer: ✎ starts a new one, ↺ lists earlier chats to continue; the bin
  on a row deletes that chat, Clear all… deletes every one. The model is chosen in ⚙ (the title's tooltip names it).
- Long chats and big sheets: each question is fitted to the model's size. A sheet that doesn't fit is sent
  from the top as far as it fits (a note says so: send only the selection instead); older parts of a long chat
  are kept as a short summary written by the AI (one extra small request). ✎ starts fresh.
- ⚙ settings also shows your OpenRouter credits left and what this key spent today / this week / this month
  (↻ Refresh), and the requests, tokens and cost counted on this computer (Reset sets them to zero).
- Choose what is sent: the sheet (up to about 3,000 cells) or only the selection. Your question and those
  cells go to OpenRouter and the model's provider. The key is stored encrypted for your Windows account.

## Printing and exporting

- Print (Ctrl+P) and Export to PDF open the Print dialog first: it shows the pages (this sheet, all sheets
  or the selection), then Print… or Export PDF….
- Orientation (Automatic / Portrait / Landscape) and Scaling (Actual size, Fit all columns on one page,
  Fit the sheet on one page) are set there and remembered. Automatic uses the file's page setup
  (margins, print area and repeated title rows always come from the file).
- Export as HTML makes a web page of the current sheet.
- View → Page breaks shows where pages end; Sheet → Insert page break adds your own.
- Files can also be saved as .xlsb (Excel binary).

## Settings

- Tools → Settings: language, how often unsaved work is kept for recovery, how many rows to load,
  interface size, chart colours (colour-blind safe), and an address to check for new versions
  (a new version is downloaded, checked and installed the next time you start).
- View → High contrast for white-on-black; F6 moves between the cells, formula bar, name box and sheet tabs;
  Ctrl+Alt+5 picks charts and pictures with the keyboard.
- Portable copy: unzip SimpleSheets-portable.zip anywhere; settings stay in its SimpleSheetsData folder.
