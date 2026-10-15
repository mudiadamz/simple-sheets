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

## Printing and exporting

- Print (Ctrl+P) and Export to PDF use the file's page setup: orientation, fit to page, margins,
  print area and repeated title rows.
- Export as HTML makes a web page of the current sheet.
- ☰ → Print preview shows the pages (this sheet, all sheets or the selection); Export all sheets to PDF.
- View → Page breaks shows where pages end; Sheet → Insert page break adds your own.
- Files can also be saved as .xlsb (Excel binary).

## Settings

- Tools → Settings: language, how often unsaved work is kept for recovery, how many rows to load,
  interface size, chart colours (colour-blind safe), and an address to check for new versions
  (a new version is downloaded, checked and installed the next time you start).
- View → High contrast for white-on-black; F6 moves between the cells, formula bar, name box and sheet tabs;
  Ctrl+Alt+5 picks charts and pictures with the keyboard.
- Portable copy: unzip SimpleSheets-portable.zip anywhere; settings stay in its SimpleSheetsData folder.
