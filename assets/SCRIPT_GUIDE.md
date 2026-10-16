# Script guide

Scripts let you change many cells at once with a few lines of Python, or hand the selected cells to any
other program. Open the script window with Tools → Script…, type a script and press F5 (or ▶ Run).

- A whole run is ONE undo step: Ctrl+Z puts everything back.
- If the script stops with an error (or you press ■ Stop), nothing is changed. The error says which line.
- print(...) shows text in the output pane under the script.
- Examples… (top right of the script window) has ready-made scripts to start from.
- Your last script is kept for next time. Open… / Save… keep scripts as .py files.

# Your first script

Select a few numbers, then run:

```
for i, row in enumerate(selection.rows()):
    for j, v in enumerate(row):
        if isinstance(v, (int, float)):
            sheet.cell(selection.row + i, selection.col + j, v * 2)
```

Every number in the selection is doubled. Ctrl+Z brings the old numbers back.

# The three names a script gets

sheet       the sheet you're looking at
selection   the selected cells
book        the whole workbook (all sheets)

# Reading cells

```
sheet["A1"]            # one cell: a number, text, True/False, or None when it's empty
sheet["A1:C3"]         # a block: a list of rows, e.g. [[1, 2, 3], [4, 5, 6], [7, 8, 9]]
sheet["B:B"]           # a whole column (down to the last row with data)
sheet.cell(2, 3)       # row 2, column 3 (= C2), counting from 1 like Excel
sheet.formula("D2")    # what was typed, e.g. "=SUM(A2:C2)" (sheet["D2"] gives its result)
sheet.rows()           # every row of the sheet as a list of values
sheet.last_row         # number of the last row with data
sheet.last_col         # number of the last column with data
```

Formulas come back calculated. Numbers come back as numbers (2.5, 42), ready for maths.

# Writing cells

```
sheet["A1"] = "Total"                 # text
sheet["B1"] = 42                      # a number
sheet["C1"] = "=SUM(A2:A100)"         # a formula: text starting with =
sheet["D1"] = None                    # empty the cell
sheet["A3"] = [[1, 2, 3], [4, 5, 6]]  # rows of values, spread from A3 to the right and down
sheet["E1:E10"] = 0                   # one value fills the whole block
sheet.cell(5, 2, "hello")             # row 5, column 2 (= B5)
```

A formula a script writes is calculated straight away, so the script can read its result on the next line.

# Formatting

```
sheet.format("A1:D1", bold=True, bg="#ffff00")
sheet.format("B2:B50", numfmt="#,##0.00", align="right")
sheet.format("A1", size=14, fg="#1a73e8", italic=True, wrap=True)
sheet.format("A1", bold=False)        # False (or None) removes that format
```

Formats: bold, italic, underline, strike, fg (text colour), bg (fill), size, font, numfmt
(number format like "0.00", "0%", "yyyy-mm-dd"), align ("left", "center", "right"),
valign ("top", "center"), wrap, indent, rotate.

Borders: border="all" (every cell), "outline" (around the block), "top" / "bottom" / "left" / "right",
or "none"; with border_style ("thin", "medium", "thick", "double", "dashed", "dotted") and border_color.

```
sheet.format("A1:D10", border="all")
sheet.format("A1:D1", border="bottom", border_style="double", border_color="#1a73e8")
```

# Layout and more

```
sheet.freeze(1)                       # keep the top row in view (freeze(cols=1) = column A, freeze("B2") = both)
sheet.unfreeze()
sheet.merge("A1:D1")                  # sheet.unmerge("A1:D1")
sheet.width("B", 20)                  # characters; width("B:D", 12) for several
sheet.autofit()                       # columns as wide as their text (autofit("A:C") for some)
sheet.height(1, 30)                   # row height in points
sheet.hide_rows("3:5")                # show_rows, hide_cols("C"), show_cols
sheet.insert_rows(before=3, count=2)  # delete_rows(3, 2), insert_cols("C"), delete_cols("C")
sheet.sort("A1:E40", by="C", descending=True)       # by=["C", "A"] for several; header=False if no header
sheet.filter("A1:E40")                              # filter buttons
sheet.filter("A1:E40", column="E", values=["Delivered", "Shipped"])   # show only those rows
sheet.clear_filter()
sheet.dropdown("D2:D50", ["Yes", "No", "Maybe"])    # a drop-down list in the cells
sheet.highlight("C2:C40", ">", 100, bg="#ffc7ce", fg="#9c0006")
sheet.highlight("A2:A40", "contains", "urgent", bold=True)
sheet.color_scale("C2:C40")                         # red - yellow - green
sheet.data_bar("D2:D40", color="#638ec6")
sheet.clear_rules()                                 # remove highlights and drop-downs (clear_rules("C:C"))
sheet.comment("A1", "Checked by Ana")
sheet.link("A2", "https://example.com")
sheet.tab_color("#0d9488")
new = book.add_sheet("Summary")                     # a new sheet; write to it like any other
new["A1"] = "=SUM(Data!C2:C40)"
book.rename_sheet("Sheet1", "Data")
```

highlight rules: > < >= <= = != between (value, value2), contains, begins, ends, blank, top / bottom
(value = how many), above_average, below_average, formula (value = "=$C2>$D2"). Formats: bg, fg, bold,
italic, underline, strike.

All of these are part of the same single undo step, and a script that fails changes none of them.

# The selection

```
selection.address      # "B2:D9"
selection.row          # first row (1-based)
selection.col          # first column (1-based)
selection.rows()       # the selected values, row by row
```

# Other sheets

```
book.active                  # the sheet you're looking at (same as sheet)
book["Sales"]                # another sheet by name
book["Sales"]["A1"] = 5
for s in book.sheets:        # every sheet
    print(s.name, s.last_row)
book.save()                  # save the workbook after the script has finished
```

# Handy helper

```
ref(2, 3)                    # "C2" - a cell address from row and column numbers
```

# More examples

Number the rows in column A:

```
for r in range(2, sheet.last_row + 1):
    sheet.cell(r, 1, r - 1)
```

Highlight negative numbers in red:

```
for r, row in enumerate(sheet.rows(), start=1):
    for c, v in enumerate(row, start=1):
        if isinstance(v, (int, float)) and v < 0:
            sheet.format(ref(r, c), fg="#d93025", bold=True)
```

Add a total under every numeric column:

```
last = sheet.last_row
for c in range(1, sheet.last_col + 1):
    col = ref(1, c)[:-1]                   # "A", "B", ...
    if isinstance(sheet.cell(2, c), (int, float)):
        sheet.cell(last + 1, c, f"=SUM({col}2:{col}{last})")
sheet.format(f"A{last + 1}:{ref(last + 1, sheet.last_col)}", bold=True)
```

Split "Last, First" names into two columns:

```
for r in range(2, sheet.last_row + 1):
    name = sheet.cell(r, 1)
    if isinstance(name, str) and "," in name:
        last, first = [p.strip() for p in name.split(",", 1)]
        sheet.cell(r, 2, first)
        sheet.cell(r, 3, last)
```

Summary statistics of the selection:

```
import statistics
nums = [v for row in selection.rows() for v in row if isinstance(v, (int, float))]
print("count", len(nums), "mean", statistics.mean(nums), "median", statistics.median(nums))
```

# Other programs (any language)

Data → Run command on selection… sends the selected cells as CSV to a program's input; the CSV the
program prints replaces them (one undo step). Use it for PowerShell, Node, R, a .bat file, or your own
Python with extra packages. For example a command such as:

```
python clean.py
powershell -File clean.ps1
node fix.js
```

where clean.py might be:

```
import csv, sys
rows = list(csv.reader(sys.stdin))
csv.writer(sys.stdout).writerows([[cell.strip().title() for cell in row] for row in rows])
```

If the program prints more rows or columns than were selected, they fill on from the selection's
top-left cell. If it fails, its error message is shown and nothing changes.

# Automation

Run a script on a file without touching it by hand:

```
SimpleSheets.exe --run job.py book.xlsx --quit
```

The file opens, job.py runs on it, and the program closes. Call book.save() in the script to keep its
changes. The result is added to job.log next to the script; the exit code is 1 when the script failed.
Leave out --quit to keep the window open and look at the result.

# Good to know

- Scripts run on the Python built into Simple Sheets. Common modules are there (math, statistics, random,
  datetime, re, json, csv, decimal, fractions, itertools, collections…). Packages such as pandas are not;
  use Run command on selection with your own Python for those.
- A long script keeps the window responsive and can be stopped with ■ Stop (or Esc).
- Protected sheets: a script can only change unlocked cells, like typing.
- Safety: a script can do anything Python can do on your computer (read and write files, use the
  network…). Only run scripts you wrote or trust. Simple Sheets never runs code stored inside a workbook.
