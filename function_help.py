"""Insert function (Shift+F3): every function by category, with what it does and an example; and
Evaluate formula: a formula broken into its parts with the value of each, to see where a result comes from."""
import tkinter as tk
from tkinter import ttk

import formula
from autocomplete import SIGNATURES

# name: (category, what it does, example)
HELP = {
    "SUM": ("Math", "Adds numbers", "=SUM(A1:A10)"),
    "AVERAGE": ("Statistics", "Average of numbers", "=AVERAGE(B2:B9)"),
    "MIN": ("Statistics", "Smallest number", "=MIN(A:A)"), "MAX": ("Statistics", "Largest number", "=MAX(A:A)"),
    "COUNT": ("Statistics", "How many cells hold numbers", "=COUNT(A1:A20)"),
    "COUNTA": ("Statistics", "How many cells aren't empty", "=COUNTA(A:A)"),
    "COUNTBLANK": ("Statistics", "How many cells are empty", "=COUNTBLANK(A1:A20)"),
    "COUNTIF": ("Statistics", "Counts cells that match a condition", '=COUNTIF(B:B, ">100")'),
    "COUNTIFS": ("Statistics", "Counts rows that match several conditions", '=COUNTIFS(A:A, "East", B:B, ">5")'),
    "SUMIF": ("Math", "Adds the cells that match a condition", '=SUMIF(A:A, "East", B:B)'),
    "SUMIFS": ("Math", "Adds where several conditions match", '=SUMIFS(C:C, A:A, "East", B:B, "Pens")'),
    "AVERAGEIF": ("Statistics", "Average of cells that match", '=AVERAGEIF(A:A, "East", B:B)'),
    "AVERAGEIFS": ("Statistics", "Average where several conditions match", '=AVERAGEIFS(C:C, A:A, "East")'),
    "MAXIFS": ("Statistics", "Largest value where conditions match", '=MAXIFS(C:C, A:A, "East")'),
    "MINIFS": ("Statistics", "Smallest value where conditions match", '=MINIFS(C:C, A:A, "East")'),
    "PRODUCT": ("Math", "Multiplies numbers", "=PRODUCT(A1:A3)"), "ABS": ("Math", "Number without its sign", "=ABS(-4)"),
    "ROUND": ("Math", "Rounds to some decimals", "=ROUND(2.345, 2)"),
    "ROUNDUP": ("Math", "Rounds away from zero", "=ROUNDUP(2.1, 0)"),
    "ROUNDDOWN": ("Math", "Rounds toward zero", "=ROUNDDOWN(2.9, 0)"),
    "INT": ("Math", "Rounds down to a whole number", "=INT(7.8)"), "TRUNC": ("Math", "Cuts off decimals", "=TRUNC(7.8)"),
    "MOD": ("Math", "Remainder of a division", "=MOD(10, 3)"), "SQRT": ("Math", "Square root", "=SQRT(16)"),
    "POWER": ("Math", "Number to a power", "=POWER(2, 10)"), "EXP": ("Math", "e to a power", "=EXP(1)"),
    "LN": ("Math", "Natural logarithm", "=LN(10)"), "LOG": ("Math", "Logarithm (base 10 or another)", "=LOG(8, 2)"),
    "LOG10": ("Math", "Base-10 logarithm", "=LOG10(1000)"), "SIGN": ("Math", "1, 0 or -1", "=SIGN(-3)"),
    "PI": ("Math", "The number π", "=PI()"), "RAND": ("Math", "Random number 0-1", "=RAND()"),
    "RANDBETWEEN": ("Math", "Random whole number", "=RANDBETWEEN(1, 6)"),
    "IF": ("Logic", "One value when true, another when false", '=IF(A1>10, "big", "small")'),
    "IFS": ("Logic", "First value whose condition is true", '=IFS(A1>90, "A", A1>80, "B", TRUE, "C")'),
    "IFERROR": ("Logic", "A value instead of an error", "=IFERROR(A1/B1, 0)"),
    "IFNA": ("Logic", "A value instead of #N/A", '=IFNA(VLOOKUP(A1, D:E, 2, 0), "none")'),
    "AND": ("Logic", "TRUE if all are true", "=AND(A1>0, B1>0)"), "OR": ("Logic", "TRUE if any is true", "=OR(A1>0, B1>0)"),
    "NOT": ("Logic", "The opposite", "=NOT(A1)"), "XOR": ("Logic", "TRUE if an odd number are true", "=XOR(A1, B1)"),
    "TRUE": ("Logic", "The value TRUE", "=TRUE()"), "FALSE": ("Logic", "The value FALSE", "=FALSE()"),
    "SWITCH": ("Logic", "Pick a result by matching a value", '=SWITCH(A1, 1, "one", 2, "two", "other")'),
    "CHOOSE": ("Lookup", "The n-th value from a list", '=CHOOSE(2, "a", "b", "c")'),
    "VLOOKUP": ("Lookup", "Finds a row by its first column", "=VLOOKUP(A2, D:F, 3, FALSE)"),
    "HLOOKUP": ("Lookup", "Finds a column by its first row", "=HLOOKUP(\"Q1\", A1:E5, 3, FALSE)"),
    "XLOOKUP": ("Lookup", "Finds a value and returns the matching one", "=XLOOKUP(A2, D:D, F:F, \"none\")"),
    "INDEX": ("Lookup", "The value at a row / column of a range", "=INDEX(A1:C10, 3, 2)"),
    "MATCH": ("Lookup", "Position of a value in a range", "=MATCH(\"Pens\", A:A, 0)"),
    "XMATCH": ("Lookup", "Position of a value (exact, nearest, wildcard)", '=XMATCH("P*", A:A, 2)'),
    "ROW": ("Lookup", "Row number", "=ROW(B5)"), "COLUMN": ("Lookup", "Column number", "=COLUMN(C1)"),
    "ROWS": ("Lookup", "Number of rows in a range", "=ROWS(A1:A10)"),
    "COLUMNS": ("Lookup", "Number of columns in a range", "=COLUMNS(A1:D1)"),
    "OFFSET": ("Lookup", "A range moved from a starting cell", "=SUM(OFFSET(A1, 1, 0, 3, 1))"),
    "INDIRECT": ("Lookup", "A reference written as text", '=INDIRECT("B" & A1)'),
    "ADDRESS": ("Lookup", "A cell address as text", "=ADDRESS(2, 3)"),
    "LEN": ("Text", "Number of characters", "=LEN(A1)"), "LEFT": ("Text", "First characters", "=LEFT(A1, 3)"),
    "RIGHT": ("Text", "Last characters", "=RIGHT(A1, 2)"), "MID": ("Text", "Characters from the middle", "=MID(A1, 2, 3)"),
    "UPPER": ("Text", "CAPITALS", "=UPPER(A1)"), "LOWER": ("Text", "small letters", "=LOWER(A1)"),
    "PROPER": ("Text", "Capitalised Words", "=PROPER(A1)"), "TRIM": ("Text", "Removes extra spaces", "=TRIM(A1)"),
    "CLEAN": ("Text", "Removes unprintable characters", "=CLEAN(A1)"),
    "CONCATENATE": ("Text", "Joins text", '=CONCATENATE(A1, " ", B1)'), "CONCAT": ("Text", "Joins text", "=CONCAT(A1:C1)"),
    "TEXTJOIN": ("Text", "Joins text with a separator", '=TEXTJOIN(", ", TRUE, A1:A5)'),
    "SUBSTITUTE": ("Text", "Replaces text", '=SUBSTITUTE(A1, "-", "/")'),
    "REPLACE": ("Text", "Replaces characters at a position", '=REPLACE(A1, 1, 3, "ABC")'),
    "FIND": ("Text", "Position of text (case matters)", '=FIND("x", A1)'),
    "SEARCH": ("Text", "Position of text (any case, wildcards)", '=SEARCH("x", A1)'),
    "EXACT": ("Text", "TRUE if two texts are exactly the same", "=EXACT(A1, B1)"),
    "REPT": ("Text", "Repeats text", '=REPT("*", 5)'), "TEXT": ("Text", "A number shown with a format", '=TEXT(A1, "0.00")'),
    "VALUE": ("Text", "Text to a number", '=VALUE("12.5")'), "FIXED": ("Text", "A number as text with decimals", "=FIXED(A1, 2)"),
    "CHAR": ("Text", "Character from a code", "=CHAR(65)"), "CODE": ("Text", "Code of the first character", '=CODE("A")'),
    "TEXTBEFORE": ("Text", "Text before a separator", '=TEXTBEFORE(A1, "@")'),
    "TEXTAFTER": ("Text", "Text after a separator", '=TEXTAFTER(A1, "@")'),
    "TEXTSPLIT": ("Text", "Splits text into cells", '=TEXTSPLIT(A1, ",")'),
    "TODAY": ("Date & time", "Today's date", "=TODAY()"), "NOW": ("Date & time", "Date and time now", "=NOW()"),
    "DATE": ("Date & time", "A date from year, month, day", "=DATE(2026, 10, 15)"),
    "TIME": ("Date & time", "A time from hours, minutes, seconds", "=TIME(9, 30, 0)"),
    "YEAR": ("Date & time", "Year of a date", "=YEAR(A1)"), "MONTH": ("Date & time", "Month of a date", "=MONTH(A1)"),
    "DAY": ("Date & time", "Day of the month", "=DAY(A1)"), "HOUR": ("Date & time", "Hour of a time", "=HOUR(A1)"),
    "MINUTE": ("Date & time", "Minute of a time", "=MINUTE(A1)"), "SECOND": ("Date & time", "Second of a time", "=SECOND(A1)"),
    "WEEKDAY": ("Date & time", "Day of the week as a number", "=WEEKDAY(A1, 2)"),
    "WEEKNUM": ("Date & time", "Week of the year", "=WEEKNUM(A1)"),
    "ISOWEEKNUM": ("Date & time", "ISO week of the year", "=ISOWEEKNUM(A1)"),
    "DAYS": ("Date & time", "Days between two dates", "=DAYS(B1, A1)"),
    "DATEDIF": ("Date & time", "Years / months / days between dates", '=DATEDIF(A1, TODAY(), "Y")'),
    "DATEVALUE": ("Date & time", "Text to a date", '=DATEVALUE("15 Oct 2026")'),
    "N": ("Information", "A number as it is (TRUE 1, text 0)", "=N(A1)"),
    "T": ("Text", "Text as it is (anything else empty)", "=T(A1)"),
    "IMAGE": ("Lookup", "A picture from the web shown in the cell (sizing 0 fit, 1 fill, 2 original, "
                                    "3 custom)", '=IMAGE("https://example.com/logo.png", "Logo")'),
    "TIMEVALUE": ("Date & time", "Text to a time (fraction of a day)", '=TIMEVALUE("2:30 PM")'),
    "EDATE": ("Date & time", "Same day some months later", "=EDATE(A1, 3)"),
    "EOMONTH": ("Date & time", "Last day of a month", "=EOMONTH(A1, 0)"),
    "NETWORKDAYS": ("Date & time", "Working days between dates", "=NETWORKDAYS(A1, B1)"),
    "WORKDAY": ("Date & time", "Date some working days later", "=WORKDAY(A1, 10)"),
    "PMT": ("Finance", "Loan payment per period", "=PMT(5%/12, 60, 10000)"),
    "FV": ("Finance", "Future value of savings", "=FV(4%/12, 120, -100)"),
    "PV": ("Finance", "Present value", "=PV(8%/12, 240, 500)"),
    "NPER": ("Finance", "Number of payments", "=NPER(1%, -100, 1000)"),
    "RATE": ("Finance", "Interest rate per period", "=RATE(60, -188.71, 10000)"),
    "NPV": ("Finance", "Net present value of cash flows", "=NPV(10%, B2:B6)"),
    "IRR": ("Finance", "Internal rate of return", "=IRR(B2:B6)"),
    "MEDIAN": ("Statistics", "Middle value", "=MEDIAN(A1:A9)"), "MODE": ("Statistics", "Most common value", "=MODE(A1:A9)"),
    "MODE.SNGL": ("Statistics", "Most common value", "=MODE.SNGL(A1:A9)"),
    "LARGE": ("Statistics", "k-th largest", "=LARGE(A:A, 2)"), "SMALL": ("Statistics", "k-th smallest", "=SMALL(A:A, 2)"),
    "RANK": ("Statistics", "Rank of a number in a list", "=RANK(A2, A:A)"),
    "RANK.EQ": ("Statistics", "Rank of a number in a list", "=RANK.EQ(A2, A:A)"),
    "STDEV": ("Statistics", "Standard deviation (sample)", "=STDEV(A1:A9)"),
    "STDEV.S": ("Statistics", "Standard deviation (sample)", "=STDEV.S(A1:A9)"),
    "STDEV.P": ("Statistics", "Standard deviation (population)", "=STDEV.P(A1:A9)"),
    "STDEVP": ("Statistics", "Standard deviation (population)", "=STDEVP(A1:A9)"),
    "VAR": ("Statistics", "Variance (sample)", "=VAR(A1:A9)"), "VAR.S": ("Statistics", "Variance (sample)", "=VAR.S(A1:A9)"),
    "VAR.P": ("Statistics", "Variance (population)", "=VAR.P(A1:A9)"),
    "VARP": ("Statistics", "Variance (population)", "=VARP(A1:A9)"),
    "PERCENTILE": ("Statistics", "Value at a percentile", "=PERCENTILE(A:A, 0.9)"),
    "PERCENTILE.INC": ("Statistics", "Value at a percentile", "=PERCENTILE.INC(A:A, 0.9)"),
    "QUARTILE": ("Statistics", "Quartile of a list", "=QUARTILE(A:A, 1)"),
    "QUARTILE.INC": ("Statistics", "Quartile of a list", "=QUARTILE.INC(A:A, 1)"),
    "COUNTUNIQUE": ("Statistics", "Number of different values", "=COUNTUNIQUE(A:A)"),
    "SUMPRODUCT": ("Math", "Multiplies ranges and adds the results", "=SUMPRODUCT(B2:B9, C2:C9)"),
    "SUMSQ": ("Math", "Sum of squares", "=SUMSQ(3, 4)"), "CEILING": ("Math", "Rounds up to a multiple", "=CEILING(2.1, 0.5)"),
    "CEILING.MATH": ("Math", "Rounds up to a multiple", "=CEILING.MATH(2.1)"),
    "FLOOR": ("Math", "Rounds down to a multiple", "=FLOOR(2.9, 0.5)"),
    "FLOOR.MATH": ("Math", "Rounds down to a multiple", "=FLOOR.MATH(2.9)"),
    "MROUND": ("Math", "Rounds to the nearest multiple", "=MROUND(10, 3)"),
    "EVEN": ("Math", "Rounds up to an even number", "=EVEN(3)"), "ODD": ("Math", "Rounds up to an odd number", "=ODD(2)"),
    "FACT": ("Math", "Factorial", "=FACT(5)"), "GCD": ("Math", "Greatest common divisor", "=GCD(24, 36)"),
    "LCM": ("Math", "Least common multiple", "=LCM(4, 6)"),
    "SIN": ("Math", "Sine (radians)", "=SIN(PI()/2)"), "COS": ("Math", "Cosine (radians)", "=COS(0)"),
    "TAN": ("Math", "Tangent (radians)", "=TAN(1)"), "ASIN": ("Math", "Arcsine", "=ASIN(1)"),
    "ACOS": ("Math", "Arccosine", "=ACOS(0)"), "ATAN": ("Math", "Arctangent", "=ATAN(1)"),
    "ATAN2": ("Math", "Arctangent of x, y", "=ATAN2(1, 1)"), "DEGREES": ("Math", "Radians to degrees", "=DEGREES(PI())"),
    "RADIANS": ("Math", "Degrees to radians", "=RADIANS(180)"),
    "ISBLANK": ("Information", "TRUE if the cell is empty", "=ISBLANK(A1)"),
    "ISNUMBER": ("Information", "TRUE for a number", "=ISNUMBER(A1)"), "ISTEXT": ("Information", "TRUE for text", "=ISTEXT(A1)"),
    "ISERROR": ("Information", "TRUE for any error", "=ISERROR(A1)"), "ISNA": ("Information", "TRUE for #N/A", "=ISNA(A1)"),
    "ISEVEN": ("Information", "TRUE for an even number", "=ISEVEN(A1)"),
    "ISODD": ("Information", "TRUE for an odd number", "=ISODD(A1)"),
    "NA": ("Information", "The #N/A error", "=NA()"), "ISOMITTED": ("Information", "In a LAMBDA: was an argument left out", ""),
    "SEQUENCE": ("Dynamic arrays", "A list of numbers (spills)", "=SEQUENCE(10)"),
    "FILTER": ("Dynamic arrays", "Rows that match a condition (spills)", '=FILTER(A2:C99, B2:B99="East")'),
    "UNIQUE": ("Dynamic arrays", "Each value once (spills)", "=UNIQUE(A2:A99)"),
    "SORT": ("Dynamic arrays", "Sorted copy (spills)", "=SORT(A2:C99, 2, -1)"),
    "SORTBY": ("Dynamic arrays", "Sorted by another range (spills)", "=SORTBY(A2:A99, B2:B99)"),
    "TRANSPOSE": ("Dynamic arrays", "Rows become columns", "=TRANSPOSE(A1:C2)"),
    "RANDARRAY": ("Dynamic arrays", "Random numbers (spills)", "=RANDARRAY(5)"),
    "VSTACK": ("Dynamic arrays", "Ranges one under another", "=VSTACK(A1:B3, D1:E3)"),
    "HSTACK": ("Dynamic arrays", "Ranges side by side", "=HSTACK(A1:A3, C1:C3)"),
    "TAKE": ("Dynamic arrays", "First / last rows", "=TAKE(A1:C99, 5)"),
    "DROP": ("Dynamic arrays", "Without the first rows", "=DROP(A1:C99, 1)"),
    "CHOOSEROWS": ("Dynamic arrays", "Some rows of a range", "=CHOOSEROWS(A1:C9, 1, 3)"),
    "CHOOSECOLS": ("Dynamic arrays", "Some columns of a range", "=CHOOSECOLS(A1:C9, 3, 1)"),
    "TOCOL": ("Dynamic arrays", "A range as one column", "=TOCOL(A1:C3)"),
    "TOROW": ("Dynamic arrays", "A range as one row", "=TOROW(A1:C3)"),
    "WRAPROWS": ("Dynamic arrays", "A list wrapped into rows", "=WRAPROWS(A1:A9, 3)"),
    "WRAPCOLS": ("Dynamic arrays", "A list wrapped into columns", "=WRAPCOLS(A1:A9, 3)"),
    "MMULT": ("Dynamic arrays", "Matrix product", "=MMULT(A1:B2, D1:D2)"),
    "FREQUENCY": ("Dynamic arrays", "How many values fall in each bin", "=FREQUENCY(A1:A99, D1:D5)"),
    "LET": ("Lambda", "Names for parts of a formula", "=LET(x, A1*2, x+x)"),
    "LAMBDA": ("Lambda", "Your own function", "=LAMBDA(x, x*2)(A1)"),
    "MAP": ("Lambda", "A LAMBDA applied to each value", "=MAP(A1:A5, LAMBDA(v, v*2))"),
    "BYROW": ("Lambda", "A LAMBDA applied to each row", "=BYROW(A1:C5, LAMBDA(r, SUM(r)))"),
    "BYCOL": ("Lambda", "A LAMBDA applied to each column", "=BYCOL(A1:C5, LAMBDA(c, MAX(c)))"),
    "REDUCE": ("Lambda", "Combines all values into one", "=REDUCE(0, A1:A5, LAMBDA(a, v, a+v))"),
    "SCAN": ("Lambda", "Running results", "=SCAN(0, A1:A5, LAMBDA(a, v, a+v))"),
    "MAKEARRAY": ("Lambda", "A block built by a LAMBDA", "=MAKEARRAY(3, 3, LAMBDA(r, c, r*c))"),
}


def describe(name):
    cat, what, example = HELP.get(name, ("Other", "", ""))
    args = SIGNATURES.get(name)
    return cat, what, example, f"{name}({args})" if args is not None else f"{name}(…)"


def categories():
    cats = {}
    for name in sorted(formula.FUNCS):
        cats.setdefault(describe(name)[0], []).append(name)
    return dict(sorted(cats.items()))


class InsertFunction(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app)
        from theme import C
        self.app = app
        self.title("Insert function")
        self.transient(app)
        f = ttk.Frame(self, padding=10)
        f.pack(fill="both", expand=True)
        top = ttk.Frame(f)
        top.pack(fill="x")
        ttk.Label(top, text="Search").pack(side="left")
        self.search = tk.StringVar(master=self)
        e = ttk.Entry(top, textvariable=self.search, width=24)
        e.pack(side="left", padx=4)
        self.cat = ttk.Combobox(top, values=["All"] + list(categories()), state="readonly", width=16)
        self.cat.set("All")
        self.cat.pack(side="left")
        self.list = tk.Listbox(f, height=14, width=24, exportselection=False, background=C["field_bg"],
                               foreground=C["fg"])
        self.list.pack(side="left", fill="y", pady=6)
        right = ttk.Frame(f, padding=(10, 6, 0, 0))
        right.pack(side="left", fill="both", expand=True)
        self.sig = ttk.Label(right, text="", font=("Consolas", 10, "bold"), wraplength=320)
        self.sig.pack(anchor="w")
        self.what = ttk.Label(right, text="", wraplength=320)
        self.what.pack(anchor="w", pady=4)
        self.example = ttk.Label(right, text="", foreground=C["muted"], wraplength=320)
        self.example.pack(anchor="w")
        ttk.Button(right, text="Insert", command=self.insert).pack(anchor="e", side="bottom")
        self.search.trace_add("write", lambda *a: self.fill())
        self.cat.bind("<<ComboboxSelected>>", lambda e: self.fill())
        self.list.bind("<<ListboxSelect>>", lambda e: self.show())
        self.list.bind("<Double-1>", lambda e: self.insert())
        self.bind("<Return>", lambda e: self.insert())
        self.bind("<Escape>", lambda e: self.destroy())
        self.fill()
        e.focus_set()

    def fill(self):
        q = self.search.get().strip().upper()
        cat = self.cat.get()
        names = [n for n in sorted(formula.FUNCS)
                 if (cat == "All" or describe(n)[0] == cat)
                 and (not q or q in n or q.lower() in describe(n)[1].lower())]
        self.names = names
        self.list.delete(0, "end")
        for n in names:
            self.list.insert("end", n)
        if names:
            self.list.selection_set(0)
            self.show()

    def show(self):
        sel = self.list.curselection()
        if not sel:
            return
        cat, what, example, sig = describe(self.names[sel[0]])
        self.sig.configure(text=sig)
        self.what.configure(text=f"{what}  ({cat})" if what else cat)
        self.example.configure(text=f"Example: {example}" if example else "")

    def insert(self):
        """Put =NAME( into the active cell's editor, ready for the arguments."""
        sel = self.list.curselection()
        if not sel:
            return None
        name = self.names[sel[0]]
        g = self.app.table
        self.destroy()
        if g.editor is None:
            g.edit("=")
        e = g.editor
        text = e.get()
        if not text.startswith("="):
            e.delete(0, "end")
            e.insert(0, "=")
        e.insert("insert", name + "(")
        e.focus_set()
        return name


# ---------- evaluate formula ----------

def unparse(node):
    """Expression tree -> formula text (for showing the parts of a formula)."""
    k = node[0]
    if k == "val":
        v = node[1]
        if isinstance(v, str) and not isinstance(v, formula.XLError):
            return '"' + v.replace('"', '""') + '"'
        if isinstance(v, bool):
            return "TRUE" if v else "FALSE"
        if isinstance(v, float):
            return str(int(v)) if v.is_integer() else repr(v)
        if isinstance(v, formula.Array):
            return "{" + ";".join(",".join(unparse(("val", x)) for x in row) for row in v.rows) + "}"
        return str(v)
    if k == "missing":
        return ""
    if k == "ref":
        _, sheet, r1, c1, r2, c2 = node
        a = f"{formula.col_letters(c1)}{r1 + 1}"
        if r2 >= formula.MAXR - 1:
            ref = f"{formula.col_letters(c1)}:{formula.col_letters(c2)}"
        elif c2 >= formula.MAXC - 1:
            ref = f"{r1 + 1}:{r2 + 1}"
        else:
            ref = a if (r1, c1) == (r2, c2) else f"{a}:{formula.col_letters(c2)}{r2 + 1}"
        return f"{formula.quote_sheet(sheet)}!{ref}" if sheet else ref
    if k == "neg":
        return "-" + unparse(node[1])
    if k == "pct":
        return unparse(node[1]) + "%"
    if k == "bin":
        return f"{unparse(node[2])}{node[1]}{unparse(node[3])}"
    if k == "call":
        return f"{node[1]}({','.join(unparse(a) for a in node[2])})"
    if k == "apply":
        return f"{unparse(node[1])}({','.join(unparse(a) for a in node[2])})"
    if k in ("name", "table", "ext"):
        return node[1]
    if k == "spillref":
        return f"{formula.col_letters(node[3])}{node[2] + 1}#"
    return "?"


def show_value(v):
    if isinstance(v, formula.Range):
        if v.h * v.w <= 6:
            rows = formula.as_rows(v)
            return "{" + ";".join(",".join(formula.text(x) if x is not None else "" for x in r) for r in rows) + "}"
        return f"[{v.h}×{v.w} cells]"
    if isinstance(v, formula.Lambda):
        return "LAMBDA"
    if v is None:
        return "(empty)"
    return formula.text(v) if not isinstance(v, formula.XLError) else str(v)


def steps(engine, text, sheet_name):
    """[(depth, part of the formula, its value)], inner parts first, the whole formula last."""
    out = []
    tree = formula.Parser(text.lstrip("=")).parse()

    def walk(node, depth):
        kids = []
        if node[0] in ("neg", "pct"):
            kids = [node[1]]
        elif node[0] == "bin":
            kids = [node[2], node[3]]
        elif node[0] in ("call", "apply") and node[1] not in ("LET", "LAMBDA", "IF", "IFS", "SWITCH", "IFERROR"):
            kids = list(node[2]) if node[0] == "call" else [node[1]] + list(node[2])
        for kid in kids:
            if kid[0] not in ("val", "missing"):
                walk(kid, depth + 1)
        try:
            v = engine.ev(node, sheet_name)
        except formula.FormulaError as e:
            v = e.err
        out.append((depth, unparse(node), show_value(v)))
    walk(tree, 0)
    return out


class EvaluateDialog(tk.Toplevel):
    """Tools → Evaluate formula: each part of the active cell's formula with its value."""

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        g = app.table
        r, c = g.span(*g.cur)[:2]
        text = g.text(r, c)
        self.title(f"Evaluate formula - {formula.col_letters(c)}{r + 1}")
        self.transient(app)
        f = ttk.Frame(self, padding=10)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text=text if text.startswith("=") else "This cell doesn't hold a formula.",
                  font=("Consolas", 10)).pack(anchor="w")
        self.tree = ttk.Treeview(f, columns=("part", "value"), show="headings", height=12)
        self.tree.heading("part", text="Part")
        self.tree.heading("value", text="Value")
        self.tree.column("part", width=320)
        self.tree.column("value", width=160)
        self.tree.pack(fill="both", expand=True, pady=6)
        self.rows = []
        if text.startswith("=") and len(text) > 1:
            app.engine.stack.append((app.sheet, r, c))
            try:
                self.rows = steps(app.engine, text, app.sheet)
            except formula.FormulaError:
                self.rows = []
            finally:
                app.engine.stack.pop()
            for depth, part, value in self.rows:
                self.tree.insert("", "end", values=("   " * depth + part, value))
        ttk.Button(f, text="Close", command=self.destroy).pack(anchor="e")
        self.bind("<Escape>", lambda e: self.destroy())
