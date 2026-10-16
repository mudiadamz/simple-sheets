"""Excel-style display formatting: number/date formats, colors, cell styles."""
import calendar
import colorsys
import datetime as dt
import re
from decimal import ROUND_HALF_UP, Decimal

NUMBER = re.compile(r"-?(0|[1-9]\d*)(\.\d+)?([eE][-+]?\d+)?")
ISO_DATE = re.compile(r"\d{4}-\d\d-\d\d([ T]\d\d:\d\d(:\d\d(\.\d+)?)?)?")
EPOCH = dt.datetime(1899, 12, 30)  # Excel day 0
DEFAULT_SIZE = 11.0                # Excel's default font size (pt)
DEFAULT_FONT = "Calibri"           # Excel's default font (cells in it aren't given a "font" key)
# cell borders: style["border"] = {side: [line style, "#rrggbb"]}, sides from this list
SIDES = ("left", "right", "top", "bottom")
BORDER_STYLES = ("thin", "medium", "thick", "dashed", "dotted", "double")

# text-style presets: font keys only, other formatting (fill, align, number format) is kept
FONT_KEYS = ("bold", "italic", "underline", "strike", "size", "fg")
PRESETS = {
    "Title": {"size": 20.0, "bold": True},
    "Heading 1": {"size": 16.0, "bold": True},
    "Heading 2": {"size": 13.0, "bold": True},
    "Paragraph": {},
    "Caption": {"size": 9.0, "italic": True, "fg": "#595959"},
}


def _box(color, style="thin"):
    return {s: [style, color] for s in ("left", "right", "top", "bottom")}


# Excel's built-in cell styles, applied as plain formatting (Format bar → Styles)
CELL_STYLES = {
    "Normal": {},
    "Good": {"bg": "#c6efce", "fg": "#006100"},
    "Bad": {"bg": "#ffc7ce", "fg": "#9c0006"},
    "Neutral": {"bg": "#ffeb9c", "fg": "#9c5700"},
    "Calculation": {"bg": "#f2f2f2", "fg": "#fa7d00", "bold": True, "border": _box("#7f7f7f")},
    "Check Cell": {"bg": "#a5a5a5", "fg": "#ffffff", "bold": True, "border": _box("#3f3f3f", "double")},
    "Explanatory Text": {"fg": "#7f7f7f", "italic": True},
    "Input": {"bg": "#ffcc99", "fg": "#3f3f76", "border": _box("#7f7f7f")},
    "Linked Cell": {"fg": "#fa7d00", "border": {"bottom": ["double", "#ff8001"]}},
    "Note": {"bg": "#ffffcc", "border": _box("#b2b2b2")},
    "Output": {"bg": "#f2f2f2", "fg": "#3f3f3f", "bold": True, "border": _box("#3f3f3f")},
    "Warning Text": {"fg": "#ff0000"},
    "Title": {"size": 18.0, "bold": True, "fg": "#44546a"},
    "Heading 1": {"size": 15.0, "bold": True, "fg": "#44546a", "border": {"bottom": ["thick", "#4472c4"]}},
    "Heading 2": {"size": 13.0, "bold": True, "fg": "#44546a", "border": {"bottom": ["thick", "#a2b8e1"]}},
    "Heading 3": {"bold": True, "fg": "#44546a", "border": {"bottom": ["medium", "#8ea9db"]}},
    "Heading 4": {"bold": True, "fg": "#44546a"},
    "Total": {"bold": True, "border": {"top": ["thin", "#4472c4"], "bottom": ["double", "#4472c4"]}},
}

# default Office theme, in Excel's theme-index order (lt1, dk1, lt2, dk2, accents, links)
THEME = ["FFFFFF", "000000", "E7E6E6", "44546A", "4472C4", "ED7D31",
         "A5A5A5", "FFC000", "5B9BD5", "70AD47", "0563C1", "954F72"]
THEME_KEYS = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2",
              "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"]


def parse_number(s):
    return float(s) if NUMBER.fullmatch(s) else None


def parse_date(s):
    if ISO_DATE.fullmatch(s):
        try:
            return dt.datetime.fromisoformat(s)
        except ValueError:
            return None
    return None


# ---------- formulas in .xlsx files: newer functions carry Excel's "_xlfn." prefix ----------

FUTURE = set("""CONCAT TEXTJOIN IFS SWITCH MAXIFS MINIFS IFNA XOR DAYS ISOWEEKNUM XLOOKUP XMATCH UNIQUE SORTBY SEQUENCE
RANDARRAY LET IMAGE TEXTBEFORE TEXTAFTER TEXTSPLIT VSTACK HSTACK TOCOL TOROW WRAPROWS WRAPCOLS TAKE DROP
CHOOSECOLS CHOOSEROWS EXPAND ARRAYTOTEXT VALUETOTEXT STDEV.S STDEV.P VAR.S VAR.P MODE.SNGL PERCENTILE.INC
PERCENTILE.EXC QUARTILE.INC QUARTILE.EXC RANK.EQ RANK.AVG NORM.DIST NORM.INV CEILING.MATH FLOOR.MATH
AGGREGATE FORECAST.LINEAR""".split())
WORKSHEET_FUTURE = {"FILTER", "SORT"}                 # these are "_xlfn._xlws."
_STRINGS = re.compile(r'("(?:[^"]|"")*")')
_PREFIXED = re.compile(r"(?<![\w.])(?:_xlfn\.)?(?:_xlws\.)?(?=[A-Za-z][\w.]*\s*\()")
_NAMES = re.compile(r"(?<![\w.])([A-Za-z][\w.]*)(?=\s*\()")


def from_file(f):
    """'=_xlfn.XLOOKUP(…)' as stored by Excel -> '=XLOOKUP(…)' (text in quotes untouched)."""
    return "".join(p if i % 2 else _PREFIXED.sub("", p) for i, p in enumerate(_STRINGS.split(f)))


def to_file(f):
    """'=XLOOKUP(…)' -> '=_xlfn.XLOOKUP(…)' for the file, so Excel knows the newer functions."""
    def fix(m):
        name = m.group(1).upper()
        if name in WORKSHEET_FUTURE:
            return "_xlfn._xlws." + m.group(1)
        return "_xlfn." + m.group(1) if name in FUTURE else m.group(1)
    if "(" not in f:
        return f
    return "".join(p if i % 2 else _NAMES.sub(fix, p) for i, p in enumerate(_STRINGS.split(f)))


# ---------- dates and times written as text (DATEVALUE, TIMEVALUE, "15 Oct 2026" + 1) ----------

MONTHS = {}
for _i, _names in enumerate((
        "january jan januari", "february feb febr februari peb", "march mar maret mrt",
        "april apr", "may mei", "june jun juni", "july jul juli", "august aug agustus agu agt",
        "september sep sept", "october oct oktober okt", "november nov nopember", "december dec desember des"),
        start=1):
    for _n in _names.split():
        MONTHS[_n] = _i
WEEKDAYS = ("monday mon tuesday tue tues wednesday wed thursday thu thur thurs friday fri saturday sat sunday sun "
            "senin selasa rabu kamis jumat sabtu minggu").split()
DATE_ORDER = None                  # "dmy" / "mdy" / "ymd" for 3/4/2026; None = Windows' short date setting

_TIME = re.compile(r"(\d{1,2}):(\d{2})(?::(\d{2})(?:\.\d+)?)?\s*([ap]\.?m\.?)?|(\d{1,2})\s*([ap]\.?m\.?)", re.I)


def date_order():
    """Day / month order for dates like 3/4/2026, from Windows' short date format (else month first)."""
    global DATE_ORDER
    if DATE_ORDER:
        return DATE_ORDER
    order = "mdy"
    try:
        import ctypes
        buf = ctypes.create_unicode_buffer(80)
        if ctypes.windll.kernel32.GetLocaleInfoEx(None, 0x1F, buf, 80):     # LOCALE_SSHORTDATE, e.g. dd/MM/yyyy
            first = next((ch for ch in buf.value.lower() if ch in "dmy"), "m")
            order = {"d": "dmy", "y": "ymd"}.get(first, "mdy")
    except (AttributeError, OSError):
        pass
    DATE_ORDER = order
    return order


def _year(y):
    y = int(y)
    return y + (2000 if y < 30 else 1900) if y < 100 else y       # Excel: 00-29 -> 20xx, 30-99 -> 19xx


def _split_time(s):
    """'15 Oct 2026 2:30 PM' -> ('15 Oct 2026', (14, 30, 0)); no time -> (s, None); a bad time -> (s, False)."""
    found = list(_TIME.finditer(s))
    m = found[-1] if found else None                                  # the time is at the end
    if not m or s[m.end():].strip():
        return s, None
    if m.group(1):
        h, mi, sec, ap = int(m.group(1)), int(m.group(2)), int(m.group(3) or 0), m.group(4)
    else:
        h, mi, sec, ap = int(m.group(5)), 0, 0, m.group(6)
    if ap:
        if not 1 <= h <= 12:
            return s, False
        h = h % 12 + (12 if ap.lower().startswith("p") else 0)
    if h > 23 or mi > 59 or sec > 59:
        return s, False
    return s[:m.start()].strip().rstrip(","), (h, mi, sec)


def parse_date_text(s, today=None):
    """A date and/or time written as text, the way Excel's DATEVALUE / TIMEVALUE read it -> datetime or None.
    "2026-10-15", "15/10/2026" (order from Windows), "2026/10/15", "15 Oct 2026", "15-Oct-26", "Oct 15, 2026",
    "October 2026" (day 1), "15 Oct" (this year), "Thu, 15 Oct 2026 14:30", "2:30 PM" (time only, day 0);
    English and Indonesian month names."""
    s = (s or "").strip()
    if not s or len(s) > 60:
        return None
    iso = parse_date(s)
    if iso is not None:
        return iso
    rest, tm = _split_time(s)
    if tm is False:
        return None
    if not rest:
        return dt.datetime(1899, 12, 30, *tm) if tm else None          # a time alone: Excel's day 0
    today = today or dt.date.today()
    words = [w for w in re.split(r"[\s,./-]+", rest.lower()) if w]
    if words and words[0].rstrip(".") in WEEKDAYS:                   # "Thu, 15 Oct 2026"
        words = words[1:]
    if not words or len(words) > 3:
        return None
    names = [i for i, w in enumerate(words) if w.rstrip(".") in MONTHS]
    nums = [w for w in words if w.isdigit()]
    if len(names) + len(nums) != len(words) or len(names) > 1:
        return None
    try:
        if names:                                                     # a month name with 1-2 numbers
            month = MONTHS[words[names[0]].rstrip(".")]
            if len(nums) == 2:
                a, b = nums
                if len(a) == 4:                                       # 2026 Oct 15
                    y, d = int(a), int(b)
                elif names[0] == 0:                                   # Oct 15 2026
                    d, y = int(a), _year(b)
                else:                                                 # 15 Oct 2026
                    d, y = int(a), _year(b)
            elif len(nums) == 1:
                n = int(nums[0])
                if len(nums[0]) <= 2 and n <= 31:                     # 15 Oct / Oct 15: this year
                    d, y = n, today.year
                else:                                                 # Oct 2026 / Oct-99: the 1st
                    d, y = 1, _year(nums[0])
            else:
                return None
        else:                                                         # numbers only: 15/10/2026, 2026/10/15
            if len(nums) == 3:
                if len(nums[0]) == 4:
                    y, month, d = int(nums[0]), int(nums[1]), int(nums[2])
                else:
                    a, b, y = int(nums[0]), int(nums[1]), _year(nums[2])
                    order = date_order()
                    if a > 12 >= b:
                        d, month = a, b
                    elif b > 12 >= a:
                        month, d = a, b
                    else:
                        d, month = (a, b) if order == "dmy" else (b, a)
            elif len(nums) == 2 and all(len(n) <= 2 for n in nums) and re.search(r"[/-]", rest):  # 15/10: this year
                a, b = int(nums[0]), int(nums[1])
                d, month = (a, b) if (date_order() == "dmy" or a > 12) and b <= 12 else (b, a)
                y = today.year
            else:
                return None
        return dt.datetime(y, month, d, *(tm or (0, 0, 0)))
    except ValueError:                                                # 31 Feb, month 13…
        return None


def general(n):
    if n.is_integer() and abs(n) < 1e15:
        return str(int(n))
    return f"{n:.10g}"


# ---------- number formats ----------

def _sections(fmt):
    out, cur, quoted = [], "", False
    for ch in fmt:
        if ch == '"':
            quoted = not quoted
        if ch == ";" and not quoted:
            out.append(cur)
            cur = ""
        else:
            cur += ch
    return out + [cur]


LIT = 0xE000  # literals become private-use chars so they never look like format codes
LIT_RE = re.compile("[-]")
TOKEN = re.compile(r'"([^"]*)"|\\(.)|_(.)|\*(.)|\[\$([^\]-]*)[^\]]*\]|\[([^\]]*)\]|(.)', re.S)


def _parts(sec):
    """-> (code, lits): format code with literal text i replaced by one private-use char LIT+i."""
    code, lits = [], []
    for m in TOKEN.finditer(sec):
        quoted, esc, pad, fill, cur, bracket, ch = m.groups()
        if ch is not None:
            code.append(ch)
            continue
        if bracket is not None:  # [Red], [>100] ... ; keep elapsed time [h] [mm] [ss]
            if re.fullmatch(r"[hms]+", bracket, re.I):
                code.append(bracket)
            continue
        lit = next(v for v in (quoted, esc, " " if pad else None, "" if fill else None, cur) if v is not None)
        code.append(chr(LIT + len(lits)))
        lits.append(lit)
    return "".join(code), lits


def _lits(code, lits):
    return LIT_RE.sub(lambda m: lits[ord(m.group(0)) - LIT], code)


def is_date(sec):
    code = re.sub(r"\.0+", "", _parts(sec)[0])
    code = LIT_RE.sub("", code)
    return bool(re.search(r"[ymdhs]", code, re.I)) and not re.search(r"[#?0]", code) \
        and "general" not in code.lower()


DATE_TOKEN = re.compile(r"([-])|yyyy|yy|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s"
                        r"|am/pm|a/p|\.0+|.", re.I | re.S)


def fmt_date(d, sec):
    code, lits = _parts(sec)
    toks = list(DATE_TOKEN.finditer(code))
    ampm = re.search(r"am/pm|a/p", code, re.I) is not None
    out, last = [], None
    for i, m in enumerate(toks):
        t = m.group(0)
        tl = t.lower()
        if m.group(1):
            out.append(lits[ord(m.group(1)) - LIT])
            continue
        if tl in ("mm", "m"):  # minutes when after an hour or before seconds
            nxt = next((x.group(0).lower() for x in toks[i + 1:] if x.group(0).lower()[:1] in "ymdhs"), "")
            if last == "h" or nxt.startswith("s"):
                out.append(f"{d.minute:02d}" if tl == "mm" else str(d.minute))
                last = "m"
                continue
        if tl == "yyyy":
            s = f"{d.year:04d}"
        elif tl == "yy":
            s = f"{d.year % 100:02d}"
        elif tl == "mmmmm":
            s = calendar.month_name[d.month][0]
        elif tl == "mmmm":
            s = calendar.month_name[d.month]
        elif tl == "mmm":
            s = calendar.month_abbr[d.month]
        elif tl == "mm":
            s = f"{d.month:02d}"
        elif tl == "m":
            s = str(d.month)
        elif tl == "dddd":
            s = calendar.day_name[d.weekday()]
        elif tl == "ddd":
            s = calendar.day_abbr[d.weekday()]
        elif tl == "dd":
            s = f"{d.day:02d}"
        elif tl == "d":
            s = str(d.day)
        elif tl in ("hh", "h"):
            h = (d.hour % 12 or 12) if ampm else d.hour
            s = f"{h:02d}" if tl == "hh" else str(h)
        elif tl in ("ss", "s"):
            s = f"{d.second:02d}" if tl == "ss" else str(d.second)
        elif tl == "am/pm":
            s = "AM" if d.hour < 12 else "PM"
        elif tl == "a/p":
            s = "A" if d.hour < 12 else "P"
        elif tl.startswith(".0"):
            s = ("." + f"{d.microsecond:06d}")[:len(tl)]
        else:
            s = t
        out.append(s)
        if tl[:1] in "ymdhs":
            last = tl[0]
    return "".join(out)


def fmt_number(n, fmt):
    secs = _sections(fmt)
    i = 0 if n > 0 or len(secs) == 1 else (1 if n < 0 else (2 if len(secs) > 2 else 0))
    sec = secs[i]
    neg = n < 0 and i == 0
    n = abs(n)
    if sec == "":
        return ""
    if is_date(sec):
        return fmt_date(EPOCH + dt.timedelta(days=n), sec)
    code, lits = _parts(sec)
    if "general" in code.lower():
        return ("-" if neg else "") + general(n)
    idx = [k for k, ch in enumerate(code) if ch in "0#?"]
    if not idx:  # literal-only section, e.g. "-"
        return _lits(code.replace("@", ""), lits)
    pre, core, post = code[:idx[0]], code[idx[0]:idx[-1] + 1], code[idx[-1] + 1:]
    if "/" in core:  # fractions: not supported
        return ("-" if neg else "") + general(n)
    if pre.endswith("."):
        pre, core = pre[:-1], "." + core
    scale = 0
    while post.startswith(","):
        scale, post = scale + 1, post[1:]
    v = n * 100 ** (pre + post).count("%") / 1000 ** scale
    sci = re.fullmatch(r"([0#?.,]*)[eE]([+-])([0#?]+)", core)
    if sci:
        mant, sign, expd = sci.groups()
        dec = sum(ch in "0#?" for ch in mant.partition(".")[2])
        m, e = f"{v:.{dec}E}".split("E")
        e = int(e)
        s = f"{m}E{'-' if e < 0 else ('+' if sign == '+' else '')}{abs(e):0{len(expd)}d}"
        zero = v == 0
    else:
        ip, dot, dp = core.partition(".")
        dmax = sum(ch in "0#?" for ch in dp)
        dmin = dp.count("0")
        q = Decimal(repr(v)).quantize(Decimal(1).scaleb(-dmax), rounding=ROUND_HALF_UP)
        zero = q == 0
        s = f"{q:{',' if ',' in ip else ''}.{dmax}f}"
        whole, _, frac = s.partition(".")
        while len(frac) > dmin and frac.endswith("0"):
            frac = frac[:-1]
        imin = ip.count("0")
        if whole == "0" and imin == 0:
            whole = ""
        elif "," not in whole and len(whole) < imin:
            whole = whole.zfill(imin)
        s = whole + ("." + frac if dot and (frac or dmax) else "")
    out = _lits(pre, lits) + s + _lits(post, lits)
    return ("-" + out) if neg and not zero else out


NUMBER_FORMATS = [   # (menu label, format code or None for General)
    ("General", None),
    ("Number   1,234.56", "#,##0.00"),
    ("Whole number   1,235", "#,##0"),
    ("Currency   $1,234.56", '"$"#,##0.00'),
    ("Percent   12%", "0%"),
    ("Percent   12.34%", "0.00%"),
    ("Scientific   1.23E+03", "0.00E+00"),
    ("Date   2024-03-05", "yyyy-mm-dd"),
    ("Date   5 Mar 2024", "d mmm yyyy"),
    ("Long date   Tuesday, March 5, 2024", "dddd, mmmm d, yyyy"),
    ("Time   2:07 PM", "h:mm AM/PM"),
    ("Date and time   2024-03-05 14:07", "yyyy-mm-dd hh:mm"),
    ("Text", "@"),
]


def change_decimals(fmt, delta, sample=""):
    """Number format with one decimal place more (delta=1) or fewer (-1), like Excel's buttons.
    General takes the number of decimals the value shows now. Dates / text are left alone."""
    if not fmt or fmt == "General":
        n = parse_number(sample) if sample else None
        if n is None:
            return fmt
        dec = len(sample.split(".")[1]) if "." in sample and "e" not in sample.lower() else 0
        dec = max(dec + delta, 0)
        return "0" + ("." + "0" * dec if dec else "")
    sections = _sections(fmt)
    if fmt == "@" or is_date(sections[0]):
        return fmt
    out = []
    for sec in sections:
        code = sec
        cut = re.search(r"[eE][+-]", code)
        head, tail = (code[:cut.start()], code[cut.start():]) if cut else (code, "")
        found = list(re.finditer(r"[0#?](?:\.[0#?]*)?", head))   # the number's placeholders
        if not found:
            out.append(sec)
            continue
        m = found[-1]
        whole, _, dec = m.group(0).partition(".")
        dec = "0" * max(len(dec) + delta, 0)
        new = whole + ("." + dec if dec else "")
        out.append(head[:m.start()] + new + head[m.end():] + tail)
    return ";".join(out)


def render(text, numfmt):
    """Cell text -> what Excel would show for that number format."""
    if not numfmt or not text or numfmt == "@":
        return text
    try:
        n = parse_number(text)
        if n is not None:
            return fmt_number(n, numfmt)
        d = parse_date(text)
        if d is not None:
            sec = _sections(numfmt)[0]
            return fmt_date(d, sec) if is_date(sec) else text
    except Exception:
        pass
    return text


# ---------- colors / styles ----------

def theme_colors(xml):
    """Workbook theme XML -> 12 hex colors in theme-index order."""
    if not xml:
        return THEME
    if isinstance(xml, bytes):
        xml = xml.decode("utf-8", "replace")
    found = dict(re.findall(r'<a:(dk1|lt1|dk2|lt2|accent[1-6]|hlink|folHlink)>\s*'
                            r'<a:(?:srgbClr val|sysClr[^>]*?lastClr)="([0-9A-Fa-f]{6})"', xml))
    return [found.get(k, d) for k, d in zip(THEME_KEYS, THEME)]


def tint(hex6, t):
    r, g, b = (int(hex6[i:i + 2], 16) / 255 for i in (0, 2, 4))
    h, light, s = colorsys.rgb_to_hls(r, g, b)
    light = light * (1 + t) if t < 0 else light * (1 - t) + t
    return "#%02x%02x%02x" % tuple(round(x * 255) for x in colorsys.hls_to_rgb(h, light, s))


def color_hex(color, theme):
    """openpyxl Color -> '#rrggbb' or None."""
    if color is None:
        return None
    try:
        if color.type == "rgb":
            if not isinstance(color.rgb, str) or len(color.rgb) < 6:
                return None
            base = color.rgb[-6:]
        elif color.type == "theme":
            base = theme[color.theme] if color.theme < len(theme) else None
        elif color.type == "indexed":
            from openpyxl.styles.colors import COLOR_INDEX
            if color.indexed >= len(COLOR_INDEX) or color.indexed in (64, 65):
                return None
            base = COLOR_INDEX[color.indexed][-6:]
        else:
            return None
        if base is None:
            return None
        return tint(base, color.tint) if color.tint else "#" + base.lower()
    except Exception:
        return None


def cell_style(cell, theme, cache):
    """openpyxl cell -> small style dict (or None for default look). Cached per style combo."""
    # Cache by openpyxl's style ids. NOT by id(cell.font): openpyxl returns a new temporary proxy
    # on every access, Python reuses freed ids, and cells then got another cell's (often empty) style.
    arr = getattr(cell, "_style", None)
    if not hasattr(arr, "fontId"):
        arr = getattr(cell, "style_array", None)       # read-only cells
    if arr is None:
        return None                                     # empty read-only cell: no style
    key = (arr.fontId, arr.fillId, arr.alignmentId, arr.numFmtId, arr.borderId, getattr(arr, "protectionId", 0))
    if key in cache:
        return cache[key]
    font, fill = getattr(cell, "font", None), getattr(cell, "fill", None)
    al, nf = getattr(cell, "alignment", None), getattr(cell, "number_format", None)
    st = {}
    if font is not None:
        for attr, name in (("b", "bold"), ("i", "italic"), ("u", "underline"), ("strike", "strike")):
            if getattr(font, attr, None):
                st[name] = True
        fg = color_hex(font.color, theme)
        if fg and fg != "#000000":
            st["fg"] = fg
        if font.sz and float(font.sz) != DEFAULT_SIZE:
            st["size"] = float(font.sz)
        if font.name and font.name != DEFAULT_FONT:
            st["font"] = font.name
    if getattr(fill, "fill_type", None) == "solid":
        bg = color_hex(fill.fgColor, theme)
        if bg:
            st["bg"] = bg
    elif getattr(fill, "stop", None):              # a gradient: shown as the average of its colours
        cols = [color_hex(s.color, theme) for s in (fill.stop or [])]
        cols = [x for x in cols if x]
        if cols:
            st["bg"] = "#%02x%02x%02x" % tuple(round(sum(int(x[i:i + 2], 16) for x in cols) / len(cols))
                                                for i in (1, 3, 5))
            st["gradient"] = True
    h = getattr(al, "horizontal", None)
    if h in ("left", "right", "center", "centerContinuous", "justify", "distributed"):
        st["align"] = {"centerContinuous": "center", "distributed": "justify"}.get(h, h)
    if getattr(al, "wrap_text", None):
        st["wrap"] = True
    if getattr(al, "indent", None):
        st["indent"] = int(al.indent)
    rot = getattr(al, "text_rotation", None)
    if rot:
        st["rotate"] = 90 - rot if 90 < rot <= 180 else rot   # Excel: 91-180 means 1-90 degrees downwards
    v = getattr(al, "vertical", None)
    if v in ("top", "center", "justify", "distributed"):   # bottom (the default) isn't stored
        st["valign"] = "top" if v == "top" else "center"
    if nf and nf != "General":
        st["numfmt"] = nf
    border = {}
    b = getattr(cell, "border", None)
    for name in SIDES:
        side = getattr(b, name, None) if b is not None else None
        if side is not None and side.style:
            border[name] = [side.style, color_hex(side.color, theme) or "#000000"]
    if border:
        st["border"] = border
    prot = getattr(cell, "protection", None)
    if prot is not None and prot.locked is False:      # may be edited on a protected sheet
        st["unlocked"] = True
    cache[key] = st or None
    return cache[key]


def describe(st):
    """Style dict -> short human text for the format panel."""
    if not st:
        return "General"
    parts = [st.get("numfmt", "General")]
    if st.get("font"):
        parts.append(st["font"])
    if st.get("size"):
        parts.append(f"{st['size']:g}pt")
    parts += [k.capitalize() for k in ("bold", "italic", "underline", "strike") if st.get(k)]
    if st.get("align"):
        parts.append(f"Align {st['align']}")
    if st.get("valign"):
        parts.append({"top": "Top", "center": "Middle"}.get(st["valign"], st["valign"]))
    if st.get("wrap"):
        parts.append("Wrap")
    if st.get("border"):
        parts.append("Borders")
    if st.get("fg"):
        parts.append(f"Text {st['fg']}")
    if st.get("bg"):
        parts.append(f"Fill {st['bg']}")
    return "  ·  ".join(parts)
