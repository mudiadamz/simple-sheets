"""Sheet-level settings (gridlines, headings, tab colour, protection, outline, page setup) and workbook
properties (title, author…, calculation mode): read from .xlsx and written back on save.

Sheet.props keys (only what differs from Excel's default is stored):
  grid: False · headings: False · tab: '#rrggbb' · protect: {'password': legacy hash or '', ...}
  page: {...} (read only, used for printing / PDF)
Workbook props (App.book_props): title, subject, creator, keywords, description, category,
  calc: 'auto' | 'manual', iterate: bool, iterate_count: int, iterate_delta: float
"""
import base64
import hashlib

from formatting import color_hex

DOC_KEYS = ("title", "subject", "creator", "keywords", "description", "category")


# ---------- sheets ----------

def read(ws, theme):
    p = {}
    v = ws.sheet_view
    if v.showGridLines is False:
        p["grid"] = False
    if v.showRowColHeaders is False:
        p["headings"] = False
    tab = ws.sheet_properties.tabColor
    if tab is not None:
        c = color_hex(tab, theme)
        if c:
            p["tab"] = c
    prot = ws.protection
    if prot.sheet:
        p["protect"] = {"password": prot.password or "", "algo": prot.algorithmName, "hash": prot.hashValue,
                        "salt": prot.saltValue, "spin": prot.spinCount}
    page = read_page(ws)
    if page:
        p["page"] = page
    import print_preview
    brk = print_preview.read(ws)
    if brk:
        p["breaks"] = brk
    import whatif
    scen = whatif.read(ws)
    if scen:
        p["scenarios"] = scen
    return p


def read_page(ws):
    """Page setup used for printing: orientation, fit-to-page, margins, print area, print titles."""
    from formula import col_index
    out = {}
    ps = ws.page_setup
    if ps.orientation in ("landscape", "portrait"):
        out["orientation"] = ps.orientation
    if ws.sheet_properties.pageSetUpPr is not None and ws.sheet_properties.pageSetUpPr.fitToPage:
        out["fit"] = (ps.fitToWidth if ps.fitToWidth is not None else 1, ps.fitToHeight if ps.fitToHeight is not None else 1)
    if ps.scale and ps.scale != 100:
        out["scale"] = ps.scale / 100
    m = ws.page_margins
    if m is not None:
        out["margins"] = (m.left, m.right, m.top, m.bottom)
    area = ws.print_area
    if area:
        first = (area if isinstance(area, str) else area[0]).split(",")[0]
        ref = first.split("!")[-1].replace("$", "")
        a, _, b = ref.partition(":")
        import re
        ma, mb = re.fullmatch(r"([A-Z]+)(\d+)", a), re.fullmatch(r"([A-Z]+)(\d+)", b or a)
        if ma and mb:
            out["area"] = (int(ma.group(2)) - 1, col_index(ma.group(1)), int(mb.group(2)) - 1, col_index(mb.group(1)))
    rows = ws.print_title_rows
    if rows:
        a, _, b = rows.replace("$", "").partition(":")
        if a.isdigit():
            out["title_rows"] = (int(a) - 1, int(b or a) - 1)
    cols = ws.print_title_cols
    if cols:
        a, _, b = cols.replace("$", "").partition(":")
        if a.isalpha():
            out["title_cols"] = (col_index(a), col_index(b or a))
    if ws.print_options.gridLines:
        out["gridlines"] = True
    return out


def apply(ws, key, val):
    """One changed sheet setting -> openpyxl worksheet."""
    from openpyxl.styles.colors import Color
    if key == "grid":
        ws.sheet_view.showGridLines = val is not False
    elif key == "headings":
        ws.sheet_view.showRowColHeaders = val is not False
    elif key == "tab":
        ws.sheet_properties.tabColor = Color(rgb="FF" + val[1:].upper()) if val else None
    elif key == "protect":
        from openpyxl.worksheet.protection import SheetProtection
        if not val:
            ws.protection = SheetProtection()
        else:
            p = SheetProtection(sheet=True)
            if val.get("hash"):          # kept as the file had it
                p.algorithmName, p.hashValue, p.saltValue, p.spinCount = val["algo"], val["hash"], val["salt"], val["spin"]
            elif val.get("password"):
                p.set_password(val["password"], already_hashed=True)
            ws.protection = p
    elif key == "breaks":
        import print_preview
        print_preview.write(ws, val)
    elif key == "scenarios":
        import whatif
        whatif.write(ws, val)


def apply_all(wb, changes):
    """changes: {(sheet, 0, key): value} from App.changes('sp')."""
    for (sheet, _, key), val in changes.items():
        if sheet in wb.sheetnames:
            apply(wb[sheet], key, val)


# ---------- protection passwords ----------

def legacy_hash(password):
    """Excel's old 16-bit sheet password hash (what openpyxl writes), e.g. 'CC1A'."""
    from openpyxl.worksheet.protection import hash_password
    return hash_password(password)


def check_password(prot, password):
    """True if `password` unlocks protection `prot` (legacy hash or SHA-512 with salt and spin count)."""
    if not prot:
        return True
    if prot.get("hash") and prot.get("salt"):
        algo = (prot.get("algo") or "SHA-512").replace("-", "").lower()
        try:
            h = hashlib.new(algo, base64.b64decode(prot["salt"]) + password.encode("utf-16-le")).digest()
            for i in range(int(prot.get("spin") or 0)):
                h = hashlib.new(algo, h + i.to_bytes(4, "little")).digest()
        except (ValueError, TypeError):
            return False
        return base64.b64encode(h).decode() == prot["hash"]
    if prot.get("password"):
        return legacy_hash(password).upper() == prot["password"].upper()
    return True


# ---------- workbook ----------

def read_book(path):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True)
    try:
        out = {}
        pr = wb.properties
        for k in DOC_KEYS:
            v = getattr(pr, k, None)
            if v:
                out[k] = str(v)
        scripts = read_scripts(wb)
        if scripts:
            out["scripts"] = scripts
        structure = read_structure(getattr(wb, "security", None))
        if structure is not None:
            out["structure"] = structure
        calc = getattr(wb, "calculation", None)
        if calc is not None:
            if calc.calcMode == "manual":
                out["calc"] = "manual"
            if calc.iterate:
                out["iterate"] = True
                out["iterate_count"] = int(calc.iterateCount or 100)
                out["iterate_delta"] = float(calc.iterateDelta or 0.001)
        return out
    finally:
        wb.close()


SCRIPTS_PROP = "SimpleSheets.CellScripts"    # custom document properties: the JSON in 250-character parts
PART = 250                                    # (Excel keeps custom properties, up to 255 characters each)


def read_scripts(wb):
    """Cell scripts kept in the file: {sheet: {"G2": code}} (cell_scripts.py), or {}."""
    import json
    try:
        props = {p.name: p.value for p in wb.custom_doc_props.props}
    except AttributeError:
        return {}
    parts, i = [], 1
    while f"{SCRIPTS_PROP}.{i}" in props:
        parts.append(str(props[f"{SCRIPTS_PROP}.{i}"] or ""))
        i += 1
    if not parts:
        return {}
    try:
        data = json.loads("".join(parts))
    except ValueError:
        return {}
    return {str(s): {str(ref): str(code) for ref, code in d.items()} for s, d in data.items()
            if isinstance(d, dict)} if isinstance(data, dict) else {}


def write_scripts(wb, scripts):
    import json

    from openpyxl.packaging.custom import StringProperty
    plist = wb.custom_doc_props
    for p in [p for p in plist.props if p.name.startswith(SCRIPTS_PROP + ".")]:
        plist.props.remove(p)
    if scripts:
        text = json.dumps(scripts, ensure_ascii=False, separators=(",", ":"))
        for i in range(0, len(text), PART):
            plist.append(StringProperty(name=f"{SCRIPTS_PROP}.{i // PART + 1}", value=text[i:i + PART]))


def read_structure(sec):
    """Workbook-structure protection (sheets can't be added, deleted, renamed, moved, hidden) -> the dict
    check_password takes ({} = no password), or None when the structure isn't locked."""
    if sec is None or not getattr(sec, "lockStructure", None):
        return None
    out = {"hash": sec.workbookHashValue, "salt": sec.workbookSaltValue, "spin": sec.workbookSpinCount,
           "algo": sec.workbookAlgorithmName, "password": sec.workbookPassword}
    return {k: v for k, v in out.items() if v}


def write_structure(wb, structure):
    """Lock (structure: the dict from read_structure / protect_book) or unlock the workbook's structure; the
    other locks (windows, revisions) stay as the file had them."""
    from openpyxl.workbook.protection import WorkbookProtection
    sec = wb.security
    if structure is None:
        if sec is not None and sec.lockStructure:
            sec.lockStructure = False
            sec._workbookPassword = None              # (openpyxl's setter would try to hash None)
            for k in ("workbookHashValue", "workbookSaltValue", "workbookSpinCount", "workbookAlgorithmName"):
                setattr(sec, k, None)
        return
    if sec is None:
        sec = wb.security = WorkbookProtection()
    sec.lockStructure = True
    if structure.get("hash"):
        sec.workbookHashValue, sec.workbookSaltValue = structure["hash"], structure.get("salt")
        sec.workbookSpinCount, sec.workbookAlgorithmName = structure.get("spin"), structure.get("algo")
    elif structure.get("password"):              # (the legacy hash as stored: openpyxl's setter would hash it again)
        sec.set_workbook_password(structure["password"], already_hashed=True)


def apply_book(wb, props):
    pr = wb.properties
    for k in DOC_KEYS:
        setattr(pr, k, props.get(k) or None)
    write_scripts(wb, props.get("scripts"))
    write_structure(wb, props.get("structure"))
    calc = wb.calculation
    calc.calcMode = "manual" if props.get("calc") == "manual" else "auto"
    calc.iterate = bool(props.get("iterate"))
    if props.get("iterate"):
        calc.iterateCount = int(props.get("iterate_count", 100))
        calc.iterateDelta = float(props.get("iterate_delta", 0.001))
