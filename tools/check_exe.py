"""Build check: start dist/SimpleSheets.exe with --run and import every app module inside it. Fails (exit 1)
when the exe is missing one - modules loaded only from a menu (__import__ in a lambda) are easy to lose.
Run by build.bat after PyInstaller; also: python tools/check_exe.py"""
import glob
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SKIP = {"installer", "simple_sheets"}      # the setup program is its own exe; the app is the main script


def main():
    exe = os.path.join(HERE, "dist", "SimpleSheets.exe")
    names = sorted(os.path.splitext(os.path.basename(p))[0] for p in glob.glob(os.path.join(HERE, "*.py")))
    names = [n for n in names if n not in SKIP]
    import simplecore  # the shared core goes in too (collected by build.bat)
    names += [f"simplecore.{m}" for m in simplecore.MODULES]
    work = tempfile.mkdtemp(prefix="ss-check-")
    try:
        job = os.path.join(work, "check.py")
        with open(job, "w", encoding="utf-8") as f:
            f.write("import importlib\nmissing = []\n"
                    f"for n in {names!r}:\n"
                    "    try:\n        importlib.import_module(n)\n"
                    "    except Exception as e:\n        missing.append(f'{n} ({type(e).__name__}: {e})')\n"
                    "sheet['A1'] = '\\n'.join(missing) or 'ALL OK'\nbook.save()\n")
        book = os.path.join(work, "check.xlsx")
        with zipfile.ZipFile(book, "w") as z:     # the smallest valid workbook, without needing openpyxl here
            z.writestr("[Content_Types].xml",
                       '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                       '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                       '<Default Extension="xml" ContentType="application/xml"/>'
                       '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.'
                       'spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="'
                       'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
            z.writestr("_rels/.rels",
                       '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
                       'relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/'
                       '2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
            z.writestr("xl/workbook.xml",
                       '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                       'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
                       '<sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>')
            z.writestr("xl/_rels/workbook.xml.rels",
                       '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
                       'relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/'
                       '2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')
            z.writestr("xl/worksheets/sheet1.xml",
                       '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                       '<sheetData/></worksheet>')
        env = dict(os.environ, SIMPLE_SHEETS_HOME=os.path.join(work, "home"))
        subprocess.run([exe, "--run", job, book, "--quit"], env=env, timeout=240)
        with zipfile.ZipFile(book) as z:
            data = z.read("xl/worksheets/sheet1.xml").decode("utf-8", "replace")
            strings = z.read("xl/sharedStrings.xml").decode("utf-8", "replace") if "xl/sharedStrings.xml" in \
                z.namelist() else ""
        result = "ALL OK" if "ALL OK" in data + strings else (strings or data)
        if "ALL OK" not in result:
            log = os.path.join(work, "check.log")
            detail = open(log, encoding="utf-8").read() if os.path.exists(log) else result
            print("The exe is missing modules (or the check script failed):\n" + detail)
            return 1
        print(f"Exe check: all {len(names)} app modules import inside {os.path.basename(exe)}")
        return 0
    finally:
        shutil.rmtree(work, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
