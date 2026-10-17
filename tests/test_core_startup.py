"""The real start-up order (a fresh process, not conftest's): version.py must tell simplecore who the app is
before any module reads the settings folder - else simplecore refuses (the exe hung on this once)."""
import os
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def test_run_script_from_a_fresh_process(tmp_path):
    script = tmp_path / "s.py"
    script.write_text("sheet['A1'] = 'from script'\nbook.save()\n", encoding="utf-8")
    book = tmp_path / "b.csv"
    book.write_text("x\n", encoding="utf-8")
    env = dict(os.environ, SIMPLE_SHEETS_HOME=str(tmp_path / "home"))
    r = subprocess.run([sys.executable, "simple_sheets.py", "--run", str(script), str(book), "--quit"], cwd=ROOT,
                       env=env, capture_output=True, text=True, timeout=120)
    assert r.returncode == 0, r.stderr[-2000:]
    assert "from script" in book.read_text(encoding="utf-8")


def test_every_module_imports_first_in_a_fresh_process(tmp_path):
    """Any app module may be the first one imported (e.g. by the exe check): none may read settings before
    simplecore knows the app - they import version (directly or through another module) first."""
    code = ("import glob, importlib, os, sys\n"
            "names = sorted(os.path.splitext(os.path.basename(p))[0] for p in glob.glob('*.py'))\n"
            "bad = []\n"
            "for n in names:\n"
            "    if n in ('installer', 'simple_sheets'): continue\n"
            "    r = __import__('subprocess').run([sys.executable, '-c', f'import {n}'], capture_output=True, text=True)\n"
            "    if r.returncode: bad.append(n + ': ' + r.stderr.strip().splitlines()[-1])\n"
            "print('\\n'.join(bad))\n")
    env = dict(os.environ, SIMPLE_SHEETS_HOME=str(tmp_path / "home"))
    r = subprocess.run([sys.executable, "-c", code], cwd=ROOT, env=env, capture_output=True, text=True, timeout=600)
    assert r.stdout.strip() == "", r.stdout
