"""Modules loaded only on demand (from a menu) must still be real modules and get into the exe."""
import glob
import os
import re

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def app_modules():
    return {os.path.splitext(os.path.basename(p))[0] for p in glob.glob(os.path.join(ROOT, "*.py"))}


def test_on_demand_imports_name_real_modules():
    """__import__("x") / importlib.import_module("x") in app code: a typo would only fail when clicked."""
    mods = app_modules()
    found = set()
    for p in glob.glob(os.path.join(ROOT, "*.py")):
        text = open(p, encoding="utf-8").read()
        found |= set(re.findall(r"""__import__\(\s*["']([\w.]+)["']""", text))
        found |= set(re.findall(r"""import_module\(\s*["']([\w.]+)["']""", text))
    assert "ai_assist" in found and "compare" in found       # the ones the exe once lost
    import importlib.util
    for name in sorted(found):
        assert name in mods or importlib.util.find_spec(name) is not None, f"{name} doesn't exist"


def test_build_bundles_every_app_module():
    bat = open(os.path.join(ROOT, "build.bat"), encoding="utf-8").read()
    assert 'for %%f in ("%~dp0*.py")' in bat and "--hidden-import %%~nf" in bat and "!HIDDEN!" in bat
    assert "EnableDelayedExpansion" in bat and "tools\\check_exe.py" in bat
