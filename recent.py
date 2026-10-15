"""Recently opened files, newest first, kept in a small JSON file."""
import json
import os

MAX = 8


def portable_dir(exe=None):
    """Portable mode: a 'SimpleSheetsData' folder (or a 'portable.txt' file) next to the program keeps
    settings, recent files and recovery copies there - nothing in the user profile. -> that folder or None."""
    import sys
    if exe is None:
        if not getattr(sys, "frozen", False):
            return None
        exe = sys.executable
    here = os.path.dirname(os.path.abspath(exe))
    data = os.path.join(here, "SimpleSheetsData")
    if os.path.isdir(data) or os.path.exists(os.path.join(here, "portable.txt")):
        return data
    return None


def config_dir():
    """Where the app keeps its small files (tests point SIMPLE_SHEETS_HOME elsewhere)."""
    if os.environ.get("SIMPLE_SHEETS_HOME"):
        return os.environ["SIMPLE_SHEETS_HOME"]
    portable = portable_dir()
    if portable:
        return portable
    if os.environ.get("APPDATA"):
        return os.path.join(os.environ["APPDATA"], "SimpleSheets")
    base = os.environ.get("XDG_CONFIG_HOME") or os.path.join(os.path.expanduser("~"), ".config")   # Linux / macOS
    return os.path.join(base, "SimpleSheets")


def _file():
    return os.path.join(config_dir(), "recent.json")


def load():
    try:
        with open(_file(), encoding="utf-8") as f:
            items = json.load(f)
        return [p for p in items if isinstance(p, str)][:MAX]
    except (OSError, ValueError):
        return []


def _save(items):
    try:
        os.makedirs(os.path.dirname(_file()), exist_ok=True)
        with open(_file(), "w", encoding="utf-8") as f:
            json.dump(items[:MAX], f, indent=1)
    except OSError:
        pass   # a missing recent list is not worth an error


def add(path):
    path = os.path.abspath(path)
    items = [p for p in load() if os.path.normcase(p) != os.path.normcase(path)]
    _save([path] + items)


def remove(path):
    _save([p for p in load() if os.path.normcase(p) != os.path.normcase(os.path.abspath(path))])


def clear():
    _save([])
