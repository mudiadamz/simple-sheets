"""Recently opened files, newest first, kept in a small JSON file."""
import json
import os

MAX = 8


def config_dir():
    """Where the app keeps its small files (tests point SIMPLE_SHEETS_HOME elsewhere)."""
    return os.environ.get("SIMPLE_SHEETS_HOME") or os.path.join(
        os.environ.get("APPDATA") or os.path.expanduser("~"), "SimpleSheets")


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
