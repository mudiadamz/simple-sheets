"""Small user preferences (e.g. dark mode), kept next to the recent-files list."""
import json
import os

from recent import config_dir


def _file():
    return os.path.join(config_dir(), "settings.json")


def load():
    try:
        with open(_file(), encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def get(key, default=None):
    return load().get(key, default)


def set(key, value):
    data = load()
    data[key] = value
    try:
        os.makedirs(config_dir(), exist_ok=True)
        with open(_file(), "w", encoding="utf-8") as f:
            json.dump(data, f, indent=1)
    except OSError:
        pass
