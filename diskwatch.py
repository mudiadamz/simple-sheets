"""Is the open file still the one we loaded? Is another program (Excel) holding it open?"""
import os


def stamp(path):
    """(modified time, size) of a file, or None if it's gone."""
    try:
        st = os.stat(path)
    except OSError:
        return None
    return st.st_mtime_ns, st.st_size


def is_locked(path):
    """True if another program has the file open for writing (Excel does while a workbook is open)."""
    if not path or not os.path.exists(path):
        return False
    try:
        with open(path, "r+b"):
            return False
    except PermissionError:
        return True
    except OSError:
        return False
