"""New workbook from a template (.xltx / .xltm, or any .xlsx used as one): the template's sheets,
formatting and layout, as a new untitled workbook — saving asks for a name and never touches the template."""
import os
from tkinter import filedialog

from simplecore import recent

TYPES = [("Excel templates", "*.xltx *.xltm"), ("Workbooks", "*.xlsx *.xlsm"), ("All files", "*.*")]


def new_from_template(app, path=None):
    if not app.confirm_discard():
        return False
    if path is None:
        path = filedialog.askopenfilename(filetypes=TYPES, title="New from template", parent=app)
        if not path:
            return False
    if not app.open_file(path):
        return False
    recent.remove(path)
    app.template_src = path            # saving patches the template's content into the new file
    app.path = None
    app.disk_stamp = None
    app.refresh()
    app.status["text"] = f"New workbook from {os.path.basename(path)}"
    return True
