"""Help: keyboard shortcuts (F1), the user guide (assets/USER_GUIDE.md), About. Simple Sheets' shortcut
list; the window is simplecore.help_ui."""
from simplecore import help_ui as _core
from simplecore.help_ui import HelpWindow, guide_text, shortcuts_text, show  # noqa: F401

SHORTCUTS = [
    ("Files", [("Ctrl+N", "New window"), ("Ctrl+O", "Open"), ("Ctrl+S", "Save"), ("Ctrl+Shift+S", "Save as"),
               ("Ctrl+P", "Print"), ("F10", "Open the ☰ menu")]),
    ("Moving", [("Arrows", "Move"), ("Ctrl+Arrows", "Jump to the edge of the data"), ("Ctrl+Home / End", "First / last cell"),
                ("Page Up / Down", "Scroll a page"), ("Ctrl+G", "Go to a cell, range or name"),
                ("Ctrl+Page Up / Down", "Previous / next sheet")]),
    ("Selecting", [("Shift+Arrows", "Extend the selection"), ("Ctrl+A", "Select the data block, again for all"),
                   ("Ctrl+click", "Add another block"), ("Shift+click", "Extend to the clicked cell")]),
    ("Editing", [("F2", "Edit the cell"), ("Enter / Tab", "Finish and move down / right"), ("Esc", "Cancel"),
                 ("Ctrl+Enter / Alt+Enter", "New line inside the cell"), ("Delete", "Clear the cells"),
                 ("Ctrl+Z / Ctrl+Y", "Undo / redo"), ("Ctrl+C / X / V", "Copy / cut / paste"),
                 ("Ctrl+D / Ctrl+R", "Fill down / right"), ("Ctrl+E", "Flash fill"),
                 ("Alt+↓", "Open a cell's drop-down list"), ("Ctrl+K", "Insert a link"),
                 ("Drag the selection's edge", "Move the cells"),
                 ("Drag a selected column / row header", "Move whole columns / rows")]),
    ("Formulas", [("=", "Start a formula; a list of functions appears"), ("Tab", "Take the suggested function"),
                  ("Arrows (after = ( , or an operator)", "Point at a cell; Shift+arrows for a range"),
                  ("F9", "Recalculate (manual calculation)"), ("Ctrl+`", "Show formulas")]),
    ("Formatting", [("Ctrl+B / I / U", "Bold / italic / underline"),
                    ("Ctrl+Shift+> / <", "Bigger / smaller text"), ("🖌", "Format painter"),
                    ("Alt+Shift+→ / ←", "Group / ungroup rows or columns")]),
    ("Find & view", [("Ctrl+F / Ctrl+H", "Find / replace"), ("F7", "Spelling"),
                     ("Ctrl + / Ctrl - / Ctrl+wheel", "Zoom"), ("F11", "Full screen"), ("F1", "This list")]),
    ("Keyboard only", [("Ctrl+Shift+P", "Command palette: type to find any command"),
                       ("Ctrl+Shift+A", "AI assistant"),
                       ("Ctrl+Alt+5", "Pick charts / pictures (Tab: the next one)"),
                       ("Arrows / Shift+Arrows / Delete", "Move / resize / remove the picked chart or picture"),
                       ("F6 / Shift+F6", "Cells → formula bar → name box → sheet tabs"),
                       ("Shift+F10 / Menu key", "Right-click menu"), ("Shift+F3", "Insert a function")]),
]

_core.set_shortcuts(SHORTCUTS)
GROUPS = {g for g, _ in SHORTCUTS}
