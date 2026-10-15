# Screen readers in Simple Sheets — study (Roadmap 2, G3)

## Where things stand

| Part | Narrator / NVDA today | Why |
|---|---|---|
| Cell grid (`grid.py`) | Silent | A Tk `Canvas` draws text as pixels; it exposes nothing to UI Automation (UIA) or MSAA. |
| Menus (☰, right-click) | Read | Tk menus are native Win32 menus. |
| Message boxes | Read | `tkinter.messagebox` uses the native Windows dialog. |
| File open / save dialogs | Read | Native common dialogs. |
| Our dialogs (ttk) | Partly | ttk widgets are drawn by Tk, not Win32 controls: screen readers see a window and sometimes a name, rarely roles or states. |
| Keyboard | Good | Every command is reachable without a mouse (F10 menu, Shift+F10, Ctrl+Shift+P palette, F6 panes, Ctrl+Alt+5 objects). |

The grid is the real gap: a blind user can't hear which cell they're on or what it holds.

## Options

1. **UIA provider for the grid (ctypes).** Implement `IRawElementProviderSimple` + `IGridProvider` / `ITableProvider`
   with ctypes COM vtables and answer `WM_GETOBJECT` for the canvas's window.
   - + Real support: cell names, values, row / column headers, table navigation commands of the reader.
   - − Large (COM vtables by hand, ~1500 lines), hard to test without a reader; easy to crash the process with a
     wrong vtable. Tk subclasses the window procedure — we'd have to hook it (as `dropfiles.py` already does).
2. **Live region / "announce" channel.** Keep a hidden native Win32 `STATIC` / `EDIT` control (CreateWindowEx via
   ctypes) and put "B5, 1,234, formula =SUM(B1:B4)" into it on every cursor move, raising
   `EVENT_OBJECT_NAMECHANGE` / UIA `Notification` events.
   - + Small (~200 lines), uses the same ctypes window hooking we already have; works with Narrator and NVDA.
   - − Readers hear an announcement, not a table: no "row 5 of 40" navigation of their own.
3. **Self-voicing with SAPI** (`SAPI.SpVoice` through ctypes COM, like `spell.py` does with the spell checker).
   - + Independent of any reader; trivial to test (call counts).
   - − Fights with a running screen reader (two voices); users of readers dislike self-voicing apps.
4. **Accessible mirror window:** a ttk `Treeview` view of the visible block.
   - − Treeview is not exposed better than the canvas; rejected.

## Recommendation

- Build **option 2 first** (announce channel), behind *Settings → Accessibility → Announce the current cell*,
  on by default when Windows reports a screen reader (`SystemParametersInfo(SPI_GETSCREENREADER)`).
  - Announce: cell address, shown value, "formula" + formula text, comment / link present, merged size,
    selection size on Shift+arrows, sheet name on Ctrl+PgUp/PgDn, dialog results ("3 rows sorted").
  - Tests: the text pushed into the channel for each move (no reader needed).
- Then, only if users ask for table navigation, **option 1** for the grid only (dialogs stay ttk).
- Dialogs: give every entry a visible label placed *before* it (readers that do see ttk use the previous label),
  and keep the native message boxes for anything important.

## Checks to run with a real reader (manual)

- Narrator (Win+Ctrl+Enter) and NVDA: open a file, move with arrows, type in a cell, Ctrl+Z, open ☰ with F10,
  run a command from Ctrl+Shift+P, close a "save changes?" question with the keyboard.
- High contrast (View → High contrast) with Windows contrast themes on: no text drawn in fixed colours that vanish.
