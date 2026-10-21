"""The ☰ app menu: new, open, open recent, save, save as, undo/redo, exit."""
import os
import tkinter as tk
from tkinter import ttk

from simplecore import icons, recent
from simplecore.ai_button import ai_button, disk_icon, refresh_ai_button, save_images, sparkle_icon  # noqa: F401


def build(app, parent):
    """☰ button (with its menu) plus small undo/redo buttons, packed left into `parent`."""
    mb = icons.icon(ttk.Menubutton(parent, width=2), "menu")
    mb.pack(side="left")
    menu = tk.Menu(mb, tearoff=False)
    recent_menu = tk.Menu(menu, tearoff=False, postcommand=lambda: fill_recent(app, recent_menu))
    menu.add_command(label="New", accelerator="Ctrl+N", command=app.new_window)
    menu.add_command(label="Open…", accelerator="Ctrl+O", command=app.open_dialog)
    menu.add_cascade(label="Open recent", menu=recent_menu)
    import book_tabs
    menu.add_command(label="New tab", accelerator="Ctrl+T", command=lambda: book_tabs.new_tab(app))
    menu.add_command(label="Open in new tab…", command=lambda: book_tabs.open_in_tab(app))
    from simplecore import google_drive
    drive = tk.Menu(menu, tearoff=False)
    drive.add_command(label="Open from Google Drive…", command=lambda: google_drive.open_from_drive(app))
    drive.add_command(label="Save to Google Drive…", command=lambda: google_drive.save_to_drive(app))
    drive.add_separator()
    drive.add_command(label="Google account…", command=lambda: google_drive.AccountDialog(app))
    menu.add_cascade(label="Google Drive", menu=drive)
    menu.add_separator()
    menu.add_command(label="Save", accelerator="Ctrl+S", command=app.save)
    menu.add_command(label="Save as…", accelerator="Ctrl+Shift+S", command=app.save_as)
    menu.add_command(label="Print preview…", command=lambda: __import__("print_preview").open_preview(app))
    menu.add_command(label="Export to PDF…", command=lambda: __import__("print_preview").open_preview(app))
    menu.add_command(label="Export all sheets to PDF…",
                     command=lambda: __import__("print_preview").open_preview(app, "All sheets"))
    menu.add_command(label="Print…", accelerator="Ctrl+P", command=lambda: __import__("print_preview").open_preview(app))
    menu.add_separator()
    import find_replace
    menu.add_command(label="Find…", accelerator="Ctrl+F", command=lambda: find_replace.open_dialog(app))
    menu.add_command(label="Replace…", accelerator="Ctrl+H",
                     command=lambda: find_replace.open_dialog(app, replace=True))
    menu.add_command(label="Named ranges…", command=app.manage_names)
    menu.add_separator()
    menu.add_command(label="Undo", accelerator="Ctrl+Z", command=app.undo)
    menu.add_command(label="Redo", accelerator="Ctrl+Y", command=app.redo)
    menu.add_separator()
    import workbook_tools as wt
    view = tk.Menu(menu, tearoff=False, postcommand=lambda: sync_view(app))
    view.add_checkbutton(label="Gridlines", variable=app.grid_var,
                         command=lambda: wt.toggle_sheet_setting(app, "grid"))
    view.add_checkbutton(label="Headings", variable=app.headings_var,
                         command=lambda: wt.toggle_sheet_setting(app, "headings"))
    view.add_checkbutton(label="Formulas", accelerator="Ctrl+`", variable=app.formulas_var,
                         command=lambda: wt.toggle_formulas(app))
    view.add_checkbutton(label="Full screen", accelerator="F11", variable=app.full_var,
                         command=lambda: wt.toggle_full_screen(app))
    view.add_command(label="Split window", command=lambda: app.split_window())
    view.add_checkbutton(label="Page breaks", variable=app.page_breaks_var,
                         command=lambda: __import__("print_preview").toggle(app, app.page_breaks_var.get()))
    view.add_separator()
    view.add_checkbutton(label="Dark mode", variable=app.dark_var, command=app.toggle_dark)
    view.add_checkbutton(label="Match system light / dark", variable=app.follow_var,
                         command=app.toggle_follow_system)
    view.add_checkbutton(label="High contrast", variable=app.contrast_var, command=app.toggle_contrast)
    menu.add_cascade(label="View", menu=view)
    insert = tk.Menu(menu, tearoff=False)
    insert.add_command(label="Chart…", command=lambda: app.insert_chart())
    insert.add_command(label="Picture…", command=lambda: __import__("images").insert(app))
    insert.add_command(label="Link…", accelerator="Ctrl+K", command=app.edit_link)
    insert.add_command(label="Comment…", command=app.edit_comment)
    menu.add_cascade(label="Insert", menu=insert)
    data = tk.Menu(menu, tearoff=False)
    data.add_command(label="Text to columns…", command=lambda: app.data_tool("text_to_columns"))
    data.add_command(label="Remove duplicates…", command=lambda: app.data_tool("remove_duplicates"))
    data.add_command(label="Flash fill", accelerator="Ctrl+E", command=lambda: app.data_tool("flash_fill"))
    data.add_command(label="Advanced filter…", command=lambda: app.data_tool("advanced"))
    data.add_command(label="Import text / CSV…", command=lambda: __import__("import_data").import_csv(app))
    data.add_command(label="Import table from web…", command=lambda: __import__("import_data").import_web(app))
    data.add_command(label="Run command on selection…",
                     command=lambda: __import__("scripting").run_command_dialog(app))
    data.add_separator()
    data.add_command(label="Pivot table…", command=lambda: app.data_tool("pivot"))
    data.add_command(label="Refresh pivot tables", command=lambda: __import__("pivots").refresh(app))
    data.add_command(label="Goal seek…", command=lambda: app.data_tool("goal_seek"))
    data.add_command(label="Scenario manager…", command=lambda: __import__("whatif").scenario_manager(app))
    data.add_command(label="Data table…", command=lambda: __import__("whatif").data_table_dialog(app))
    data.add_command(label="Solver…", command=lambda: __import__("whatif").solver(app))
    data.add_separator()
    data.add_command(label="Data validation…", command=lambda: app.rules_dialog("validation"))
    data.add_command(label="Conditional formatting…", command=lambda: app.rules_dialog("conditional"))
    data.add_separator()
    data.add_command(label="Group", accelerator="Alt+Shift+→", command=lambda: app.group(True))
    data.add_command(label="Ungroup", accelerator="Alt+Shift+←", command=lambda: app.group(False))
    data.add_separator()
    data.add_command(label="Update links", command=lambda: app.update_links())
    menu.add_cascade(label="Data", menu=data)
    sheet = tk.Menu(menu, tearoff=False, postcommand=lambda: sync_sheet(app, sheet))
    menu.add_cascade(label="Sheet", menu=sheet)
    tools = tk.Menu(menu, tearoff=False)
    tools.add_command(label="Spelling…", accelerator="F7", command=lambda: app.spell_check())
    tools.add_command(label="Recalculate", accelerator="F9", command=lambda: wt.recalculate(app))
    tools.add_command(label="Insert function…", accelerator="Shift+F3",
                      command=lambda: __import__("function_help").InsertFunction(app))
    tools.add_command(label="Evaluate formula…", command=lambda: __import__("function_help").EvaluateDialog(app))
    tools.add_command(label="Compare with another workbook…", command=lambda: __import__("compare").open_compare(app))
    tools.add_command(label="Script…", command=lambda: __import__("scripting").open_window(app))
    tools.add_command(label="AI assistant…", accelerator="Ctrl+Shift+A",
                      command=lambda: __import__("ai_assist").open_assistant(app))
    tools.add_command(label="Command palette…", accelerator="Ctrl+Shift+P",
                      command=lambda: __import__("palette").open_palette(app))
    tools.add_command(label="Calculation…", command=lambda: wt.CalculationDialog(app))
    tools.add_command(label="Properties…", command=lambda: wt.PropertiesDialog(app))
    tools.add_command(label="Password…", command=lambda: app.set_password())
    tools.add_command(label="Settings…", command=lambda: app.settings_dialog())
    menu.add_cascade(label="Tools", menu=tools)
    menu.add_command(label="New from template…", command=lambda: app.new_from_template())
    menu.add_command(label="Export as HTML…", command=lambda: app.export_html())
    help_ = tk.Menu(menu, tearoff=False)
    help_.add_command(label="Keyboard shortcuts", accelerator="F1", command=lambda: app.show_help("shortcuts"))
    help_.add_command(label="User guide", command=lambda: app.show_help("guide"))
    help_.add_command(label="Script guide", command=lambda: app.show_help("scripts"))
    help_.add_command(label="Tips", command=lambda: __import__("palette").show_tips(app))
    help_.add_command(label="Check for updates…", command=lambda: app.check_updates())
    help_.add_command(label="About", command=lambda: app.show_help("about"))
    menu.add_cascade(label="Help", menu=help_)
    menu.add_separator()
    menu.add_command(label="Exit", command=app.on_close)
    app.view_menu, app.sheet_menu = view, sheet
    import i18n
    i18n.translate_menu(menu)
    mb["menu"] = menu
    app.app_menu, app.recent_menu = menu, recent_menu

    from format_bar import Tooltip
    Tooltip(mb, "Menu")
    # save: blue disk when there are unsaved changes, grey (disabled) when everything is saved
    app.save_icons = save_images(app)
    app.save_btn = ttk.Button(parent, image=(app.save_icons["on"], "disabled", app.save_icons["off"]),
                              style="Icon.Toolbutton", command=app.save, state="disabled")
    app.save_btn.pack(side="left")
    Tooltip(app.save_btn, "Save (Ctrl+S)")
    app.undo_btn = icons.icon(ttk.Button(parent, width=2, command=app.undo, state="disabled"), "undo")
    app.undo_btn.pack(side="left")
    Tooltip(app.undo_btn, "Undo (Ctrl+Z)")
    app.redo_btn = icons.icon(ttk.Button(parent, width=2, command=app.redo, state="disabled"), "redo")
    app.redo_btn.pack(side="left")
    Tooltip(app.redo_btn, "Redo (Ctrl+Y)")
    ttk.Separator(parent, orient="vertical").pack(side="left", fill="y", padx=6, pady=(0, 4))


def sync_view(app):
    s = app.current
    app.grid_var.set(s.props.get("grid") is not False)
    app.headings_var.set(s.props.get("headings") is not False)


def sync_sheet(app, m):
    """Sheet submenu, rebuilt when opened (protect / unprotect depends on the sheet)."""
    import workbook_tools as wt
    m.delete(0, "end")
    name = app.sheet
    m.add_command(label="Duplicate", command=lambda: wt.duplicate_sheet(app, name))
    m.add_command(label="Copy to a new workbook…", command=lambda: wt.copy_to_new_workbook(app, name))
    m.add_command(label="Tab colour…", command=lambda: wt.pick_tab_color(app, name))
    m.add_separator()
    if app.current.props.get("protect"):
        m.add_command(label="Unprotect sheet…", command=lambda: wt.unprotect_sheet(app))
    else:
        m.add_command(label="Protect sheet…", command=lambda: wt.protect_sheet(app))
    m.add_command(label="Lock / unlock cells", command=lambda: wt.toggle_locked(app))
    if app.book_props.get("structure") is not None:            # sheets can't be added / deleted / renamed…
        m.add_command(label="Unprotect workbook…", command=lambda: wt.unprotect_book(app))
    else:
        m.add_command(label="Protect workbook…", command=lambda: wt.protect_book(app))
    m.add_separator()
    import print_preview
    m.add_command(label="Insert page break", command=lambda: print_preview.insert_break(app))
    m.add_command(label="Remove page breaks", command=lambda: print_preview.remove_breaks(app))
    import i18n
    i18n.translate_menu(m)


def fill_recent(app, menu):
    """Rebuilt each time the submenu opens."""
    menu.delete(0, "end")
    items = recent.load()
    if not items:
        menu.add_command(label="(none yet)", state="disabled")
        return
    for i, path in enumerate(items):
        label = f"{i + 1}  {os.path.basename(path)}   —   {os.path.dirname(path)}"
        menu.add_command(label=label, underline=0, command=lambda p=path: app.open_recent(p))
    menu.add_separator()
    menu.add_command(label="Clear list", command=recent.clear)
