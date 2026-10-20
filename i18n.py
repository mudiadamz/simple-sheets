"""Menu translations. The language is chosen in Tools → Settings (takes effect on the next start).
translate_menu() relabels a built tk.Menu (and its submenus) in place, so menu code stays in English."""
import version  # noqa: F401  (first: this module reads the settings folder when it loads)

# isort: split
import re

from simplecore import settings

LANGS = {"en": "English", "id": "Bahasa Indonesia", "es": "Español"}

T = {
    "id": {
        "New": "Baru", "Open…": "Buka…", "Open recent": "Buka terbaru", "Save": "Simpan", "Save as…": "Simpan sebagai…",
        "Export to PDF…": "Ekspor ke PDF…", "Print…": "Cetak…", "Find…": "Cari…", "Replace…": "Ganti…",
        "Named ranges…": "Rentang bernama…", "Undo": "Batalkan", "Redo": "Ulangi", "View": "Tampilan",
        "Gridlines": "Garis kisi", "Headings": "Judul baris/kolom", "Formulas": "Rumus", "Full screen": "Layar penuh",
        "Split window": "Bagi jendela", "Dark mode": "Mode gelap",
        "Match system light / dark": "Ikuti terang / gelap sistem", "Insert": "Sisipkan", "Chart…": "Bagan…",
        "Link…": "Tautan…", "Comment…": "Komentar…", "Data": "Data", "Text to columns…": "Teks ke kolom…",
        "Remove duplicates…": "Hapus duplikat…", "Flash fill": "Isi kilat", "Pivot table…": "Tabel pivot…",
        "Goal seek…": "Cari tujuan…", "Data validation…": "Validasi data…",
        "Conditional formatting…": "Format bersyarat…", "Group": "Kelompokkan", "Ungroup": "Pisahkan kelompok",
        "Update links": "Perbarui tautan", "Sheet": "Lembar", "Tools": "Alat", "Spelling…": "Ejaan…",
        "Recalculate": "Hitung ulang", "Calculation…": "Perhitungan…", "Properties…": "Properti…",
        "Settings…": "Pengaturan…", "New from template…": "Baru dari templat…", "Export as HTML…": "Ekspor sebagai HTML…",
        "Help": "Bantuan", "Keyboard shortcuts": "Pintasan keyboard", "User guide": "Panduan pengguna",
        "Check for updates…": "Periksa pembaruan…", "About": "Tentang", "Exit": "Keluar",
        "Duplicate": "Duplikat", "Copy to a new workbook…": "Salin ke buku kerja baru…", "Tab colour…": "Warna tab…",
        "Tab colour": "Warna tab", "Protect sheet…": "Lindungi lembar…", "Unprotect sheet…": "Buka perlindungan lembar…",
        "Lock / unlock cells": "Kunci / buka sel", "Add sheet": "Tambah lembar", "Rename": "Ganti nama",
        "Protect workbook…": "Lindungi buku kerja…", "Unprotect workbook…": "Buka perlindungan buku kerja…",
        "Shrink to fit (on / off)": "Perkecil agar muat (nyala / mati)",
        "Delete": "Hapus", "Hide": "Sembunyikan", "Unhide": "Tampilkan", "Other…": "Lainnya…", "No colour": "Tanpa warna",
        "Open link": "Buka tautan", "Edit link…": "Edit tautan…", "Insert link…": "Sisipkan tautan…", "Cut": "Potong",
        "Copy": "Salin", "Paste": "Tempel", "Values only": "Hanya nilai", "Format only": "Hanya format",
        "Paste special": "Tempel khusus", "Shift right": "Geser ke kanan", "Shift down": "Geser ke bawah",
        "Insert cells": "Sisipkan sel", "Shift left": "Geser ke kiri", "Shift up": "Geser ke atas",
        "Delete cells": "Hapus sel", "Edit comment": "Edit komentar", "Insert comment": "Sisipkan komentar",
        "Delete comment": "Hapus komentar", "Sort A → Z": "Urutkan A → Z", "Sort Z → A": "Urutkan Z → A",
        "Sort range…": "Urutkan rentang…", "Remove filter": "Hapus filter", "Create filter": "Buat filter",
        "Top row": "Baris atas", "First column": "Kolom pertama", "Unfreeze": "Lepas bekuan", "Freeze": "Bekukan",
        "Merge cells": "Gabung sel", "Merge & center": "Gabung & tengahkan", "Unmerge cells": "Pisahkan sel",
        "Unhide rows": "Tampilkan baris", "Unhide columns": "Tampilkan kolom", "Clear contents": "Hapus isi",
        "Delete chart": "Hapus bagan", "Clear list": "Kosongkan daftar", "(none yet)": "(belum ada)",
    },
    "es": {
        "New": "Nuevo", "Open…": "Abrir…", "Open recent": "Abrir reciente", "Save": "Guardar",
        "Save as…": "Guardar como…", "Export to PDF…": "Exportar a PDF…", "Print…": "Imprimir…", "Find…": "Buscar…",
        "Replace…": "Reemplazar…", "Named ranges…": "Rangos con nombre…", "Undo": "Deshacer", "Redo": "Rehacer",
        "View": "Vista", "Gridlines": "Líneas de cuadrícula", "Headings": "Encabezados", "Formulas": "Fórmulas",
        "Full screen": "Pantalla completa", "Split window": "Dividir ventana", "Dark mode": "Modo oscuro",
        "Match system light / dark": "Seguir claro / oscuro del sistema",
        "Insert": "Insertar", "Chart…": "Gráfico…", "Link…": "Vínculo…", "Comment…": "Comentario…", "Data": "Datos",
        "Text to columns…": "Texto en columnas…", "Remove duplicates…": "Quitar duplicados…",
        "Flash fill": "Relleno rápido", "Pivot table…": "Tabla dinámica…", "Goal seek…": "Buscar objetivo…",
        "Data validation…": "Validación de datos…", "Conditional formatting…": "Formato condicional…",
        "Group": "Agrupar", "Ungroup": "Desagrupar", "Update links": "Actualizar vínculos", "Sheet": "Hoja",
        "Tools": "Herramientas", "Spelling…": "Ortografía…", "Recalculate": "Recalcular",
        "Calculation…": "Cálculo…", "Properties…": "Propiedades…", "Settings…": "Configuración…",
        "New from template…": "Nuevo desde plantilla…", "Export as HTML…": "Exportar como HTML…", "Help": "Ayuda",
        "Keyboard shortcuts": "Atajos de teclado", "User guide": "Guía del usuario",
        "Check for updates…": "Buscar actualizaciones…", "About": "Acerca de", "Exit": "Salir",
        "Duplicate": "Duplicar", "Copy to a new workbook…": "Copiar a un libro nuevo…", "Tab colour…": "Color de pestaña…",
        "Tab colour": "Color de pestaña", "Protect sheet…": "Proteger hoja…", "Unprotect sheet…": "Desproteger hoja…",
        "Lock / unlock cells": "Bloquear / desbloquear celdas", "Add sheet": "Agregar hoja", "Rename": "Cambiar nombre",
        "Protect workbook…": "Proteger libro…", "Unprotect workbook…": "Desproteger libro…",
        "Shrink to fit (on / off)": "Reducir hasta ajustar (sí / no)",
        "Delete": "Eliminar", "Hide": "Ocultar", "Unhide": "Mostrar", "Other…": "Otro…", "No colour": "Sin color",
        "Open link": "Abrir vínculo", "Edit link…": "Editar vínculo…", "Insert link…": "Insertar vínculo…",
        "Cut": "Cortar", "Copy": "Copiar", "Paste": "Pegar", "Values only": "Solo valores", "Format only": "Solo formato",
        "Paste special": "Pegado especial", "Shift right": "Desplazar a la derecha", "Shift down": "Desplazar hacia abajo",
        "Insert cells": "Insertar celdas", "Shift left": "Desplazar a la izquierda", "Shift up": "Desplazar hacia arriba",
        "Delete cells": "Eliminar celdas", "Edit comment": "Editar comentario", "Insert comment": "Insertar comentario",
        "Delete comment": "Eliminar comentario", "Sort A → Z": "Ordenar A → Z", "Sort Z → A": "Ordenar Z → A",
        "Sort range…": "Ordenar rango…", "Remove filter": "Quitar filtro", "Create filter": "Crear filtro",
        "Top row": "Fila superior", "First column": "Primera columna", "Unfreeze": "Movilizar", "Freeze": "Inmovilizar",
        "Merge cells": "Combinar celdas", "Merge & center": "Combinar y centrar", "Unmerge cells": "Separar celdas",
        "Unhide rows": "Mostrar filas", "Unhide columns": "Mostrar columnas", "Clear contents": "Borrar contenido",
        "Delete chart": "Eliminar gráfico", "Clear list": "Borrar lista", "(none yet)": "(ninguno todavía)",
    },
}

PATTERNS = {   # labels with numbers or names inside
    "id": [(r"^Insert (\d+ )?rows? above$", lambda m: f"Sisipkan {m.group(1) or ''}baris di atas"),
           (r"^Insert (\d+ )?rows? below$", lambda m: f"Sisipkan {m.group(1) or ''}baris di bawah"),
           (r"^Insert (\d+ )?columns? left$", lambda m: f"Sisipkan {m.group(1) or ''}kolom di kiri"),
           (r"^Insert (\d+ )?columns? right$", lambda m: f"Sisipkan {m.group(1) or ''}kolom di kanan")],
    "es": [(r"^Insert (\d+ )?rows? above$", lambda m: f"Insertar {m.group(1) or ''}fila(s) arriba"),
           (r"^Insert (\d+ )?rows? below$", lambda m: f"Insertar {m.group(1) or ''}fila(s) debajo"),
           (r"^Insert (\d+ )?columns? left$", lambda m: f"Insertar {m.group(1) or ''}columna(s) a la izquierda"),
           (r"^Insert (\d+ )?columns? right$", lambda m: f"Insertar {m.group(1) or ''}columna(s) a la derecha")],
}

_lang = None


def _merge():
    """Dialog / tooltip / message texts, newer menu items, and your own translations: a file
    lang_<code>.json in the settings folder ({"_name": "Deutsch", "Save": "Speichern", ...}) adds or
    completes a language."""
    import json
    import os

    from simplecore.recent import config_dir

    import i18n_strings
    for table in (i18n_strings.UI, i18n_strings.MENUS):
        for code, words in table.items():
            for k, v in words.items():
                T.setdefault(code, {}).setdefault(k, v)
    try:
        names = os.listdir(config_dir())
    except OSError:
        names = []
    for name in names:
        m = re.fullmatch(r"lang_([A-Za-z-]{2,10})\.json", name)
        if not m:
            continue
        try:
            with open(os.path.join(config_dir(), name), encoding="utf-8") as f:
                words = json.load(f)
        except (OSError, ValueError):
            continue
        code = m.group(1)
        LANGS.setdefault(code, str(words.pop("_name", code)))
        T.setdefault(code, {}).update({str(k): str(v) for k, v in words.items()})


_merge()


def tr_text(text):
    """A message: whole, else line by line (fixed lines translate, the rest stays as it is)."""
    if lang() == "en" or not text:
        return text
    whole = tr(text)
    if whole != text:
        return whole
    return "\n".join(tr(line) for line in text.split("\n"))


TEXT_CLASSES = {"Label", "TLabel", "Button", "TButton", "Checkbutton", "TCheckbutton", "Radiobutton",
                "TRadiobutton", "Labelframe", "TLabelframe", "Menubutton", "TMenubutton"}


def translate_widgets(widget):
    """Relabel a window's widgets (labels, buttons, check boxes, frames, list headings, notebook tabs)."""
    if lang() == "en":
        return 0
    n = 0
    try:
        cls = widget.winfo_class()
    except Exception:
        return 0
    if cls in TEXT_CLASSES:
        try:
            text = widget.cget("text")
            new = tr_text(str(text)) if text else text
            if new != text:
                widget.configure(text=new)
                n += 1
        except Exception:
            pass
    elif cls == "Treeview":
        for col in ("#0",) + tuple(widget["columns"]):
            try:
                text = widget.heading(col, "text")
                if text and tr(text) != text:
                    widget.heading(col, text=tr(text))
                    n += 1
            except Exception:
                pass
    elif cls == "TNotebook":
        for tab in widget.tabs():
            text = widget.tab(tab, "text")
            if text and tr(text) != text:
                widget.tab(tab, text=tr(text))
                n += 1
    elif cls == "Toplevel":
        try:
            title = widget.title()
            if title and tr(title) != title:
                widget.title(tr(title))
                n += 1
        except Exception:
            pass
    for child in widget.winfo_children():
        n += translate_widgets(child)
    return n


def install(root):
    """Translate every new window when it appears, and the texts of message boxes / questions."""
    if lang() == "en":
        return False
    seen = set()

    def on_map(e):
        w = e.widget
        if getattr(w, "winfo_class", None) and w.winfo_class() == "Toplevel" and str(w) not in seen:
            seen.add(str(w))
            translate_widgets(w)
    root.bind_class("Toplevel", "<Map>", on_map, add="+")
    wrap_dialogs()
    return True


_originals = {}


def wrap_dialogs():
    """messagebox.* / simpledialog.ask*: title and message translated before they are shown."""
    from tkinter import messagebox, simpledialog
    for mod, names in ((messagebox, ("showinfo", "showwarning", "showerror", "askokcancel", "askyesno",
                                     "askyesnocancel", "askquestion", "askretrycancel")),
                       (simpledialog, ("askstring", "askinteger", "askfloat"))):
        for name in names:
            fn = getattr(mod, name)
            if getattr(fn, "_translated", False):
                continue
            _originals[(mod, name)] = fn

            def wrapped(title=None, prompt=None, *a, _fn=fn, **k):
                return _fn(tr_text(title) if isinstance(title, str) else title,
                           tr_text(prompt) if isinstance(prompt, str) else prompt, *a, **k)
            wrapped._translated = True
            setattr(mod, name, wrapped)


def unwrap_dialogs():
    for (mod, name), fn in _originals.items():
        setattr(mod, name, fn)
    _originals.clear()


def lang():
    global _lang
    if _lang is None:
        _lang = settings.get("lang", "en") if settings.get("lang", "en") in LANGS else "en"
    return _lang


def set_lang(code):
    global _lang
    _lang = code if code in LANGS else "en"


def tr(text):
    table = T.get(lang())
    if not table or not text:
        return text
    if text in table:
        return table[text]
    for pat, fn in PATTERNS.get(lang(), []):
        m = re.match(pat, text)
        if m:
            return fn(m)
    return text


def translate_menu(menu, seen=None):
    """Relabel every entry of a tk.Menu (recursively) in the current language."""
    if lang() == "en":
        return
    seen = seen if seen is not None else set()
    if str(menu) in seen:
        return
    seen.add(str(menu))
    end = menu.index("end")
    if end is None:
        return
    for i in range(end + 1):
        kind = menu.type(i)
        if kind in ("command", "cascade", "checkbutton", "radiobutton"):
            label = menu.entrycget(i, "label")
            new = tr(label)
            if new != label:
                menu.entryconfigure(i, label=new)
            if kind == "cascade":
                sub = menu.entrycget(i, "menu")
                if sub:
                    translate_menu(menu.nametowidget(sub), seen)


# simplecore's widgets (Tooltip…) translate through this too
from simplecore import appinfo as _appinfo  # noqa: E402

_appinfo.tr = tr
