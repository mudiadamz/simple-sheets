"""Threaded comments (Excel 365 comments with replies). Read from the file, shown as a conversation in the
comment dialog and on hover, replies added here, and written back (openpyxl itself drops them; Excel's
plain-note copy of each thread is kept up to date too, for older programs).

Sheet.threads: {(r, c): [[author, text, time ISO, id], ...]} - the first message is the comment."""
import datetime
import os
import posixpath
import re
import uuid
import zipfile
from xml.sax.saxutils import escape, quoteattr

import keep_parts

NS = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"
SHEET_REL = "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment"
PERSON_REL = "http://schemas.microsoft.com/office/2017/10/relationships/person"
THREAD_CT = "application/vnd.ms-excel.threadedcomments+xml"
PERSON_CT = "application/vnd.ms-excel.person+xml"


def new_id():
    return "{" + str(uuid.uuid4()).upper() + "}"


def now():
    return datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S.00")


def _cell(ref):
    from formula import col_index
    m = re.fullmatch(r"([A-Z]+)(\d+)", ref)
    return int(m.group(2)) - 1, col_index(m.group(1))


def _attr(tag, name):
    m = re.search(rf'\b{name}="([^"]*)"', tag)
    from html import unescape
    return unescape(m.group(1)) if m else None


def read(path):
    """{sheet name: {(r, c): [[author, text, time, id], ...]}} from an .xlsx."""
    from html import unescape

    from sheet_io import _sheet_files
    try:
        with zipfile.ZipFile(path) as z:
            files = {n: z.read(n) for n in z.namelist() if n.endswith((".xml", ".rels"))}
    except (OSError, zipfile.BadZipFile):
        return {}
    people = {}
    for n, data in files.items():
        if n.startswith("xl/persons/"):
            for tag in re.findall(r"<person\b[^>]*>", data.decode("utf-8")):
                people[_attr(tag, "id")] = _attr(tag, "displayName") or ""
    out = {}
    try:
        where = _sheet_files(files)
    except KeyError:
        return {}
    for name, part in where.items():
        for target, kind in keep_parts._rels(files, part).values():
            if not kind.endswith("/threadedComment"):
                continue
            xml = files.get(keep_parts._resolve(part, target), b"").decode("utf-8")
            threads, by_id = {}, {}
            for m in re.finditer(r"<threadedComment\b([^>]*)>(.*?)</threadedComment>", xml, re.S):
                tag, body = m.group(1), m.group(2)
                t = re.search(r"<text>(.*?)</text>", body, re.S)
                text = unescape(t.group(1)) if t else ""
                msg = [people.get(_attr(tag, "personId"), ""), text, _attr(tag, "dT") or "", _attr(tag, "id") or new_id()]
                parent = _attr(tag, "parentId")
                cell = _cell(_attr(tag, "ref") or "A1")
                if parent and parent in by_id:
                    threads[by_id[parent]].append(msg)
                else:
                    threads[cell] = [msg]
                    by_id[msg[3]] = cell
            if threads:
                out[name] = threads
    return out


def fallback(thread):
    """The plain note Excel keeps next to a thread for programs that don't know threads."""
    lines = ["[Threaded comment]", "", "Your version of Excel allows you to read this threaded comment; however, any "
             "edits to it will get removed if the file is opened in a newer version of Excel.", "", "Comment:",
             "    " + thread[0][1]]
    for msg in thread[1:]:
        lines += ["Reply:", "    " + msg[1]]
    return "\n".join(lines)


def write(path, sheets):
    """Put every sheet's threads into the saved .xlsx (with the people who wrote them)."""
    from sheet_io import _sheet_files
    if not any(getattr(s, "threads", None) for s in sheets.values()):
        return 0
    with zipfile.ZipFile(path) as z:
        infos = z.infolist()
        files = {i.filename: z.read(i.filename) for i in infos}
    order = [i.filename for i in infos]
    where = _sheet_files(files)
    persons, count = {}, 0
    ct = files["[Content_Types].xml"].decode("utf-8")
    for name, s in sheets.items():
        threads = getattr(s, "threads", None)
        part = where.get(name)
        if not threads or not part:
            continue
        items = []
        from formula import col_letters
        for (r, c), msgs in sorted(threads.items()):
            ref = f"{col_letters(c)}{r + 1}"
            first = msgs[0][3]
            for k, (author, text, when, mid) in enumerate(msgs):
                pid = persons.setdefault(author or "Unknown", new_id())
                parent = f" parentId={quoteattr(first)}" if k else ""
                items.append(f'<threadedComment ref="{ref}" dT={quoteattr(when or now())} personId={quoteattr(pid)} '
                             f"id={quoteattr(mid)}{parent}><text>{escape(text)}</text></threadedComment>")
        n = 1
        while f"xl/threadedComments/threadedComment{n}.xml" in files:
            n += 1
        tpart = f"xl/threadedComments/threadedComment{n}.xml"
        files[tpart] = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<ThreadedComments xmlns="{NS}" '
                        'xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                        + "".join(items) + "</ThreadedComments>").encode("utf-8")
        order.append(tpart)
        ct = ct.replace("</Types>", f'<Override PartName="/{tpart}" ContentType="{THREAD_CT}"/></Types>')
        keep_parts._add_rel(files, part, SHEET_REL, posixpath.relpath(tpart, posixpath.dirname(part)))
        if keep_parts._rels_path(part) not in order:
            order.append(keep_parts._rels_path(part))
        count += len(items)
    if persons:
        ppart = "xl/persons/person.xml"
        files[ppart] = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<personList xmlns="{NS}" '
                        'xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                        + "".join(f"<person displayName={quoteattr(a)} id={quoteattr(i)} userId={quoteattr(a)} "
                                  'providerId="None"/>' for a, i in persons.items())
                        + "</personList>").encode("utf-8")
        if ppart not in order:
            order.append(ppart)
            ct = ct.replace("</Types>", f'<Override PartName="/{ppart}" ContentType="{PERSON_CT}"/></Types>')
            keep_parts._add_rel(files, "xl/workbook.xml", PERSON_REL, "persons/person.xml")
    files["[Content_Types].xml"] = ct.encode("utf-8")
    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for n in order:
            if n in files:
                z.writestr(n, files[n])
    os.replace(tmp, path)
    return count
