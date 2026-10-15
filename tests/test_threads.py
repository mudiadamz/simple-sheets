"""Threaded comments: start a conversation, reply, saved as Excel threaded comments, read back."""
import re
import zipfile

import openpyxl

import comments
import threads
import workbooks
from conftest import open_in


def test_start_reply_save_and_read(app, tmp_path):
    path = workbooks.plain(tmp_path / "p.xlsx")
    g = open_in(app, path)
    assert app.add_reply(1, 1, "Is this right?", start=True)
    assert app.add_reply(1, 1, "Yes, checked.")
    t = g.sheet.threads[(1, 1)]
    assert [m[1] for m in t] == ["Is this right?", "Yes, checked."]
    note = g.sheet.comments[(1, 1)]
    assert note[1] == "tc=" + t[0][3] and "Reply:" in note[0]          # Excel's fallback note
    assert app.save()
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        tpart = next(n for n in names if n.startswith("xl/threadedComments/"))
        xml = z.read(tpart).decode()
        assert xml.count("<threadedComment ") == 2 and f'parentId="{t[0][3]}"' in xml
        assert "xl/persons/person.xml" in names
        assert "threadedcomments+xml" in z.read("[Content_Types].xml").decode()
        assert re.search(r"threadedComment", z.read("xl/worksheets/_rels/sheet1.xml.rels").decode())
    openpyxl.load_workbook(path)                                        # still a valid file
    found = threads.read(str(path))
    assert [m[1] for m in found["Data"][(1, 1)]] == ["Is this right?", "Yes, checked."]
    g = open_in(app, path)                                              # opened again: the conversation is back
    assert len(g.sheet.threads[(1, 1)]) == 2
    app.undo()                                                          # (nothing to undo after opening)
    app.add_reply(1, 1, "Third")
    app.undo()
    assert len(g.sheet.threads[(1, 1)]) == 2


def test_dialog(app):
    d = comments.CommentDialog(app, 0, 0)
    d.text.insert("1.0", "Let's talk")
    d.start.set(True)
    d.ok()
    assert app.current.threads[(0, 0)][0][1] == "Let's talk"
    d = comments.CommentDialog(app, 0, 0)
    assert "Let's talk" in d.conversation.get("1.0", "end")
    d.text.insert("1.0", "ok")
    d.ok()
    assert [m[1] for m in app.current.threads[(0, 0)]] == ["Let's talk", "ok"]
    d = comments.CommentDialog(app, 0, 0)
    d.delete_thread()
    assert (0, 0) not in app.current.threads and (0, 0) not in app.current.comments


def test_threads_follow_inserted_rows(app):
    app.add_reply(2, 0, "hello", start=True)
    app.restructure("row", 0, 1)
    assert (3, 0) in app.current.threads and (3, 0) in app.current.comments
