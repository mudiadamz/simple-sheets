"""Sign in with Google, then open and save Google Sheets / Drive spreadsheets (☰ → Google Drive).

Sign-in is OAuth for desktop apps: the browser shows Google's own sign-in page and sends the answer back to a
one-time local address (127.0.0.1); PKCE protects it. The user's own Google Cloud "Desktop app" OAuth client
(Client ID + secret, entered in Google account…) is used; the secret and the sign-in (refresh token) are kept
encrypted for the Windows account (DPAPI). Files: a Google Sheet opens as .xlsx (exported), Excel / CSV /
ODS files download as they are; Save puts the file back (a Google Sheet is converted again), Save to Google
Drive… uploads a new copy. Only the standard library is used."""
import base64
import hashlib
import json
import os
import queue
import secrets
import sys
import threading
import time
import urllib.error
import urllib.parse
import urllib.request

import settings

AUTH_URL = "https://accounts.google.com/o/oauth2/v2/auth"
TOKEN_URL = "https://oauth2.googleapis.com/token"
REVOKE_URL = "https://oauth2.googleapis.com/revoke"
DRIVE = "https://www.googleapis.com/drive/v3"
UPLOAD = "https://www.googleapis.com/upload/drive/v3"
SCOPES = "openid email https://www.googleapis.com/auth/drive"
CONSOLE = "https://console.cloud.google.com/apis/credentials"

SHEET = "application/vnd.google-apps.spreadsheet"
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
MIMES = {".xlsx": XLSX, ".xlsm": "application/vnd.ms-excel.sheet.macroEnabled.12", ".xls": "application/vnd.ms-excel",
         ".csv": "text/csv", ".tsv": "text/tab-separated-values", ".ods": "application/vnd.oasis.opendocument.spreadsheet"}
OPENABLE = {SHEET, *MIMES.values()}


class GoogleError(Exception):
    pass


# ---------- HTTP (tests swap `transport`) ----------

def request(method, url, token=None, data=None, headers=None, timeout=60):
    """-> (status, body bytes). Raises GoogleError with Google's message on errors."""
    h = dict(headers or {})
    if token:
        h["Authorization"] = f"Bearer {token}"
    req = urllib.request.Request(url, data=data, headers=h, method=method)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return r.status, r.read()
    except urllib.error.HTTPError as e:
        detail = ""
        try:
            err = json.loads(e.read().decode("utf-8"))
            inner = err.get("error")
            detail = err.get("error_description") or (inner.get("message") if isinstance(inner, dict) else inner)
        except Exception:
            pass
        hint = {401: "Signed out by Google - sign in again (☰ → Google Drive → Google account…).",
                403: "Google refused: check that the Drive API is enabled for your Google Cloud project.",
                404: "Not found on Google Drive."}.get(e.code, "")
        raise GoogleError(f"Google said {e.code}. {hint} {detail or ''}".strip()) from None
    except (urllib.error.URLError, TimeoutError, OSError) as e:
        raise GoogleError(f"Couldn't reach Google: {getattr(e, 'reason', e)}") from None


transport = request


def call_json(method, url, token=None, form=None, body=None, headers=None):
    data, h = None, dict(headers or {})
    if form is not None:
        data = urllib.parse.urlencode(form).encode()
        h["Content-Type"] = "application/x-www-form-urlencoded"
    elif body is not None:
        data = json.dumps(body).encode()
        h["Content-Type"] = "application/json"
    _status, raw = transport(method, url, token, data, h)
    return json.loads(raw.decode("utf-8") or "{}") if raw else {}


# ---------- secrets kept for the Windows account ----------

def _store(name, value):
    value = (value or "").strip()
    if not value:
        settings.set(name, "")
        return
    if sys.platform == "win32":
        try:
            from ai_assist import _dpapi
            settings.set(name, "dpapi:" + base64.b64encode(_dpapi(value.encode(), True)).decode())
            return
        except OSError:
            pass
    settings.set(name, "plain:" + value)


def _load(name):
    v = settings.get(name, "") or ""
    if v.startswith("dpapi:"):
        try:
            from ai_assist import _dpapi
            return _dpapi(base64.b64decode(v[6:]), False).decode()
        except (OSError, ValueError):
            return ""
    return v[6:] if v.startswith("plain:") else v


def client():
    return (settings.get("google_client_id") or "").strip(), _load("google_client_secret")


def set_client(client_id, secret):
    settings.set("google_client_id", (client_id or "").strip())
    _store("google_client_secret", secret)


def email():
    return settings.get("google_email") or ""


def signed_in():
    return bool(_load("google_refresh"))


# ---------- sign in / out ----------

_access = {"token": None, "until": 0.0}


def _email_of(id_token):
    try:
        payload = id_token.split(".")[1]
        return json.loads(base64.urlsafe_b64decode(payload + "=" * (-len(payload) % 4))).get("email", "")
    except Exception:
        return ""


def _keep(tokens):
    if tokens.get("refresh_token"):
        _store("google_refresh", tokens["refresh_token"])
    if tokens.get("id_token"):
        settings.set("google_email", _email_of(tokens["id_token"]))
    if tokens.get("access_token"):
        _access["token"] = tokens["access_token"]
        _access["until"] = time.time() + int(tokens.get("expires_in") or 3600) - 60


def sign_in(open_browser=None, timeout=300, cancelled=None):
    """Google's sign-in page in the browser; the answer comes back to a local address. -> the email.
    cancelled() -> True stops waiting (the Cancel button). Raises GoogleError."""
    import http.server
    cid, secret = client()
    if not cid:
        raise GoogleError("Enter your Google OAuth Client ID first (☰ → Google Drive → Google account…).")
    verifier = secrets.token_urlsafe(64)
    challenge = base64.urlsafe_b64encode(hashlib.sha256(verifier.encode()).digest()).rstrip(b"=").decode()
    state = secrets.token_urlsafe(24)
    got, done = {}, threading.Event()

    class Answer(http.server.BaseHTTPRequestHandler):
        def do_GET(self):
            q = dict(urllib.parse.parse_qsl(urllib.parse.urlsplit(self.path).query))
            if "code" not in q and "error" not in q:          # favicon and the like
                self.send_response(404)
                self.end_headers()
                return
            got.update(q)
            ok = "code" in q and q.get("state") == state
            page = ("<h2>Signed in.</h2><p>You can close this tab and go back to Simple Sheets.</p>" if ok else
                    "<h2>Sign-in didn't finish.</h2><p>Go back to Simple Sheets and try again.</p>")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(f"<html><body style='font-family:sans-serif;margin:3em'>{page}</body></html>".encode())
            done.set()

        def log_message(self, *a):
            pass

    server = http.server.HTTPServer(("127.0.0.1", 0), Answer)
    redirect = f"http://127.0.0.1:{server.server_port}"
    threading.Thread(target=server.serve_forever, daemon=True).start()
    url = AUTH_URL + "?" + urllib.parse.urlencode({
        "client_id": cid, "redirect_uri": redirect, "response_type": "code", "scope": SCOPES,
        "code_challenge": challenge, "code_challenge_method": "S256", "state": state,
        "access_type": "offline", "prompt": "consent"})
    try:
        (open_browser or __import__("webbrowser").open)(url)
        end = time.time() + timeout
        while not done.wait(0.2):
            if (cancelled and cancelled()) or time.time() > end:
                raise GoogleError("Sign-in was cancelled." if cancelled and cancelled() else "Sign-in timed out.")
    finally:
        server.shutdown()
        server.server_close()
    if got.get("state") != state:
        raise GoogleError("Sign-in answer didn't match (state) - try again.")
    if "error" in got:
        raise GoogleError(f"Google sign-in: {got['error']}")
    tokens = call_json("POST", TOKEN_URL, form={
        "code": got["code"], "client_id": cid, "client_secret": secret, "redirect_uri": redirect,
        "grant_type": "authorization_code", "code_verifier": verifier})
    if not tokens.get("access_token"):
        raise GoogleError("Google didn't send a sign-in token.")
    _keep(tokens)
    return email()


def token():
    """A current access token (refreshed when needed). Raises GoogleError when signed out."""
    if _access["token"] and time.time() < _access["until"]:
        return _access["token"]
    refresh = _load("google_refresh")
    if not refresh:
        raise GoogleError("Sign in with Google first (☰ → Google Drive → Google account…).")
    cid, secret = client()
    tokens = call_json("POST", TOKEN_URL, form={"client_id": cid, "client_secret": secret,
                                                "refresh_token": refresh, "grant_type": "refresh_token"})
    if not tokens.get("access_token"):
        raise GoogleError("Google didn't renew the sign-in - sign in again.")
    _keep(tokens)
    return _access["token"]


def sign_out():
    refresh = _load("google_refresh")
    if refresh:
        try:                                     # tell Google too (best effort)
            transport("POST", REVOKE_URL + "?" + urllib.parse.urlencode({"token": refresh}), None, b"",
                      {"Content-Type": "application/x-www-form-urlencoded"})
        except GoogleError:
            pass
    _store("google_refresh", "")
    settings.set("google_email", "")
    _access.update(token=None, until=0.0)


# ---------- Drive ----------

def list_files(search=""):
    """Spreadsheets on Drive, newest first: [{"id", "name", "mimeType", "modifiedTime"}]."""
    kinds = " or ".join(f"mimeType='{m}'" for m in sorted(OPENABLE))
    q = f"({kinds}) and trashed=false"
    if search.strip():
        q += " and name contains '" + search.strip().replace("\\", "\\\\").replace("'", "\\'") + "'"
    url = DRIVE + "/files?" + urllib.parse.urlencode({
        "q": q, "orderBy": "modifiedTime desc", "pageSize": 100,
        "fields": "files(id,name,mimeType,modifiedTime)", "supportsAllDrives": "true",
        "includeItemsFromAllDrives": "true"})
    return call_json("GET", url, token()).get("files", [])


def cache_dir(file_id):
    from recent import config_dir
    d = os.path.join(config_dir(), "drive", "".join(ch for ch in file_id if ch.isalnum() or ch in "-_"))
    os.makedirs(d, exist_ok=True)
    return d


def local_name(f):
    """The file name it gets on this computer (a Google Sheet becomes .xlsx)."""
    name = "".join("_" if ch in '<>:"/\\|?*' else ch for ch in f["name"]).strip() or "Untitled"
    if f["mimeType"] == SHEET:
        return name if name.lower().endswith(".xlsx") else name + ".xlsx"
    ext = next((e for e, m in MIMES.items() if m == f["mimeType"]), "")
    return name if not ext or name.lower().endswith(ext) else name + ext


def download(f):
    """Download a Drive file into the cache -> local path."""
    if f["mimeType"] == SHEET:
        url = f"{DRIVE}/files/{f['id']}/export?" + urllib.parse.urlencode({"mimeType": XLSX})
    else:
        url = f"{DRIVE}/files/{f['id']}?alt=media&supportsAllDrives=true"
    _status, data = transport("GET", url, token(), None, {})
    path = os.path.join(cache_dir(f["id"]), local_name(f))
    with open(path, "wb") as out:
        out.write(data)
    return path


def upload(f, path):
    """Put a saved file back on Drive (a Google Sheet is converted from the .xlsx again)."""
    mime = MIMES.get(os.path.splitext(path)[1].lower(), "application/octet-stream")
    with open(path, "rb") as fh:
        data = fh.read()
    url = f"{UPLOAD}/files/{f['id']}?uploadType=media&supportsAllDrives=true"
    _status, raw = transport("PATCH", url, token(), data, {"Content-Type": mime})
    return json.loads(raw.decode("utf-8") or "{}") if raw else {}


def create(path, name, as_sheet=True):
    """Upload a new file to Drive (as a Google Sheet, or as it is) -> its file dict."""
    mime = MIMES.get(os.path.splitext(path)[1].lower(), "application/octet-stream")
    meta = {"name": name, "mimeType": SHEET if as_sheet else mime}
    with open(path, "rb") as fh:
        data = fh.read()
    boundary = "ss" + secrets.token_hex(12)
    body = (f"--{boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n{json.dumps(meta)}\r\n"
            f"--{boundary}\r\nContent-Type: {mime}\r\n\r\n").encode() + data + f"\r\n--{boundary}--\r\n".encode()
    url = f"{UPLOAD}/files?uploadType=multipart&fields=id,name,mimeType,modifiedTime&supportsAllDrives=true"
    _status, raw = transport("POST", url, token(), body, {"Content-Type": f"multipart/related; boundary={boundary}"})
    return json.loads(raw.decode("utf-8"))


# ---------- in the app ----------

def in_background(app, work, done):
    """Run work() on a thread; done(result, error text or None) runs on the main thread after."""
    q = queue.Queue()

    def run():
        try:
            q.put((work(), None))
        except GoogleError as e:
            q.put((None, str(e)))
        except Exception as e:                       # never leave the window waiting
            q.put((None, f"{type(e).__name__}: {e}"))
    threading.Thread(target=run, daemon=True).start()

    def check():
        try:
            result, error = q.get_nowait()
        except queue.Empty:
            app._root().after(100, check)
            return
        done(result, error)
    app._root().after(100, check)


def open_drive_file(app, f, path):
    """Open a downloaded Drive file; Save then puts it back on Drive."""
    app.open_file(path)
    if app.path and os.path.abspath(app.path) == os.path.abspath(path):
        app.drive_file = dict(f, path=os.path.abspath(path))
        app.status["text"] = f"Opened {f['name']} from Google Drive"
        app.refresh()
    return path


def after_save(app, dst):
    """Called after every save: a file from Drive goes back there (in the background)."""
    f = getattr(app, "drive_file", None)
    if not f or os.path.abspath(dst) != f.get("path"):
        return None
    app.status["text"] = f"Saving {f['name']} to Google Drive…"

    def done(_result, error):
        if error:
            from tkinter import messagebox
            messagebox.showerror("Google Drive", f"Saved on this computer, but not on Google Drive:\n\n{error}",
                                 parent=app)
            app.status["text"] = "Not saved to Google Drive"
        else:
            app.status["text"] = f"Saved {f['name']} to Google Drive"
    in_background(app, lambda: upload(f, dst), done)
    return f


def need_sign_in(app):
    """True when signed in; else offers the Google account window."""
    if signed_in():
        return True
    from tkinter import messagebox
    if messagebox.askyesno("Google Drive", "Sign in with Google first?", parent=app):
        AccountDialog(app)
    return False


def save_to_drive(app):
    """☰ → Google Drive → Save to Google Drive…: a new copy on Drive (as a Google Sheet or as the file)."""
    if not need_sign_in(app):
        return None
    if not app.path or app.dirty():
        if not app.save():
            return None
    d = SaveDialog(app)
    return d


class AccountDialog:
    """☰ → Google Drive → Google account…: your OAuth client (ID + secret) and Sign in / Sign out."""

    def __init__(self, app):
        import tkinter as tk
        from tkinter import ttk

        from theme import C
        self.app = app
        w = self.win = tk.Toplevel(app)
        w.withdraw()
        w.title("Google account")
        w.transient(app)
        w.resizable(False, False)
        f = ttk.Frame(w, padding=(18, 14))
        f.pack(fill="both", expand=True)
        f.columnconfigure(1, weight=1)
        wrap = int(440 * getattr(app, "scale", 1))
        ttk.Label(f, text="Sign in with Google", font=("Segoe UI", 11, "bold")).grid(row=0, column=0, columnspan=2,
                                                                                    sticky="w")
        self.state = ttk.Label(f, text="", wraplength=wrap)
        self.state.grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 8))
        self.btn = ttk.Button(f, text="Sign in with Google", command=self.toggle)
        self.btn.grid(row=2, column=0, columnspan=2, sticky="w")
        ttk.Separator(f).grid(row=3, column=0, columnspan=2, sticky="ew", pady=12)
        ttk.Label(f, text="Your Google OAuth client", font=("Segoe UI", 10, "bold")).grid(row=4, column=0,
                                                                                        columnspan=2, sticky="w")
        steps = tk.Label(f, justify="left", anchor="w", wraplength=wrap, background=C["bg"], foreground=C["muted"],
                         font=("Segoe UI", 8), text=(
                             "1. Google Cloud console → APIs & Services: enable the Google Drive API.\n"
                             "2. OAuth consent screen: add yourself as a test user.\n"
                             "3. Credentials → Create credentials → OAuth client ID → type \"Desktop app\".\n"
                             "4. Paste its Client ID and Client secret here."))
        steps.grid(row=5, column=0, columnspan=2, sticky="w", pady=(2, 4))
        link = tk.Label(f, text="Open the Google Cloud console", cursor="hand2", background=C["bg"],
                        foreground=C["accent"], font=("Segoe UI", 9, "underline"))
        link.grid(row=6, column=0, columnspan=2, sticky="w", pady=(0, 8))
        link.bind("<Button-1>", lambda e: __import__("webbrowser").open(CONSOLE))
        cid, secret = client()
        self.cid = tk.StringVar(master=w, value=cid)
        self.secret = tk.StringVar(master=w, value=secret)
        ttk.Label(f, text="Client ID").grid(row=7, column=0, sticky="w", padx=(0, 8))
        ttk.Entry(f, textvariable=self.cid, width=52).grid(row=7, column=1, sticky="ew", pady=2)
        ttk.Label(f, text="Client secret").grid(row=8, column=0, sticky="w", padx=(0, 8))
        ttk.Entry(f, textvariable=self.secret, show="•").grid(row=8, column=1, sticky="ew", pady=2)
        tk.Label(f, text="Kept on this computer only (the secret and your sign-in are encrypted for your Windows "
                         "account).", background=C["bg"], foreground=C["muted"], font=("Segoe UI", 8),
                 wraplength=wrap, justify="left").grid(row=9, column=0, columnspan=2, sticky="w", pady=(4, 0))
        ttk.Separator(f).grid(row=10, column=0, columnspan=2, sticky="ew", pady=12)
        b = ttk.Frame(f)
        b.grid(row=11, column=0, columnspan=2, sticky="e")
        ttk.Button(b, text="Close", command=self.close).pack(side="right")
        w.bind("<Escape>", lambda e: self.close())
        self.waiting = False
        self.show_state()
        from ai_assist import center
        center(w, app)
        w.deiconify()

    def save_client(self):
        set_client(self.cid.get(), self.secret.get())

    def show_state(self):
        if self.waiting:
            self.state.configure(text="Waiting for you to sign in in the browser…")
            self.btn.configure(text="Cancel")
        elif signed_in():
            self.state.configure(text=f"Signed in as {email() or 'your Google account'}. "
                                      "☰ → Google Drive opens and saves your Drive spreadsheets.")
            self.btn.configure(text="Sign out")
        else:
            self.state.configure(text="Not signed in. Sign in to open and save Google Sheets and Drive files.")
            self.btn.configure(text="Sign in with Google")

    def toggle(self):
        if self.waiting:
            self.waiting = False                      # Cancel: sign_in stops waiting
            return self.show_state()
        if signed_in():
            sign_out()
            return self.show_state()
        self.save_client()
        self.waiting = True
        self.show_state()

        def done(who, error):
            self.waiting = False
            if not self.win.winfo_exists():
                return
            self.show_state()
            if error:
                from tkinter import messagebox
                messagebox.showerror("Google sign-in", error, parent=self.win)
        in_background(self.app, lambda: sign_in(cancelled=lambda: not self.waiting), done)

    def close(self):
        self.save_client()
        self.waiting = False
        self.win.destroy()


class DrivePicker:
    """☰ → Google Drive → Open from Google Drive…: your spreadsheets, newest first; type to search."""

    def __init__(self, app):
        import tkinter as tk
        from tkinter import ttk
        self.app = app
        w = self.win = tk.Toplevel(app)
        w.withdraw()
        w.title("Open from Google Drive")
        w.transient(app)
        f = ttk.Frame(w, padding=12)
        f.pack(fill="both", expand=True)
        top = ttk.Frame(f)
        top.pack(fill="x")
        ttk.Label(top, text=f"Google Drive · {email()}").pack(side="left")
        self.search = tk.StringVar(master=w)
        e = ttk.Entry(top, textvariable=self.search, width=28)
        e.pack(side="right")
        ttk.Label(top, text="Search").pack(side="right", padx=(0, 6))
        e.bind("<Return>", lambda ev: self.load())
        cols = ("modified", "kind")
        self.tree = ttk.Treeview(f, columns=cols, height=14, selectmode="browse")
        self.tree.heading("#0", text="Name", anchor="w")
        self.tree.heading("modified", text="Modified", anchor="w")
        self.tree.heading("kind", text="Type", anchor="w")
        scale = getattr(app, "scale", 1)
        self.tree.column("#0", width=int(330 * scale))
        self.tree.column("modified", width=int(130 * scale))
        self.tree.column("kind", width=int(110 * scale))
        self.tree.pack(fill="both", expand=True, pady=8)
        self.tree.bind("<Double-1>", lambda ev: self.open())
        self.tree.bind("<Return>", lambda ev: self.open())
        bottom = ttk.Frame(f)
        bottom.pack(fill="x")
        self.note = ttk.Label(bottom, text="Loading…")
        self.note.pack(side="left")
        ttk.Button(bottom, text="Cancel", command=w.destroy).pack(side="right")
        self.open_btn = ttk.Button(bottom, text="Open", command=self.open, default="active")
        self.open_btn.pack(side="right", padx=6)
        w.bind("<Escape>", lambda ev: w.destroy())
        self.files = {}
        from ai_assist import center
        center(w, app)
        w.deiconify()
        e.focus_set()
        self.load()

    def load(self):
        self.note.configure(text="Loading…")
        text = self.search.get()
        in_background(self.app, lambda: list_files(text), self.show)

    def show(self, files, error):
        if not self.win.winfo_exists():
            return
        self.tree.delete(*self.tree.get_children())
        self.files = {}
        if error:
            self.note.configure(text=error[:120])
            return
        for f in files:
            kind = "Google Sheet" if f["mimeType"] == SHEET else next(
                (e[1:].upper() for e, m in MIMES.items() if m == f["mimeType"]), "File")
            when = (f.get("modifiedTime") or "")[:16].replace("T", " ")
            iid = self.tree.insert("", "end", text=f["name"], values=(when, kind))
            self.files[iid] = f
        self.note.configure(text=f"{len(files)} spreadsheet{'s' if len(files) != 1 else ''}")
        kids = self.tree.get_children()
        if kids:
            self.tree.selection_set(kids[0])
            self.tree.focus(kids[0])

    def open(self):
        sel = self.tree.selection()
        if not sel:
            return None
        f = self.files[sel[0]]
        self.note.configure(text=f"Downloading {f['name']}…")

        def done(path, error):
            if error:
                if self.win.winfo_exists():
                    self.note.configure(text=error[:120])
                return
            if self.win.winfo_exists():
                self.win.destroy()
            open_drive_file(self.app, f, path)
        in_background(self.app, lambda: download(f), done)
        return f


class SaveDialog:
    """Save to Google Drive…: name, and whether it becomes a Google Sheet."""

    def __init__(self, app):
        import tkinter as tk
        from tkinter import ttk
        self.app = app
        w = self.win = tk.Toplevel(app)
        w.withdraw()
        w.title("Save to Google Drive")
        w.transient(app)
        w.resizable(False, False)
        f = ttk.Frame(w, padding=16)
        f.pack(fill="both", expand=True)
        stem = os.path.splitext(os.path.basename(app.path))[0]
        ttk.Label(f, text="Name").grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.name = tk.StringVar(master=w, value=stem)
        ttk.Entry(f, textvariable=self.name, width=40).grid(row=0, column=1, sticky="ew")
        self.as_sheet = tk.BooleanVar(master=w, value=True)
        ttk.Checkbutton(f, text="As a Google Sheet (open it in the browser too)", variable=self.as_sheet).grid(
            row=1, column=1, sticky="w", pady=(8, 0))
        self.note = ttk.Label(f, text=f"Saving as {email()}")
        self.note.grid(row=2, column=0, columnspan=2, sticky="w", pady=(10, 0))
        b = ttk.Frame(f)
        b.grid(row=3, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(b, text="Cancel", command=w.destroy).pack(side="right")
        ttk.Button(b, text="Save", command=self.save, default="active").pack(side="right", padx=6)
        w.bind("<Escape>", lambda e: w.destroy())
        w.bind("<Return>", lambda e: self.save())
        from ai_assist import center
        center(w, app)
        w.deiconify()

    def save(self):
        name, as_sheet, path = self.name.get().strip() or "Untitled", self.as_sheet.get(), self.app.path
        if as_sheet and name.lower().endswith(".xlsx"):
            name = name[:-5]
        self.note.configure(text="Uploading…")

        def done(f, error):
            if error:
                if self.win.winfo_exists():
                    self.note.configure(text=error[:120])
                return
            if self.win.winfo_exists():
                self.win.destroy()
            # from now on this workbook is that Drive file: Save puts it back there
            self.app.drive_file = dict(f, path=os.path.abspath(path))
            self.app.status["text"] = f"Saved {f['name']} to Google Drive"
        in_background(self.app, lambda: create(path, name, as_sheet), done)
        return True


def open_from_drive(app):
    if not need_sign_in(app):
        return None
    return DrivePicker(app)
