# Local window for Partnership Tools. Binds 127.0.0.1 only, serves app\ui and runs each tool
# as a child process (worker.py) so a run can be stopped and a crash never takes the window down.
# Started by "Partnership Tools.bat" with pyembed\pythonw.exe (no console window).
#   python.exe app\server.py --diag   -> same server with the log in a console (for troubleshooting)
import os, re, sys, json, time, glob, codecs, shutil, socket, zipfile, ctypes, tempfile, threading, subprocess, webbrowser
import urllib.request, urllib.error, unicodedata
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs, quote

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import common as C
from common import APP, ROOT, DUMP, OUTPUT

PORT    = 8713
HOST    = "127.0.0.1"
NAME    = "Partnership Tools"
UI      = os.path.join(APP, "ui")
VERFILE = os.path.join(ROOT, "VERSION")
CFGFILE = os.path.join(APP, "update.json")
TOKFILE = os.path.join(DUMP, "ui_token.txt")
LOGFILE = os.path.join(DUMP, "server.log")
UPLOADS = os.path.join(DUMP, "uploads")
BAT     = os.path.join(ROOT, "Partnership Tools.bat")
DIAG    = "--diag" in sys.argv

PY = [sys.executable] if sys.executable else ["py", "-3.12"]
if PY[0].lower().endswith("pythonw.exe"):          # children need a real stdout to stream from
    pc = os.path.join(os.path.dirname(PY[0]), "python.exe")
    if os.path.exists(pc): PY = [pc]
NO_WINDOW = 0x08000000 if os.name == "nt" else 0
ANSI = re.compile(r"\x1b\[[0-9;]*[A-Za-z]")


def log(msg):
    line = time.strftime("%H:%M:%S ") + str(msg)
    if DIAG: print(line, flush=True)
    try:
        os.makedirs(DUMP, exist_ok=True)
        with open(LOGFILE, "a", encoding="utf-8") as f: f.write(line + "\n")
    except Exception: pass


def box(msg, title=NAME):
    """Fatal error popup: pythonw has no console, so a message box is the only way to be seen."""
    if DIAG: print(msg, flush=True); return
    try: ctypes.windll.user32.MessageBoxW(0, str(msg), title, 0x10)
    except Exception: pass


def version():
    try: return open(VERFILE, encoding="utf-8").read().strip() or "0.0.0"
    except Exception: return "0.0.0"


def cfg():
    return C.update_cfg()


def inside(path, roots=(OUTPUT, UPLOADS)):
    p = os.path.realpath(path)
    return any(p == os.path.realpath(r) or p.startswith(os.path.realpath(r) + os.sep) for r in roots)


# ------------------------------------------------------------------ jobs
class Job:
    """One run of one tool: worker.py in its own folder output\\<Tool>\\<date time>\\."""
    def __init__(self, tool, args, label):
        self.tool, self.label = tool, label
        stamp = time.strftime("%Y%m%d_%H%M%S")
        base = os.path.join(OUTPUT, C.TOOLS.get(tool, tool), stamp)
        d, k = base, 2
        while os.path.exists(d): d = "%s_%d" % (base, k); k += 1
        self.dir = d
        os.makedirs(self.dir, exist_ok=True)
        self.args = self.stage(args)
        json.dump(self.args, open(os.path.join(self.dir, "args.json"), "w", encoding="utf-8"), ensure_ascii=False)
        self.lines, self.progress, self.code, self.error = [], None, None, ""
        self.state = "running"               # running | done | failed | stopped
        self.result = None
        self.started, self.ended = time.time(), None
        self.proc, self.stopped = None, False
        self.lock = threading.Lock()

    def stage(self, args):
        """Copy uploaded inputs into the run folder, so the output sits next to its input and
        the run folder alone says what went in (the engines write their output beside it)."""
        out = {}
        for k, v in args.items():
            if isinstance(v, str) and v.startswith("upload:"):
                out[k] = self.copy_in(v[7:])
            elif isinstance(v, list) and v and all(isinstance(x, str) and x.startswith("upload:") for x in v):
                out[k] = [self.copy_in(x[7:]) for x in v]
            else:
                out[k] = v
        return out

    def copy_in(self, uid):
        src = upload_path(uid)
        if not src: raise ValueError("Uploaded file is gone, add it again.")
        dst = os.path.join(self.dir, os.path.basename(src))
        b, e = os.path.splitext(dst); k = 2
        while os.path.exists(dst): dst = "%s (%d)%s" % (b, k, e); k += 1   # two weekly files with one name
        shutil.copy2(src, dst)
        return dst

    def put(self, text):
        with self.lock: self.lines.append(text)

    def feed(self, text):
        text = ANSI.sub("", text).rstrip()
        if not text: return
        if text.startswith("[[P]]"):
            try:
                d, t, lb = text[5:].strip().split("|", 2)
                self.progress = {"done": int(d), "total": int(t), "label": lb}
            except Exception: pass
        elif text.startswith("[[CODE]]"):
            url, _, code = text[8:].strip().partition("|")
            self.code = {"url": url, "code": code}
        elif text.startswith("[[ERROR]]"):
            self.error = text[9:].strip()
        else:
            self.put(text)

    def run(self):
        try:
            e = dict(os.environ); e["PYTHONIOENCODING"] = "utf-8"; e["PYTHONUTF8"] = "1"
            p = subprocess.Popen(PY + ["-u", os.path.join(APP, "worker.py"), self.tool, self.dir], cwd=APP, env=e,
                                 stdin=subprocess.DEVNULL, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                 bufsize=0, creationflags=NO_WINDOW)
            self.proc = p
            dec = codecs.getincrementaldecoder("utf-8")(errors="replace")
            buf = ""
            while True:
                chunk = p.stdout.read(4096)
                if not chunk: break
                for ch in dec.decode(chunk):
                    if ch in "\r\n":
                        self.feed(buf); buf = ""
                    else:
                        buf += ch
            if buf: self.feed(buf)
            rc = p.wait()
            try: self.result = json.load(open(os.path.join(self.dir, "result.json"), encoding="utf-8"))
            except Exception: self.result = None
            if self.stopped and not (self.result and self.result.get("ok")):
                self.state = "stopped"
            elif rc == 0 and self.result and self.result.get("ok"):
                self.state = "done"
            else:
                self.state = "failed"
                if not self.error: self.error = "The run ended with code %s." % rc
        except Exception as ex:
            self.state = "failed"; self.error = str(ex); log("job %s failed: %r" % (self.tool, ex))
        self.ended = time.time()
        self.progress = None

    def kill(self):
        p = self.proc
        if not p or p.poll() is not None: return
        try: subprocess.call(["taskkill", "/F", "/T", "/PID", str(p.pid)], creationflags=NO_WINDOW,
                             stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception: pass

    def stop(self):
        self.stopped = True
        self.kill()

    def view(self, since=0):
        with self.lock: lines = self.lines[since:]
        return {"tool": self.tool, "label": self.label, "state": self.state, "total": len(self.lines),
                "since": since, "lines": lines, "progress": self.progress, "code": self.code,
                "error": self.error, "result": self.result if self.state != "running" else None,
                "dir": self.dir, "elapsed": round((self.ended or time.time()) - self.started)}


JOBS = {}            # tool -> its latest Job
JLOCK = threading.Lock()


def busy(tool):
    j = JOBS.get(tool)
    return bool(j and j.state == "running")


def start_job(tool, args, label):
    with JLOCK:
        if busy(tool): return None, "This tool is already running."
        if tool == "update" and any(busy(t) for t in JOBS): return None, "Wait for the running tools to finish."
        j = Job(tool, args, label)
        JOBS[tool] = j
    threading.Thread(target=j.run, daemon=True).start()
    return j, ""


# ------------------------------------------------------------------ uploads
def upload_save(name, data):
    name = os.path.basename(unicodedata.normalize("NFC", name or "").replace("\\", "/")) or "file"
    uid = os.urandom(6).hex()
    d = os.path.join(UPLOADS, uid); os.makedirs(d, exist_ok=True)
    with open(os.path.join(d, name), "wb") as f: f.write(data)
    return {"id": "upload:" + uid, "name": name, "size": len(data)}


def upload_path(uid):
    d = os.path.join(UPLOADS, re.sub(r"[^0-9a-f]", "", uid))
    fs = os.listdir(d) if os.path.isdir(d) else []
    return os.path.join(d, fs[0]) if fs else ""


def cleanup():
    """Uploads older than a day, runs older than 30 days."""
    now = time.time()
    for d in glob.glob(os.path.join(UPLOADS, "*")):
        if now - os.path.getmtime(d) > 86400: shutil.rmtree(d, ignore_errors=True)
    for d in glob.glob(os.path.join(OUTPUT, "*", "*")):
        try:
            if os.path.isdir(d) and now - os.path.getmtime(d) > 30 * 86400: shutil.rmtree(d, ignore_errors=True)
        except Exception: pass


# ------------------------------------------------------------------ quick look at the inputs (child process)
def inspect(kind, ids):
    """Read the uploaded files the way the run will (file types, suggested period, link count) in a
    short-lived worker, so pandas and the engines never load into the window's process."""
    paths = [upload_path(i[7:] if i.startswith("upload:") else i) for i in ids]
    if not paths or not all(paths): return {"error": "Uploaded file is gone, add it again."}
    fd, af = tempfile.mkstemp(suffix=".json", dir=DUMP); os.close(fd)
    json.dump({"files": paths}, open(af, "w", encoding="utf-8"), ensure_ascii=False)
    e = dict(os.environ); e["PYTHONIOENCODING"] = "utf-8"; e["PYTHONUTF8"] = "1"
    try:
        p = subprocess.run(PY + ["-u", os.path.join(APP, "worker.py"), "--inspect", kind, af], cwd=APP, env=e,
                           capture_output=True, timeout=300, creationflags=NO_WINDOW)
    finally:
        try: os.remove(af)
        except OSError: pass
    out = p.stdout.decode("utf-8", "replace")
    for ln in out.splitlines():
        if ln.startswith("[[JSON]] "): return json.loads(ln[9:])
    log("inspect %s failed: %s" % (kind, (out + p.stderr.decode("utf-8", "replace"))[-800:]))
    return {"error": "Cannot read these files."}


# ------------------------------------------------------------------ environment checks
def installed(progid):
    try:
        import winreg
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid + r"\CLSID"))
        return True
    except Exception:
        return False


def env_info():
    edge = any(os.path.exists(x) for x in (r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                                            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"))
    return {"excel": installed("Excel.Application"), "edge": edge}


# ------------------------------------------------------------------ XHD: Tax ID printouts (Cuc Thue)
# One lookup session for the whole list, kept in this process because the user types captchas
# between calls. The captcha is read by ddddocr first; only when that fails the user types it.
TAX = {"lk": None, "out": "", "zip": "", "todo": [], "state": {}, "rows": {}, "img": None, "msg": "", "manual": set()}
TLOCK = threading.Lock()


def _taxsite():
    p = os.path.join(APP, "engines", "xhd", "tool")
    if p not in sys.path: sys.path.insert(0, p)
    import taxsite
    return taxsite


def _captcha_view(png):
    """The captcha has a transparent background and black text: on the dark theme it is nearly
    invisible. White background, twice the size."""
    try:
        import io
        from PIL import Image
        im = Image.open(io.BytesIO(png)).convert("RGBA")
        bg = Image.new("RGBA", im.size, (255, 255, 255, 255)); bg.alpha_composite(im)
        bg = bg.convert("RGB").resize((im.width * 2, im.height * 2), Image.LANCZOS)
        b = io.BytesIO(); bg.save(b, "PNG"); return b.getvalue()
    except Exception:
        return png


def _rezip():
    """The run's zip again, now with the TaxID printouts (same layout: GCN and TaxID folders)."""
    out = TAX["out"]
    with zipfile.ZipFile(TAX["zip"], "w", zipfile.ZIP_DEFLATED) as z:
        for base, _d, files in os.walk(out):
            for f in files:
                full = os.path.join(base, f)
                z.write(full, os.path.relpath(full, out))


def tax_view():
    import base64
    pending = [t for t in TAX["todo"] if t["mst"] not in TAX["state"]]
    cur = pending[0] if pending else None
    return {"todo": [{"mst": t["mst"], "name": t["name"], "status": TAX["state"].get(t["mst"], ""),
                      "result": TAX["rows"].get(t["mst"], "")} for t in TAX["todo"]],
            "cur": cur, "manual": bool(cur and cur["mst"] in TAX["manual"]), "msg": TAX["msg"],
            "img": base64.b64encode(TAX["img"]).decode() if TAX["img"] else None, "done": not pending}


def tax_start(job_dir):
    res = json.load(open(os.path.join(job_dir, "result.json"), encoding="utf-8"))
    t = res.get("tax") or {}
    with TLOCK:
        if TAX["lk"]:
            try: TAX["lk"].close()
            except Exception: pass
        TAX.update(lk=None, out=t.get("out", ""), zip=t.get("zip", ""), todo=t.get("todo") or [], state={}, rows={},
                   img=None, msg="", manual=set())
        return tax_view()


def tax_step(action, code=""):
    with TLOCK:
        pending = [t for t in TAX["todo"] if t["mst"] not in TAX["state"]]
        if not pending: return tax_view()
        mst = pending[0]["mst"]
        T = _taxsite()
        pdf_dir = os.path.join(TAX["out"], "TaxID")

        def done(r):
            TAX["state"][mst] = "Saved" if r["ok"] else ("Not found" if r["not_found"] else "Saved, check")
            TAX["rows"][mst] = " | ".join(r["rows"][0]) if r["rows"] else ""
            TAX["img"], TAX["msg"] = None, ""
            if r.get("pdf"): _rezip()
        try:
            if TAX["lk"] is None: TAX["lk"] = T.Lookup()
            lk = TAX["lk"]
            if action == "skip":
                TAX["state"][mst] = "Skipped"; TAX["img"], TAX["msg"] = None, ""
            elif action == "auto":
                r = lk.auto(mst, pdf_dir)
                if r["wrong_code"]:
                    TAX["manual"].add(mst); TAX["img"] = _captcha_view(r["captcha"]) if r.get("captcha") else None
                    TAX["msg"] = "Could not read the captcha automatically, type it below."
                else:
                    done(r)
            elif action == "captcha":
                TAX["manual"].add(mst); TAX["img"] = _captcha_view(lk.captcha(mst)); TAX["msg"] = ""
            elif action == "search" and code.strip():
                r = lk.search(mst, code, pdf_dir)
                if r["wrong_code"]:
                    TAX["img"] = _captcha_view(r["captcha"]) if r.get("captcha") else None
                    TAX["msg"] = ("The tax site kept sending the request to another server, type the new captcha."
                                  if r.get("lost") else "Wrong captcha, try the new one.")
                else:
                    done(r)
        except Exception as ex:
            TAX["manual"].add(mst); TAX["img"] = None
            TAX["msg"] = str(ex) if type(ex).__name__ == "TaxSiteError" else "Cannot reach the tax site (%s)." % type(ex).__name__
            log("tax %s %s: %r" % (action, mst, ex))
        if not [t for t in TAX["todo"] if t["mst"] not in TAX["state"]] and TAX["lk"]:
            try: TAX["lk"].close()
            except Exception: pass
            TAX["lk"] = None
        return tax_view()


def state():
    return {"version": version(), "update": UPD.get("cache"),
            "jobs": {t: {"state": j.state, "label": j.label} for t, j in JOBS.items()}}


# ------------------------------------------------------------------ update from GitHub
UPD = {"cache": None, "checked": 0}
SKIP_DIRS  = {"pyembed", "dump", "output", "dist", ".git", "__pycache__", ".vscode"}
ALLOW_EXT  = {".py", ".js", ".html", ".css", ".bat", ".txt", ".json", ".md", ".ico"}
DATA_EXT   = {".xlsx", ".xlsm", ".docx"}       # engine templates: only under app\engines
ALLOW_NAME = {"VERSION"}
KEEP_FILES = {"update.json"}                   # each install keeps its own repo + token


def wanted(rel):
    name = rel.split("/")[-1]
    ext = os.path.splitext(name)[1].lower()
    if name in KEEP_FILES or name.startswith("~$"): return False
    if name in ALLOW_NAME or ext in ALLOW_EXT: return True
    return ext in DATA_EXT and rel.startswith("app/engines/")


http_get = C.http_get


def cmp_ver(a, b):
    na = [int(x) for x in str(a).replace("v", "").split(".") if x.isdigit()]
    nb = [int(x) for x in str(b).replace("v", "").split(".") if x.isdigit()]
    for i in range(max(len(na), len(nb))):
        x, y = (na[i] if i < len(na) else 0), (nb[i] if i < len(nb) else 0)
        if x != y: return 1 if x > y else -1
    return 0


def read_remote(c, name):
    """API first (raw.githubusercontent keeps serving the old file for a few minutes)."""
    repo, branch, token = c["repo"], c["branch"] or "main", c["token"]
    try:
        api = "https://api.github.com/repos/%s/contents/%s?ref=%s" % (repo, name, branch)
        j = json.loads(http_get(api, token).decode("utf-8", "replace"))
        import base64
        return base64.b64decode(j.get("content", "")).decode("utf-8", "replace")
    except urllib.error.HTTPError as ex:
        if ex.code == 404: raise
    except Exception:
        pass
    return http_get("https://raw.githubusercontent.com/%s/%s/%s" % (repo, branch, name), token).decode("utf-8", "replace")


def update_check():
    c = cfg(); cur = version()
    if not c["repo"] or "/" not in c["repo"]:
        out = {"configured": False, "current": cur}
    else:
        try:
            latest = read_remote(c, "VERSION").strip()
            try: notes = read_remote(c, "RELEASE-NOTES.txt")
            except Exception: notes = ""
            out = {"configured": True, "current": cur, "latest": latest,
                   "available": cmp_ver(latest, cur) > 0, "notes": notes[:4000], "repo": c["repo"]}
        except urllib.error.HTTPError as ex:
            msg = ("repo %s not found on GitHub" % c["repo"]) if ex.code in (401, 403, 404) else "GitHub returned %s" % ex.code
            out = {"configured": True, "current": cur, "error": msg, "repo": c["repo"]}
        except Exception as ex:
            out = {"configured": True, "current": cur, "error": "cannot reach GitHub (%s)" % str(ex)[:120], "repo": c["repo"]}
    UPD["cache"] = out; UPD["checked"] = time.time()
    return out


def install_tree(src, j):
    backup = os.path.join(APP, "_backup_" + time.strftime("%Y%m%d_%H%M%S"))
    n = 0
    for base, dirs, files in os.walk(src):
        dirs[:] = [d for d in dirs if d not in SKIP_DIRS and not d.startswith("_backup_")]
        for f in files:
            full = os.path.join(base, f)
            rel = os.path.relpath(full, src).replace("\\", "/")
            if not wanted(rel): continue
            dst = os.path.join(ROOT, rel.replace("/", os.sep))
            if os.path.exists(dst):
                if open(dst, "rb").read() == open(full, "rb").read(): continue
                b = os.path.join(backup, rel.replace("/", os.sep))
                os.makedirs(os.path.dirname(b), exist_ok=True); shutil.copy2(dst, b)
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            shutil.copy2(full, dst); n += 1
            j.put("  " + rel)
    j.put("")
    j.put("%d file(s) updated." % n + (" Old copies kept in app\\%s" % os.path.basename(backup) if n else ""))
    j.put("Version on disk is now %s. Restart the tool to use it." % version())
    UPD["cache"] = None


def update_job(mode, zip_path=""):
    """Runs inside the server (it replaces worker.py itself), so it is a plain thread, not a child."""
    with JLOCK:
        if any(busy(t) for t in JOBS): return None, "Wait for the running tools to finish."
        j = Job.__new__(Job)
        j.tool, j.label, j.dir, j.args = "update", "Update", "", {}
        j.lines, j.progress, j.code, j.error, j.result = [], None, None, "", None
        j.state, j.started, j.ended, j.proc, j.stopped = "running", time.time(), None, None, False
        j.lock = threading.Lock()
        JOBS["update"] = j

    def body():
        tmp = tempfile.mkdtemp(prefix="ct_upd_")
        try:
            if mode == "local":
                if not os.path.exists(zip_path): raise ValueError("File not found: %s" % zip_path)
                j.put("Installing from %s" % zip_path)
                zp = zip_path
            else:
                c = cfg()
                if not c["repo"]: raise ValueError("app\\update.json has no repo yet.")
                url = "https://codeload.github.com/%s/zip/refs/heads/%s" % (c["repo"], c["branch"] or "main")
                j.put("Downloading %s ..." % url)
                zp = os.path.join(tmp, "src.zip")
                open(zp, "wb").write(http_get(url, c["token"], timeout=180))
            ex = os.path.join(tmp, "x")
            with zipfile.ZipFile(zp) as z: z.extractall(ex)
            roots = [os.path.join(ex, d) for d in os.listdir(ex) if os.path.isdir(os.path.join(ex, d))]
            src = roots[0] if len(roots) == 1 and not os.path.exists(os.path.join(ex, "VERSION")) else ex
            install_tree(src, j)
            j.state = "done"; j.result = {"ok": True}
        except urllib.error.HTTPError as ex:
            j.error = "GitHub returned %s. Check the repo and branch in app\\update.json." % ex.code; j.state = "failed"
        except Exception as ex:
            j.error = "Update failed: %s" % ex; j.state = "failed"
        finally:
            j.ended = time.time()
            shutil.rmtree(tmp, ignore_errors=True)
    threading.Thread(target=body, daemon=True).start()
    return j, ""


# ------------------------------------------------------------------ http
TOKEN = ""
LAST_PING = [time.time()]
SHUTDOWN = []
MIME = {".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".xlsm": "application/vnd.ms-excel.sheet.macroEnabled.12", ".zip": "application/zip",
        ".pdf": "application/pdf", ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"}


class Handler(BaseHTTPRequestHandler):
    server_version = "PartnershipTools"
    protocol_version = "HTTP/1.1"

    def log_message(self, *a): pass

    def _send(self, code, body=b"", ctype="application/json; charset=utf-8", extra=None):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        for k, v in (extra or {}).items(): self.send_header(k, v)
        self.end_headers()
        try: self.wfile.write(body)
        except Exception: pass

    def json(self, obj, code=200):
        self._send(code, json.dumps(obj, ensure_ascii=False, default=str).encode("utf-8"))

    def ok_token(self, q):
        ok = (self.headers.get("X-Token") or q.get("t", [""])[0]) == TOKEN
        if ok: LAST_PING[0] = time.time()        # any call from the page means a window is open
        return ok

    def body(self):
        try: n = int(self.headers.get("Content-Length") or 0)
        except ValueError: n = 0
        if not n: return {}
        try: return json.loads(self.rfile.read(n).decode("utf-8"))
        except Exception: return {}

    def send_file(self, path, attach=True):
        name = os.path.basename(path)
        data = open(path, "rb").read()
        disp = "attachment" if attach else "inline"
        self._send(200, data, MIME.get(os.path.splitext(name)[1].lower(), "application/octet-stream"),
                   {"Content-Disposition": "%s; filename*=UTF-8''%s" % (disp, quote(name))})

    def do_GET(self):
        u = urlparse(self.path); q = parse_qs(u.query); p = u.path
        g = lambda k: q.get(k, [""])[0]
        if p in ("/", "/index.html"):
            LAST_PING[0] = time.time()
            f = os.path.join(UI, "index.html")
            if not os.path.exists(f): return self._send(404, b"ui missing", "text/plain")
            return self._send(200, open(f, "rb").read(), "text/html; charset=utf-8")
        if not self.ok_token(q): return self.json({"error": "token"}, 403)
        if p == "/api/state": return self.json(state())
        if p == "/api/env": return self.json(env_info())
        if p == "/api/job":
            j = JOBS.get(g("tool"))
            if not j: return self.json({"job": None})
            return self.json({"job": j.view(int(g("since") or 0))})
        if p == "/api/update/check": return self.json(update_check())
        if p == "/api/file":
            path = g("path")
            if not path or not inside(path) or not os.path.isfile(path): return self.json({"error": "no file"}, 404)
            return self.send_file(path)
        return self._send(404, b"not found", "text/plain")

    def do_POST(self):
        u = urlparse(self.path); q = parse_qs(u.query); p = u.path
        if not self.ok_token(q): return self.json({"error": "token"}, 403)
        if p == "/api/upload":
            try: n = int(self.headers.get("Content-Length") or 0)
            except ValueError: n = 0
            if n > 300 * 1024 * 1024: return self.json({"error": "File too large (over 300 MB)."}, 413)
            data = self.rfile.read(n) if n else b""
            try: return self.json(upload_save(q.get("name", [""])[0], data))
            except Exception as ex:
                log("upload: %r" % ex); return self.json({"error": str(ex)}, 500)
        d = self.body()
        if p == "/api/ping":
            LAST_PING[0] = time.time(); return self.json({"ok": True})
        if p == "/api/inspect":
            try: return self.json(inspect(d.get("kind"), d.get("ids") or []))
            except Exception as ex:
                log("inspect: %r" % ex); return self.json({"error": "Cannot read this file: %s" % ex})
        if p == "/api/run":
            tool = d.get("tool")
            if tool == "update":
                j, err = update_job(d.get("mode") or "github", d.get("zip") or "")
            elif tool in C.TOOLS:
                try: j, err = start_job(tool, d.get("args") or {}, d.get("label") or tool)
                except Exception as ex: j, err = None, str(ex)
            else:
                return self.json({"error": "unknown tool"}, 400)
            if not j: return self.json({"error": err}, 409)
            return self.json({"ok": True})
        if p == "/api/tax/start":
            dir_ = d.get("dir") or ""
            if not inside(dir_) or not os.path.exists(os.path.join(dir_, "result.json")):
                return self.json({"error": "Run not found"}, 404)
            return self.json(tax_start(dir_))
        if p == "/api/tax/step":
            return self.json(tax_step(d.get("action") or "", d.get("code") or ""))
        if p == "/api/stop":
            j = JOBS.get(d.get("tool"))
            if j and j.state == "running": j.stop()
            return self.json({"ok": True})
        if p == "/api/open":
            # a file or folder of a run
            if d.get("url"):
                url = d["url"]
                if not url.startswith("https://"): return self.json({"error": "bad link"}, 400)
                webbrowser.open(url); return self.json({"ok": True})
            path = d.get("path") or ""
            if d.get("what") == "output": path = OUTPUT
            if not path or not (inside(path) or os.path.realpath(path) == os.path.realpath(OUTPUT)) or not os.path.exists(path):
                return self.json({"error": "Not found"}, 404)
            try:
                if d.get("select") and os.path.isfile(path):
                    subprocess.Popen(["explorer", "/select,", path])
                else:
                    os.startfile(path)
            except Exception as ex: return self.json({"error": str(ex)}, 500)
            return self.json({"ok": True})
        if p == "/api/restart":
            restart(); return self.json({"ok": True})
        if p == "/api/quit":
            for j in JOBS.values():
                if j.state == "running" and j.proc: j.kill()
            threading.Timer(0.4, lambda: SHUTDOWN and SHUTDOWN[0]()).start()
            return self.json({"ok": True})
        return self.json({"error": "not found"}, 404)


def restart():
    """Detached helper waits for this process to free the port, then runs the .bat."""
    helper = "import os,time\ntime.sleep(2.0)\nos.startfile(r'%s')\n" % BAT
    exe = sys.executable or "pythonw.exe"
    pw = os.path.join(os.path.dirname(exe), "pythonw.exe")
    subprocess.Popen([pw if os.path.exists(pw) else exe, "-c", helper], creationflags=NO_WINDOW | 0x00000008)
    threading.Timer(0.5, lambda: SHUTDOWN and SHUTDOWN[0]()).start()


def app_windows(title=NAME):
    if os.name != "nt": return set()
    u32 = ctypes.windll.user32
    proto = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
    hits = set()
    def cb(hwnd, _):
        n = u32.GetWindowTextLengthW(hwnd)
        if n and u32.IsWindowVisible(hwnd):
            buf = ctypes.create_unicode_buffer(n + 1); u32.GetWindowTextW(hwnd, buf, n + 1)
            if buf.value == title: hits.add(hwnd)
        return True
    u32.EnumWindows(proto(cb), 0)
    return hits


def maximize_window(before, wait=20):
    """Edge ignores --start-maximized when it is already running, so maximize the new window ourselves."""
    if os.name != "nt": return
    u32 = ctypes.windll.user32
    end = time.time() + wait
    while time.time() < end:
        new = app_windows() - before
        if new:
            for h in new:
                if not u32.IsZoomed(h): u32.ShowWindow(h, 3)
            return
        time.sleep(0.3)


def open_ui():
    url = "http://%s:%d/?t=%s" % (HOST, PORT, TOKEN)
    if os.environ.get("CT_NO_OPEN"): log("CT_NO_OPEN set, window not opened: " + url); return
    for exe in (r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"):
        if os.path.exists(exe):
            try:
                before = app_windows()
                subprocess.Popen([exe, "--app=" + url, "--start-maximized", "--disable-features=Translate"],
                                 creationflags=NO_WINDOW)
                t = threading.Thread(target=maximize_window, args=(before,), daemon=True)
                t.start()
                return t
            except Exception: pass
    webbrowser.open(url)


def already_running():
    s = socket.socket()
    try:
        s.connect((HOST, PORT)); return True
    except Exception:
        return False
    finally:
        s.close()


def shortcuts():
    """Unzip-and-run: the first start adds the app to the Desktop and the Start menu,
    pointing at this copy, so next time nobody has to find the extracted folder. An existing shortcut
    (from Install.bat or an earlier start) is left alone."""
    if os.name != "nt": return
    target = os.path.join(APP, "pyembed", "pythonw.exe")
    if not os.path.exists(target): return
    ps = ("$s = New-Object -ComObject WScript.Shell;"
          "foreach ($d in @([Environment]::GetFolderPath('Programs'), [Environment]::GetFolderPath('Desktop'))) {"
          " $p = Join-Path $d '%s.lnk'; if (Test-Path -LiteralPath $p) { continue };"
          " $l = $s.CreateShortcut($p); $l.TargetPath = '%s'; $l.Arguments = '\"%s\"'; $l.WorkingDirectory = '%s';"
          " $l.IconLocation = '%s,0'; $l.Description = '%s - PVI Digital'; $l.Save() }"
          % (NAME, target, os.path.join(APP, "server.py"), ROOT, os.path.join(UI, "app.ico"), NAME))
    try:
        subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", ps], timeout=60,
                       creationflags=NO_WINDOW, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception as ex:
        log("shortcuts: %r" % ex)


def watchdog():
    """No page talking to us for 2 minutes and nothing running -> close the server."""
    while True:
        time.sleep(10)
        if any(j.state == "running" for j in JOBS.values()): LAST_PING[0] = time.time(); continue
        if time.time() - LAST_PING[0] > 120:
            log("no page for 2 minutes, closing"); SHUTDOWN and SHUTDOWN[0](); return


def main():
    global TOKEN
    for d in (DUMP, UPLOADS, OUTPUT): os.makedirs(d, exist_ok=True)
    if already_running():
        try: TOKEN = open(TOKFILE, encoding="utf-8").read().strip()
        except Exception: TOKEN = ""
        log("already running, opening the window")
        t = open_ui()
        if t: t.join()
        return
    TOKEN = os.urandom(16).hex()
    open(TOKFILE, "w", encoding="utf-8").write(TOKEN)
    try:
        srv = ThreadingHTTPServer((HOST, PORT), Handler)
    except OSError as ex:
        box("Cannot open port %d (%s).\nClose the running tool, then open it again." % (PORT, ex)); return
    SHUTDOWN.append(lambda: threading.Thread(target=srv.shutdown, daemon=True).start())
    log("listening on %s:%d  v%s" % (HOST, PORT, version()))
    threading.Thread(target=watchdog, daemon=True).start()
    threading.Thread(target=cleanup, daemon=True).start()
    threading.Thread(target=lambda: (time.sleep(1.5), update_check()), daemon=True).start()
    threading.Thread(target=shortcuts, daemon=True).start()
    open_ui()
    try: srv.serve_forever()
    except KeyboardInterrupt: pass
    log("closed")


if __name__ == "__main__":
    try:
        main()
    except Exception as ex:
        import traceback
        log(traceback.format_exc())
        box("The tool failed to start:\n\n%s\n\nLog: %s" % (ex, LOGFILE))
