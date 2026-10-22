# -*- coding: utf-8 -*-
"""Phan chay cua Partnership Tools (ban desktop: chep tu Partnership-Web\runners.py, chi doi duong dan engine).

Engine CHEP tu workspace goc vao app\\engines (giu nguyen bo cuc de config.json/templates tim duoc nhau):
  - engines\\d2c\\tool   <- Documents\\D2C-VNPAY-Reconcile\\tool  (reconcile.build + recalc_win.recalc, template.xlsx)
  - engines\\ptt         <- Documents\\PPT-Partner                (generate_ptt; config.json + templates\\)
  - engines\\xhd         <- Documents\\Invoice-Issue              (xhd, taxsite; config.json + templates\\)
  - engines\\kbank       <- Documents\\Kbank-GCNBH                (kbank_engine: run, zip_bytes; template.docx)
Sua engine o workspace goc thi chep lai sang day roi day ban moi len GitHub.

Moi lan chay mot thu muc rieng runs\\<loai>\\<gio>-<ma>\\, qua KEEP_DAYS ngay tu xoa.
Mot luc chi chay 1 viec (LOCK, chung ca app): D2C dung Excel COM, PTT doi bien toan cuc cua engine.
"""
import contextlib
import datetime as dt
import glob
import io
import json
import os
import re
import shutil
import sys
import threading
import time
import unicodedata
import uuid
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
ENGINES = os.path.join(HERE, "engines")         # desktop app: copies of the 4 workspaces, same layout
D2C_TOOL = os.path.join(ENGINES, "d2c", "tool")
PTT_ROOT = os.path.join(ENGINES, "ptt")
XHD_TOOL = os.path.join(ENGINES, "xhd", "tool")
KB_TOOL = os.path.join(ENGINES, "kbank")
TEMPLATE = os.path.join(D2C_TOOL, "template.xlsx")
RUNS = os.path.join(HERE, "runs")              # worker.py points this at the run folder
KEEP_DAYS = 14

for _p in (D2C_TOOL, os.path.join(PTT_ROOT, "tool"), XHD_TOOL):
    if _p not in sys.path:
        sys.path.insert(0, _p)
os.environ["PTT_NO_OPEN"] = "1"           # generate_ptt khong tu mo Explorer tren may chu
os.environ["PTT_NO_ASK"] = "1"            # khong hoi ngay tren console, web truyen --tu/--den

import reconcile as B                      # noqa: E402
from recalc_win import recalc              # noqa: E402
import generate_ptt as G                   # noqa: E402
import xhd as X                            # noqa: E402
import taxsite as T                        # noqa: E402


def _load_kbank():
    """Nap kbank_engine KHONG qua sys.path: no 'import docx_fill', trung ten docx_fill cua PPT-Partner
    (khac noi dung, generate_ptt can replace_everywhere). Tam dat docx_fill cua KBank vao sys.modules
    trong luc nap engine roi tra lai ban cu, nen moi engine giu dung docx_fill cua minh."""
    import importlib.util

    def load(name, path):
        spec = importlib.util.spec_from_file_location(name, path)
        mod = importlib.util.module_from_spec(spec)
        sys.modules[name] = mod                    # @dataclass can tim module theo ten
        spec.loader.exec_module(mod)
        return mod

    saved = sys.modules.get("docx_fill")
    sys.modules["docx_fill"] = load("kb_docx_fill", os.path.join(KB_TOOL, "docx_fill.py"))
    try:
        return load("kbank_engine", os.path.join(KB_TOOL, "kbank_engine.py"))
    finally:
        if saved is None:
            sys.modules.pop("docx_fill", None)
        else:
            sys.modules["docx_fill"] = saved


K = _load_kbank()

G.PACE = 0                                # thanh tien trinh console: tren web khong can cho nhip

LOCK = threading.Lock()


class RunError(Exception):
    pass


def nfc(s):
    return unicodedata.normalize("NFC", s)


def _cleanup(kind):
    base = os.path.join(RUNS, kind)
    if not os.path.isdir(base):
        return
    cut = time.time() - KEEP_DAYS * 86400
    for d in os.listdir(base):
        p = os.path.join(base, d)
        try:
            if os.path.isdir(p) and os.path.getmtime(p) < cut:
                shutil.rmtree(p, ignore_errors=True)
        except OSError:
            pass


def new_run(kind):
    _cleanup(kind)
    run = os.path.join(RUNS, kind, dt.datetime.now().strftime("%Y%m%d-%H%M%S") + "-" + uuid.uuid4().hex[:6])
    os.makedirs(os.path.join(run, "input"))
    os.makedirs(os.path.join(run, "output"))
    return run


def save_files(files, folder):
    paths = []
    for name, data in files:
        p = os.path.join(folder, nfc(os.path.basename(name)))
        with open(p, "wb") as fh:
            fh.write(data)
        paths.append(p)
    return paths


def zip_bytes(paths):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for p in paths:
            z.write(p, os.path.basename(p))
    return buf.getvalue()


@contextlib.contextmanager
def one_at_a_time(on_wait=None):
    if not LOCK.acquire(blocking=False):
        if on_wait:
            on_wait()
        LOCK.acquire()
    try:
        yield
    finally:
        LOCK.release()


# ======================================================================= D2C VNPAY

# cung mau ten file voi reconcile.build (_find), de bao truoc file nao thieu/thua
D2C_KINDS = [
    ("vnpay", "VNPAY statement", (r"^vnpay.*\.xlsx$", r"danhsachgiaodich.*\.xlsx$"), 1),
    ("admin", "Admin export", (r"^admin.*\.csv$",), 2),
    ("pias", "PIAS export", (r"^pias.*\.xlsx$",), 4),
    ("invoice", "Invoice ledger", (r"hóa đơn thu phí.*\.xlsx$", r"hoa don thu phi.*\.xlsx$"), 1),
]
D2C_LABEL = {k: lab for k, lab, _, _ in D2C_KINDS}


def d2c_kind(name):
    n = nfc(os.path.basename(name))
    for k, _, pats, _ in D2C_KINDS:
        if any(re.search(p, n, re.I) for p in pats):
            return k
    return None


def d2c_problems(names):
    """Thieu loai nao / thua VNPAY, hoa don thi chan nut Run. Thieu so file Admin/PIAS chi canh bao."""
    kinds = [d2c_kind(n) for n in names]
    block, warn = [], []
    for k, lab, _, want in D2C_KINDS:
        c = kinds.count(k)
        if c == 0:
            block.append(f"Missing the {lab}.")
        elif want == 1 and c > 1:
            block.append(f"Upload only one {lab} file.")
        elif c < want:
            warn.append(f"{lab}: {c} of {want} files. Rows of the missing ones will not be matched.")
    return block, warn


def d2c_period(df):
    """Ky = thang xuat hien nhieu nhat trong Thoi gian GD cua giao dich thanh cong -> 'MM.YYYY'."""
    ts = df["_ts"].dropna() if "_ts" in df else []
    if len(ts) == 0:
        return dt.date.today().strftime("%m.%Y")
    top = ts.dt.to_period("M").value_counts().idxmax()
    return f"{top.month:02d}.{top.year}"


def d2c_run(files, log=print, progress=None, on_wait=None):
    """files = [(ten, bytes)]. Tra ve dict: run, output, period, stats, recalc, log_text."""
    block, _ = d2c_problems([n for n, _ in files])
    if block:
        raise RunError(" ".join(block))
    lines = []

    def say(m, tag=None):
        lines.append(str(m))
        log(m)

    with one_at_a_time(on_wait):
        r = new_run("d2c")
        inp, out_dir = os.path.join(r, "input"), os.path.join(r, "output")
        save_files(files, inp)
        tmp = os.path.join(out_dir, "result.xlsx")
        try:
            df, vn, tong, n = B.build(inp, TEMPLATE, tmp, progress=progress, log=say)
        except SystemExit as e:              # engine bao loi dau vao bang SystemExit(message)
            raise RunError(str(e).replace("[LỖI]", "").replace("[ERROR]", "").strip()) from None
        if progress:
            progress(92, "Recalculating formulas")
        rc = None
        try:
            import pythoncom
            pythoncom.CoInitialize()         # luong cua streamlit chua khoi tao COM cho Excel
            try:
                rc = recalc(tmp)
            finally:
                pythoncom.CoUninitialize()
            say(f"{rc['total_formulas']} formulas, {rc.get('total_errors', 0)} errors")
        except Exception as e:
            say(f"Skipped recalc ({type(e).__name__}: {e})")
        ky = d2c_period(df)
        out = os.path.join(out_dir, f"D2C - VNPAY - Reconciliation {ky}.xlsx")
        os.replace(tmp, out)
        say(B.summary(df, tong, n))
        stats = {
            "transactions": int(tong),
            "successful": int(n),
            "matched": int(df["_khop_admin"].sum()),
            "policy": int(df["_don"].notna().sum()),
            "invoice": int(df["_hoadon"].notna().sum()),
        }
        with io.open(os.path.join(out_dir, "run-log.txt"), "w", encoding="utf-8") as fh:
            fh.write("\n".join(lines) + "\n")
        if progress:
            progress(100, "Done")
        return {"run": r, "output": out, "period": ky, "stats": stats, "recalc": rc,
                "log_text": "\n".join(lines)}


# ============================================================================ PTT GSM

# "[GREEN_EXPRESS]  3 policy  |  premium 164 834 545 + VAT 16 483 455 = 181 318 000" (tien co dau cach)
SLIP_RE = re.compile(r"^\[(\w+)\]\s+(\d+) policy\s+\|\s+premium ([\d .,]+?) \+ VAT ([\d .,]+?) = ([\d .,]+?)\s*$", re.M)


def _day(v):
    return v.date() if isinstance(v, dt.datetime) else v


def ptt_config():
    with io.open(os.path.join(PTT_ROOT, "config.json"), encoding="utf-8") as fh:
        return json.load(fh)


def ptt_preview(files):
    """Luu file vao mot run moi, doc nhu tool goc va tra ky goi y de dien san vao o ngay.
    Tra ve dict: run, paths, auto_tu, auto_den, per_start {date: so don}, warns, counts."""
    r = new_run("ptt")
    paths = save_files(files, os.path.join(r, "input"))
    cfg = ptt_config()
    warns = []
    try:
        with contextlib.redirect_stdout(io.StringIO()):
            data, _dup = G.read_rows(paths, warns, G.Progress(len(paths)))
        if not any(data.values()):
            raise RunError("No policy row found. Upload the VCND and TNCN exports from PIAS.")
        tu, den, per_start = G.suggest_period(data, cfg["programs"], cfg.get("status_include") or [])
    except G.ToolError as e:
        raise RunError(str(e)) from None
    return {"run": r, "paths": paths, "auto_tu": _day(tu), "auto_den": _day(den),
            "per_start": {_day(d): n for d, n in per_start.items()}, "warns": warns,
            "counts": {k: len(v) for k, v in data.items()}}


def ptt_covered(per_start, tu, den):
    return sum(n for d, n in per_start.items() if tu <= d < den)


def ptt_run(preview, tu, den, lap, on_wait=None):
    """Chay y nhu Run_PTT.bat voi --tu/--den/--ngay-lap tren file da luu o preview.
    Tra ve dict: run, outputs (docx truoc, log.txt sau), slips, warns, console, log_text."""
    run = preview["run"]
    f = lambda d: d.strftime("%d/%m/%Y")                         # noqa: E731
    args = preview["paths"] + ["--tu", f(tu), "--den", f(den), "--ngay-lap", f(lap)]
    buf = io.StringIO()
    with one_at_a_time(on_wait):
        out_root = os.path.join(run, "output")
        shutil.rmtree(out_root, ignore_errors=True)               # chay lai cung bo file: bo ket qua cu
        os.makedirs(out_root)
        old = G.OUT_ROOT, G.LAST_DIR
        G.OUT_ROOT, G.LAST_DIR = out_root, os.path.join(run, "last_folder.txt")
        try:
            with contextlib.redirect_stdout(buf):
                G.main(args)
        except G.ToolError as e:
            raise RunError(str(e)) from None
        finally:
            G.OUT_ROOT, G.LAST_DIR = old
    console = buf.getvalue()
    outs = [p for d in glob.glob(os.path.join(run, "output", "*")) for p in glob.glob(os.path.join(d, "*"))]
    outs.sort(key=lambda p: (p.endswith(".txt"), os.path.basename(p)))
    log_path = next((p for p in outs if os.path.basename(p) == "log.txt"), None)
    log_text = io.open(log_path, encoding="utf-8").read() if log_path else ""
    warns = log_text.split("Warnings (", 1)[1].split("\n", 1)[1].strip() if "Warnings (" in log_text else ""
    warns = [] if warns in ("", "(none)") else warns.splitlines()
    slips = [{"slip": m.group(1), "policies": int(m.group(2)), "premium": m.group(3),
              "vat": m.group(4), "total": m.group(5)} for m in SLIP_RE.finditer(console)]
    return {"run": run, "outputs": outs, "slips": slips, "warns": warns,
            "console": console, "log_text": log_text}


# ======================================================================= INVOICE PDFs

# Link hoa don dien tu PVI (hddt.pvi.com.vn/GetInvoice.aspx?id=...) tra ve .zip gom 1 .pdf + 1 .xml.
# Trang nay doc cot link trong file Excel, bo link trung (giu lan dau), tai theo thu tu dong,
# lay rieng PDF va gop thanh 1 file de in hang loat. Chi goi mang, khong dung LOCK chung.
URL_RE = re.compile(r"https?://[^\s\"'<>]+", re.I)
HYPER_RE = re.compile(r'HYPERLINK\(\s*"([^"]+)"', re.I)
INV_WORKERS = 6
INV_TRIES = 3


def _cell_url(cell, value):
    """Link cua 1 o: hyperlink gan vao o > URL trong gia tri (ca gia tri da tinh cua cong thuc)
    > URL trong cong thuc =HYPERLINK("...")."""
    if cell.hyperlink is not None and (cell.hyperlink.target or "").lower().startswith("http"):
        return cell.hyperlink.target.strip()
    if isinstance(value, str):
        m = URL_RE.search(value)
        if m:
            return m.group(0)
    if isinstance(cell.value, str) and cell.value.startswith("="):
        m = HYPER_RE.search(cell.value)
        if m and m.group(1).lower().startswith("http"):
            return m.group(1)
    return None


def inv_scan(name, data):
    """Tim cot co nhieu link nhat (moi sheet), tra ve dict:
    sheet, col, header, items [{row, url, hidden}] theo thu tu dong (chua bo trung)."""
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
    if not nfc(name).lower().endswith((".xlsx", ".xlsm")):
        raise RunError("Upload an .xlsx or .xlsm file.")
    try:
        wf = load_workbook(io.BytesIO(data))                  # cong thuc + hyperlink
        wv = load_workbook(io.BytesIO(data), data_only=True)  # gia tri da tinh Excel luu san
    except Exception as e:
        raise RunError(f"Cannot open the file ({type(e).__name__}).") from None
    best = None
    for ws in wf.worksheets:
        vs = wv[ws.title]
        cols = {}
        for row in ws.iter_rows():
            for c in row:
                u = _cell_url(c, vs.cell(c.row, c.column).value)
                if u:
                    cols.setdefault(c.column, []).append((c.row, u))
        for col, hits in cols.items():
            if best is None or len(hits) > len(best[2]):
                best = (ws, col, hits)
    if best is None:
        raise RunError("No invoice link found in the file. If the links come from a formula, "
                       "open the file in Excel, let it calculate, save and upload again.")
    ws, col, hits = best
    first = hits[0][0]
    head = ws.cell(first - 1, col).value if first > 1 else None
    items = [{"row": r, "url": u, "hidden": bool(ws.row_dimensions[r].hidden)} for r, u in hits]
    return {"sheet": ws.title, "col": get_column_letter(col),
            "header": str(head).strip() if head not in (None, "") else "", "items": items}


def inv_unique(items, visible_only=False):
    """Bo dong an (neu chon) roi bo link trung, giu lan xuat hien dau. Tra ve (list, so dong trung)."""
    seen, out, dup = set(), [], 0
    for it in items:
        if visible_only and it["hidden"]:
            continue
        if it["url"] in seen:
            dup += 1
            continue
        seen.add(it["url"])
        out.append(it)
    return out, dup


def _inv_fetch(sess, url):
    """Tai 1 link, tra ve (ten pdf, bytes pdf). Nhan ca .zip (pdf+xml) lan .pdf tra thang."""
    last = None
    for i in range(INV_TRIES):
        try:
            r = sess.get(url, timeout=60)
            r.raise_for_status()
            body = r.content
            cd = r.headers.get("Content-Disposition", "")
            fname = (re.search(r'filename="?([^";]+)', cd) or [None, ""])[1].strip()
            if body[:4] == b"%PDF":
                return (os.path.splitext(fname)[0] or "invoice") + ".pdf", body
            if body[:2] == b"PK":
                with zipfile.ZipFile(io.BytesIO(body)) as z:
                    pdfs = [n for n in z.namelist() if n.lower().endswith(".pdf")]
                    if not pdfs:
                        raise RunError("no PDF inside the zip")
                    return os.path.basename(pdfs[0]), z.read(pdfs[0])
            raise RunError("the link did not return a PDF or zip (maybe the invoice id is wrong)")
        except RunError:
            raise
        except Exception as e:
            last = e
            time.sleep(1.5 * (i + 1))
    raise RunError(f"{type(last).__name__}: {last}")


def inv_run(items, progress=None):
    """items = list sau inv_unique. Tai song song, gop PDF dung thu tu items.
    Tra ve dict: run, merged, zip, rows [{No, Row, Invoice, Status}], failed, pages."""
    import requests
    from concurrent.futures import ThreadPoolExecutor, as_completed
    from pypdf import PdfReader, PdfWriter
    run = new_run("inv")
    pdf_dir = os.path.join(run, "output", "pdf")
    os.makedirs(pdf_dir)
    sess = requests.Session()
    sess.headers["User-Agent"] = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    got = [None] * len(items)
    done = 0
    with ThreadPoolExecutor(INV_WORKERS) as ex:
        futs = {ex.submit(_inv_fetch, sess, it["url"]): i for i, it in enumerate(items)}
        for f in as_completed(futs):
            i = futs[f]
            try:
                got[i] = f.result()
            except Exception as e:
                got[i] = e
            done += 1
            if progress:
                progress(int(done * 90 / len(items)), f"Downloaded {done} of {len(items)}")
    if progress:
        progress(92, "Merging PDFs")
    w, rows, files, pages = PdfWriter(), [], [], 0
    for i, (it, g) in enumerate(zip(items, got), 1):
        if isinstance(g, Exception):
            rows.append({"No": i, "Row": it["row"], "Invoice": "", "Status": f"Failed: {g}"})
            continue
        name, body = g
        p = os.path.join(pdf_dir, f"{i:03d}_{name}")
        with open(p, "wb") as fh:
            fh.write(body)
        files.append(p)
        try:
            n0 = len(w.pages)
            w.append(PdfReader(io.BytesIO(body)))
            pages += len(w.pages) - n0
            rows.append({"No": i, "Row": it["row"], "Invoice": name, "Status": "OK"})
        except Exception as e:
            rows.append({"No": i, "Row": it["row"], "Invoice": name,
                         "Status": f"Downloaded, not merged ({type(e).__name__})"})
    merged = None
    if pages:
        merged = os.path.join(run, "output", "Invoices_%s.pdf" % dt.date.today().strftime("%Y%m%d"))
        with open(merged, "wb") as fh:
            w.write(fh)
    if progress:
        progress(100, "Done")
    return {"run": run, "merged": merged, "zip": zip_bytes(files) if files else None,
            "rows": rows, "failed": [r for r in rows if r["Status"] != "OK"], "pages": pages,
            "ok": sum(r["Status"] == "OK" for r in rows)}


# ============================================================================ XHD (invoice requests)

# Engine: Documents\Invoice-Issue\tool\xhd.py (doc CLAUDE.md ben do ve luat). Tai file len la doc ngay
# (nhan loai theo tieu de) de goi y ky gom khach le; Run chay ca 2 luong: theo yeu cau + gom khach le.

def xhd_preview(files):
    """Luu file, doc, tinh ky goi y cho luong gom. Tra ve dict: run, paths, inp, files [(ten, loai)],
    auto_tu, auto_den, per_day {date: so don}, n_req, n_pias, warns."""
    r = new_run("xhd")
    paths = save_files(files, os.path.join(r, "input"))
    warns = []
    try:
        inp = X.load_inputs(paths, warns)
        tu = den = None
        per = {}
        if inp["pias"] is not None and inp["invoices"] is not None:
            cand = X.bulk_candidates(X.plan_bulk(inp, X.load_config(), []))
            tu, den = X.suggest_period(cand)
            per = X.per_day(cand)
    except X.ToolError as e:
        raise RunError(str(e)) from None
    n_req = 0
    if inp["requests"] is not None:
        n_req = sum(1 for v in inp["requests"]["Mã SV / Mã yêu cầu"].tolist() if X.txt(v))
    n_pias = sum(1 for _, k in inp["files"] if k == "pias")
    return {"run": r, "paths": paths, "inp": inp, "files": inp["files"], "auto_tu": tu, "auto_den": den,
            "per_day": per, "n_req": n_req, "n_pias": n_pias, "warns": warns}


def xhd_run(preview, tu, den, lap, online=True, log=print, on_wait=None):
    """Chay engine tren du lieu da doc o preview. Tra ve dict cua xhd.run_loaded + run, zip, log_text."""
    run = preview["run"]
    out = os.path.join(run, "output")
    lines = []

    def say(m):
        lines.append(str(m))
        log(m)

    with one_at_a_time(on_wait):
        shutil.rmtree(out, ignore_errors=True)                    # chay lai cung bo file: bo ket qua cu
        os.makedirs(out)
        try:
            res = X.run_loaded(preview["inp"], out, period=(tu, den) if tu and den else None, slip_date=lap,
                               tax=online, gcn=online, log=say, warns=list(preview["warns"]))
        except X.ToolError as e:
            raise RunError(str(e)) from None
        xls = [p for p in res["outputs"] if p.endswith(".xlsx")]
        if xls:
            say("Recalculating formulas")
            try:
                import pythoncom
                pythoncom.CoInitialize()     # luong cua streamlit chua khoi tao COM cho Excel
                try:
                    for p in xls:
                        rc = recalc(p)
                        if rc.get("total_errors"):
                            res["warns"].append(f"{os.path.basename(p)}: {rc['total_errors']} formula error(s).")
                finally:
                    pythoncom.CoUninitialize()
            except Exception as e:
                say(f"Skipped recalc ({type(e).__name__}: {e})")
    with io.open(os.path.join(out, "run-log.txt"), "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines + ["", "Warnings:"] + (res["warns"] or ["(none)"])) + "\n")
    res["outputs"].append(os.path.join(out, "run-log.txt"))
    res["run"] = run
    _xhd_zip(res)
    res["log_text"] = "\n".join(lines)
    res["tax_todo"] = [{"mst": it["mst"], "name": it.get("inv_name") or it["company"], "policy": it["policy"]}
                       for it in (res["requests"] or {}).get("items", []) if it["action"] == X.ACT_COMPANY]
    return res


def _xhd_zip(res):
    out = os.path.join(res["run"], "output")
    buf = io.BytesIO()                                            # giu thu muc con GCN\, TaxID\ trong zip
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for p in res["outputs"]:
            if os.path.isfile(p):
                z.write(p, os.path.relpath(p, out))
    res["zip"] = buf.getvalue()


# Ban in tra cuu MST tren trang Cuc Thue (tu doc captcha, doc khong ra thi nguoi dung go).
# Engine: Invoice-Issue\tool\taxsite.py
def xhd_tax_session():
    """Phien tra cuu (taxsite.Lookup): tu gui lai ma khi bo can bang tai cua trang chuyen sang may khac."""
    return T.Lookup()


def _captcha_view(png):
    """Anh captcha cua trang co nen trong suot, chu den: o dark mode gan nhu khong thay.
    Lot nen trang + phong to x2 de de doc. Loi thi tra nguyen anh."""
    try:
        from PIL import Image
        im = Image.open(io.BytesIO(png)).convert("RGBA")
        bg = Image.new("RGBA", im.size, (255, 255, 255, 255))
        bg.alpha_composite(im)
        bg = bg.convert("RGB").resize((im.width * 2, im.height * 2), Image.LANCZOS)
        out = io.BytesIO()
        bg.save(out, "PNG")
        return out.getvalue()
    except Exception:
        return png


def xhd_tax_captcha(lk, mst):
    try:
        return _captcha_view(lk.captcha(mst))
    except T.TaxSiteError as e:
        raise RunError(str(e)) from None
    except Exception as e:
        raise RunError(f"Cannot reach the tax site ({type(e).__name__}).") from None


def xhd_tax_auto(lk, res, mst):
    """Nhu xhd_tax_search nhung ddddocr tu doc captcha. Doc mai sai: r['wrong_code'] + anh moi de nguoi go."""
    return xhd_tax_search(lk, res, mst, None)


def xhd_tax_search(lk, res, mst, code):
    """Gui captcha cho 1 MST (code None = tu doc). Dung: Edge in trang ket qua ra TaxID\\TaxID_<mst>.pdf,
    them vao outputs + zip. Sai: tra kem anh captcha moi (r['captcha'])."""
    out = os.path.join(res["run"], "output", "TaxID")
    try:
        r = lk.auto(mst, out) if code is None else lk.search(mst, code, out)
    except T.TaxSiteError as e:
        raise RunError(str(e)) from None
    except Exception as e:
        raise RunError(f"Cannot reach the tax site ({type(e).__name__}).") from None
    if r.get("captcha"):
        r["captcha"] = _captcha_view(r["captcha"])
    if r["pdf"] and r["pdf"] not in res["outputs"]:
        res["outputs"].append(r["pdf"])
        _xhd_zip(res)
    return r
