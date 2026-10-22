# One run of one tool, started by server.py as a child process:
#   python -u worker.py <tool> <job dir>          a run: reads args.json, writes result.json
#   python -u worker.py --inspect <kind> <file>   a quick look at uploaded files: prints one JSON line
# Lines it prints go to the run panel in the window. Markers the server reads:
#   [[P]] done|total|label     progress bar
#   [[ERROR]] text             the message the user sees when the run fails
import os, sys, json, time, traceback, datetime, decimal

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import common as C

try: sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
except Exception: pass

DIR = ""


def say(*a):
    print(" ".join(str(x) for x in a), flush=True)


def prog(done, total, label=""):
    print("[[P]] %d|%d|%s" % (done, total, str(label).replace("\n", " ")[:120]), flush=True)


class UserError(Exception):
    """A failure the user can fix (wrong file, missing type...): shown without a traceback."""


def jsonable(v):
    if v is None or isinstance(v, (bool, int, str)): return v
    if isinstance(v, float): return None if v != v else v
    if isinstance(v, decimal.Decimal): return float(v)
    if isinstance(v, (datetime.datetime, datetime.date)): return v.strftime("%d/%m/%Y")
    if isinstance(v, (bytes, bytearray)): return None
    return str(v)


def table(title, cols, rows, views=None, limit=2000):
    recs = [r if isinstance(r, dict) else dict(zip(cols, r)) for r in rows]
    out = {"title": title, "cols": cols, "total": len(recs), "links": [],
           "rows": [[jsonable(r.get(c)) for c in cols] for r in recs[:limit]]}
    if views:
        out["views"] = [{"label": lb, "rows": [i for i, r in enumerate(recs[:limit]) if f(r)]} for lb, f in views]
    return out


def stat(label, value, tone=""):
    return {"label": label, "value": value, "tone": tone}


def day(s):
    """'2026-10-22' (date box) or '22/10/2026' -> date; '' -> None."""
    if not s: return None
    for f in ("%Y-%m-%d", "%d/%m/%Y"):
        try: return datetime.datetime.strptime(s, f).date()
        except ValueError: pass
    raise UserError("Bad date: %s" % s)


def files_of(paths):
    return [(os.path.basename(p), open(p, "rb").read()) for p in paths]


def write(path, data):
    with open(path, "wb") as f: f.write(data)
    return path


def runners():
    """Import the engines once, with their run folders inside this run (no cleanup of other runs)."""
    os.environ["PTT_NO_OPEN"] = "1"; os.environ["PTT_NO_ASK"] = "1"
    import runners as R
    R.RUNS = DIR or os.path.join(C.DUMP, "inspect")
    R._cleanup = lambda kind: None
    return R


def pbar(p, m):
    prog(min(int(p), 100), 100, m)


# ------------------------------------------------------------------ D2C VNPAY
def run_d2c(a):
    R = runners()
    say("Reading %d file(s) ..." % len(a["files"]))
    try:
        res = R.d2c_run(files_of(a["files"]), log=say, progress=pbar)
    except R.RunError as e:
        raise UserError(str(e))
    s, rc = res["stats"], res["recalc"]
    warns = []
    if rc is None:
        warns.append("Formulas were not recalculated (Excel not found). Open the file in Excel once to refresh them.")
    elif rc.get("total_errors"):
        warns.append("%d formula errors in the file. Check before sending." % rc["total_errors"])
    return {"stats": [stat("Successful transactions", s["successful"]), stat("Matched to Admin", s["matched"], "good"),
                      stat("With policy no.", s["policy"])],
            "files": [res["output"]], "warns": warns,
            "notes": ["Period %s · %s transactions in the VNPAY file · invoice no. found for %s"
                      % (res["period"], format(s["transactions"], ","), format(s["invoice"], ","))],
            "details": "Sheet Cảnh báo lists rows to check, sheet Treo lists orders with no policy yet: "
                       "re-check them next month.\n\n" + (res["log_text"] or "")}


# ------------------------------------------------------------------ PTT slips (GSM)
def run_ptt(a):
    R = runners()
    try:
        prev = R.ptt_preview(files_of(a["files"]))
        tu, den, lap = day(a.get("tu")) or prev["auto_tu"], day(a.get("den")) or prev["auto_den"], day(a.get("lap")) or datetime.date.today()
        if den <= tu: raise UserError("The Ngày cuối has to be after the Ngày đầu.")
        if not R.ptt_covered(prev["per_start"], tu, den): raise UserError("No policy starts inside this period in these files.")
        say("Making slips %s → %s ..." % (tu.strftime("%d/%m/%Y"), den.strftime("%d/%m/%Y")))
        res = R.ptt_run(prev, tu, den, lap)
    except R.RunError as e:
        raise UserError(str(e))
    outs = [p for p in res["outputs"] if os.path.isfile(p)]
    docs = [p for p in outs if p.endswith(".docx")]
    files, more = list(docs), []
    if len(docs) > 1:
        files, more = [write(os.path.join(DIR, "PTT_%s.zip" % datetime.date.today().strftime("%Y%m%d")), R.zip_bytes(outs))], docs
    rows = [[s["slip"], s["policies"], s["premium"], s["vat"], s["total"]] for s in res["slips"]]
    return {"stats": [stat("Slips", len(docs)), stat("Policies", sum(s["policies"] for s in res["slips"]))],
            "files": files, "more": more, "warns": res["warns"],
            "tables": [table("Slips", ["Slip", "Policies", "Premium", "VAT", "Total"], rows)] if rows else [],
            "details": res["log_text"]}


# ------------------------------------------------------------------ Invoice PDFs
def run_inv(a):
    R = runners()
    path = a["file"]
    try:
        scan = R.inv_scan(os.path.basename(path), open(path, "rb").read())
        items, dup = R.inv_unique(scan["items"], visible_only=bool(a.get("visible_only")))
        if not items: raise UserError("No invoice link to get.")
        say("%d invoice(s) to get, %d duplicate link(s) removed." % (len(items), dup))
        res = R.inv_run(items, progress=pbar)
    except R.RunError as e:
        raise UserError(str(e))
    files = [res["merged"]] if res["merged"] else []
    more = []
    if res["zip"]:
        base = os.path.basename(res["merged"] or "Invoices.pdf")[:-4]
        more = [write(os.path.join(DIR, base + "_separate.zip"), res["zip"])]
    n = len(res["rows"])
    warns = ["%d link(s) failed, see the table." % len(res["failed"])] if res["failed"] else []
    return {"stats": [stat("PDFs merged", "%d / %d" % (res["ok"], n), "good" if res["ok"] == n else "warn"),
                      stat("Pages", res["pages"]), stat("Duplicates removed", dup)],
            "files": files, "more": more, "warns": warns,
            "notes": ["Sheet %s, column %s%s" % (scan["sheet"], scan["col"], " (%s)" % scan["header"] if scan["header"] else "")],
            "tables": [table("Invoices", ["No", "Row", "Invoice", "Status"], res["rows"],
                             [("All (%d)" % n, lambda r: True),
                              ("Failed (%d)" % len(res["failed"]), lambda r: r["Status"] != "OK")])]}


# ------------------------------------------------------------------ Invoice requests (XHD)
def run_xhd(a):
    R = runners()
    try:
        say("Reading files ...")
        prev = R.xhd_preview(files_of(a["files"]))
        tu, den = day(a.get("tu")), day(a.get("den"))
        if prev["per_day"] and not (tu and den):
            tu, den = prev["auto_tu"], prev["auto_den"]
        lap = day(a.get("lap")) or datetime.date.today()
        res = R.xhd_run(prev, tu, den, lap, online=a.get("online", True) is not False, log=say)
    except R.RunError as e:
        raise UserError(str(e))
    X = R.X
    rq, rows = res.get("requests"), []
    if rq:
        c = rq["count"]
        rows.append(["On request", rq["ptt"], None, "%d company, %d personal, %d not eligible, %d already invoiced"
                     % (c.get(X.ACT_COMPANY, 0), c.get(X.ACT_PERSONAL, 0), rq["ineligible"],
                        c.get(X.ACT_AUTO, 0) + c.get(X.ACT_DONE, 0))])
    for g in res["bulk"]:
        rows.append(["Bulk %s %s" % (g["product"], g["channel"]), g["policies"],
                     format(g["total"], ",").replace(",", "."), ""])
    zp = write(os.path.join(DIR, "XHD_%s.zip" % datetime.date.today().strftime("%Y%m%d")), res["zip"])
    out = os.path.join(res["run"], "output")
    singles = [p for p in res["outputs"] if os.path.isfile(p)
               and os.path.basename(os.path.dirname(p)) not in ("GCN", "TaxID") and not p.endswith("run-log.txt")]
    todo = []
    for t in res.get("tax_todo") or []:
        if t["mst"] not in [x["mst"] for x in todo]: todo.append(t)
    return {"stats": [stat("Slips on request", rq["ptt"] if rq else 0),
                      stat("Bulk slips", len(res["bulk"])), stat("Files", len([p for p in res["outputs"] if os.path.isfile(p)]))],
            "files": [zp], "more": singles, "warns": res["warns"],
            "tables": [table("Slips", ["Slips", "Policies", "Total", "Note"], rows)] if rows else [],
            "details": res["log_text"],
            "tax": {"out": out, "zip": zp, "todo": todo} if todo else None}


# ------------------------------------------------------------------ KBank certificates
def run_kbank(a):
    R = runners()
    K = R.K
    ngay = day(a.get("date")) or datetime.date.today()
    merge = bool(a.get("merge"))
    say("Reading the masterfile ...")
    try:
        res = K.run(open(a["file"], "rb").read(), a.get("password") or "", ngay, merge=merge)
    except K.KbankError as e:
        raise UserError(str(e))
    if res.merged:
        name, data = res.merged
        files = [write(os.path.join(DIR, name), data)]
        n = len(res.rows)
    else:
        files = [write(os.path.join(DIR, "GCNBH KBank %s.zip" % ngay.strftime("%Y-%m-%d")), K.zip_bytes(res.files))]
        n = len(res.files)
    rows = [{"STT": r.stt, "Số GCNBH": r.so_gcn, "Người được bảo hiểm": r.ten, "CMND/CCCD": r.cccd,
             "Số tiền BH": K._money(r.stbh), "Phí": K._money(r.phi), "Từ": K._date(r.tu), "Đến": K._date(r.den)}
            for r in res.rows]
    return {"stats": [stat("Certificates", n, "good"), stat("Rows", len(res.rows))],
            "files": files, "warns": res.warnings,
            "tables": [table("Check before sending", list(rows[0].keys()) if rows else ["STT"], rows)]}


RUN = {"d2c": run_d2c, "ptt": run_ptt, "inv": run_inv, "xhd": run_xhd, "kbank": run_kbank}


# ------------------------------------------------------------------ quick looks for the form
def inspect(kind, a):
    """Read the uploaded files the way the run will, so the form can show types, the suggested
    period and counts before the user presses the button."""
    global DIR
    DIR = os.path.join(C.DUMP, "inspect", os.urandom(4).hex())
    os.makedirs(DIR, exist_ok=True)
    R = runners()
    try:
        if kind == "d2c":
            names = [os.path.basename(p) for p in a["files"]]
            block, warn = R.d2c_problems(names)
            return {"files": [{"name": n, "type": R.D2C_LABEL.get(R.d2c_kind(n), "Not used")} for n in names],
                    "block": block, "warn": warn}
        if kind == "ptt":
            p = R.ptt_preview(files_of(a["files"]))
            return {"tu": p["auto_tu"].isoformat(), "den": p["auto_den"].isoformat(),
                    "per_start": [[d.isoformat(), n] for d, n in sorted(p["per_start"].items())],
                    "warns": p["warns"], "counts": p["counts"]}
        if kind == "xhd":
            p = R.xhd_preview(files_of(a["files"]))
            names = {"requests": "Request list", "invoices": "Invoice ledger", "pias": "PIAS export"}
            return {"files": [{"name": n, "type": names.get(k, "Not recognised"), "kind": k} for n, k in p["files"]],
                    "tu": p["auto_tu"].isoformat() if p["auto_tu"] else "",
                    "den": p["auto_den"].isoformat() if p["auto_den"] else "",
                    "per_day": len(p["per_day"]), "n_req": p["n_req"], "warns": p["warns"]}
        if kind == "inv":
            path = a["files"][0]
            s = R.inv_scan(os.path.basename(path), open(path, "rb").read())
            all_, dup = R.inv_unique(s["items"], False)
            vis, dupv = R.inv_unique(s["items"], True)
            return {"sheet": s["sheet"], "col": s["col"], "header": s["header"], "links": len(s["items"]),
                    "hidden": sum(i["hidden"] for i in s["items"]),
                    "n": len(all_), "dup": dup, "n_vis": len(vis), "dup_vis": dupv}
        return {"error": "unknown"}
    except R.RunError as e:
        return {"error": str(e)}
    finally:
        import shutil
        shutil.rmtree(DIR, ignore_errors=True)


def main():
    global DIR
    if sys.argv[1] == "--inspect":
        a = json.load(open(sys.argv[3], encoding="utf-8"))
        try: out = inspect(sys.argv[2], a)
        except Exception as e: out = {"error": "%s: %s" % (type(e).__name__, e)}
        print("[[JSON]] " + json.dumps(out, ensure_ascii=False, default=jsonable), flush=True)
        return
    tool, DIR = sys.argv[1], sys.argv[2]
    a = json.load(open(os.path.join(DIR, "args.json"), encoding="utf-8"))
    res = {"ok": False}
    try:
        res = RUN[tool](a) or {}
        res["ok"] = True
        rc = 0
    except UserError as e:
        print("[[ERROR]] " + str(e), flush=True); res["error"] = str(e); rc = 2
    except Exception as e:
        tb = traceback.format_exc()
        say(tb)
        msg = "Unexpected error: %s" % e
        print("[[ERROR]] " + msg.replace("\n", " "), flush=True)
        res["error"] = msg; res["trace"] = tb; rc = 2
    for k in ("files", "more"):
        res[k] = [{"name": os.path.basename(p), "path": p} for p in res.get(k, []) if p and os.path.exists(p)]
    json.dump(res, open(os.path.join(DIR, "result.json"), "w", encoding="utf-8"), ensure_ascii=False, default=jsonable)
    sys.exit(rc)


if __name__ == "__main__":
    main()
