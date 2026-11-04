"""Premium collection slip (PTT) generator for the GSM partner programmes.

Double-click Run_PTT.bat -> pick the PIAS exports in the file dialog; Ctrl+click
to take the VCND export and the TNCN export in one go (dropping the files onto
the .bat skips the dialog). Every run writes its own output\\<date time>.

    py -3.13 tool\\generate_ptt.py [vcnd.xlsx tncn.xlsx] [--ngay-lap 24/08/2026]
                                   [--tu 20/08/2026] [--den 23/08/2026]

A slip covers 22h01 <Ngay dau> -> 22h00 <Ngay cuoi> and collects the policies
whose col `Ngay dau` starts inside it - one Monday run now covers Friday to
Sunday. The tool asks for those two dates; --tu / --den skip the question.
"""
import argparse
import datetime as dt
import glob
import io
import json
import os
import re
import sys
import time
from decimal import Decimal, ROUND_HALF_UP

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

# Ten file/sheet co dau tieng Viet: chay khong qua .bat (thieu PYTHONUTF8) thi
# codepage 1252 lam chuong trinh chet giua chung. Ep UTF-8, cho thay ky tu la.
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

import docx                                                  # noqa: E402
from openpyxl import load_workbook                           # noqa: E402

from docx_fill import replace_everywhere                     # noqa: E402

ROOT = os.path.dirname(HERE)
OUT_ROOT = os.path.join(ROOT, "output")
LAST_DIR = os.path.join(HERE, "last_folder.txt")

# thanh tien trinh: do rong, va thoi gian toi thieu moi buoc (giay). Dien mot
# phieu chi mat ~0,1s va ca lan chay chi ~5 buoc nen nhip phai cham hon tool
# GCNBH (0,06) thi thanh moi chay thay duoc. De 0 la het toc do.
BAR_WIDTH = 30
PACE = 0.30
LINE_W = 96

POLICY_RE = re.compile(r"^\d{2}/\d{2}/\d{2}/[A-Z]{2,6}/[A-Za-z]?\d+$")

# Do rong toi da cua ky goi y san. Ky tuan chi vai ngay; file TNCN tai nham
# ban dump ca nam thi trai rong hang thang, luc do goi y ca ky la mot cai Enter
# gom nham hang tram don - qua nguong nay chi goi y dung mot ngay ban cuoi.
MAX_AUTO_SPAN = 14

# ten cot lay theo TEN TIEU DE trong file export PIAS, khong theo vi tri
FIELDS = {
    "VCND": {
        "policy": ["số đơn bh/sđbs"],
        "ma_sp": ["mã sp"],
        "issue": ["ngày cấp đơn"],
        "start": ["ngày đầu"],
        "end": ["ngày cuối"],
        "fee": ["phí bảo hiểm(vnđ)", "phí bảo hiểm(vnd)"],
        "vat": ["nguyên tệ (vat)"],
        "status": ["mã t.thái"],
        "partner": ["tên đơn vị tgbh"],
    },
    "TNCN": {
        "policy": ["số đơn bh"],
        "ma_sp": ["mã sp"],
        "issue": ["ngày cấp đơn"],
        "start": ["ngày đầu"],
        "end": ["ngày cuối"],
        "fee": ["số tiền phí"],
        "status": ["tình trạng đơn"],
        "partner": ["tên đơn vị tgbh"],
        "mtn": ["mtn-gtbh (ntệ)"],
    },
}


class ToolError(Exception):
    pass


# ------------------------------------------------------------ thanh tien trinh

def _enable_ansi():
    """Bat ma mau cho cua so cmd (Windows 10 tro len). Khong bat duoc thi chay
    khong mau chu khong loi."""
    if not sys.stdout.isatty():
        return False
    try:
        import ctypes
        k = ctypes.windll.kernel32
        h = k.GetStdHandle(-11)
        mode = ctypes.c_uint32()
        if k.GetConsoleMode(h, ctypes.byref(mode)):
            k.SetConsoleMode(h, mode.value | 0x0004)   # VIRTUAL_TERMINAL_PROCESSING
            return True
    except Exception:
        pass
    return False


COLOR = _enable_ansi()
C_ON = "\033[96m" if COLOR else ""      # xanh lo
C_OK = "\033[92m" if COLOR else ""      # xanh la
C_OFF = "\033[0m" if COLOR else ""


class Progress:
    """Thanh chay tai cho: [######....]  3/9  <viec dang lam>."""

    def __init__(self, total):
        self.total = max(1, total)
        self.n = 0
        self.t0 = time.time()
        self.last = self.t0
        self.live = sys.stdout.isatty()

    def step(self, label=""):
        self.n += 1
        wait = PACE - (time.time() - self.last)
        if wait > 0:
            time.sleep(wait)
        self.last = time.time()
        if not self.live:
            return
        fill = int(round(BAR_WIDTH * min(self.n, self.total) / self.total))
        bar = "█" * fill + "░" * (BAR_WIDTH - fill)
        plain = "   %s  %*d/%d  %s" % (bar, len(str(self.total)), self.n,
                                       self.total, label)
        if len(plain) > LINE_W:
            plain = plain[:LINE_W - 3] + "..."
        sys.stdout.write("\r" + plain.replace(bar, C_ON + bar + C_OFF, 1)
                         + " " * (LINE_W - len(plain)))
        sys.stdout.flush()

    def note(self, text=""):
        """In mot dong giua luc thanh dang chay: xoa dong thanh, in xong thi
        buoc sau se ve lai thanh o dong moi. Khong lam vay thi print de
        thang len dong thanh dang do."""
        self.wipe()
        print(text)

    def wipe(self):
        if self.live:
            sys.stdout.write("\r" + " " * LINE_W + "\r")
            sys.stdout.flush()

    def finish(self, msg):
        """Giu lai vach day tren man hinh roi xuong dong (khong xoa) de con
        thay thanh da chay het."""
        if self.live:
            bar = "█" * BAR_WIDTH
            plain = "   %s  %d/%d" % (bar, self.total, self.total)
            sys.stdout.write("\r" + plain.replace(bar, C_ON + bar + C_OFF, 1)
                             + " " * max(0, LINE_W - len(plain)) + "\n")
            sys.stdout.flush()
        print("   %s>%s %s  (%.1fs)" % (C_OK, C_OFF, msg, time.time() - self.t0))


# ----------------------------------------------------------- hop thoai chon file

DIALOG_TITLE = "Select the PIAS export files (Ctrl+click for both)"
DIALOG_FILTER = "Excel files\0*.xlsx;*.xlsm;*.xls\0All files\0*.*\0\0"


def _remember(path):
    try:
        with io.open(LAST_DIR, "w", encoding="utf-8") as f:
            f.write(os.path.dirname(os.path.abspath(path)))
    except OSError:
        pass


def _start_dir():
    """Thu muc mo san trong hop thoai: lan truoc -> input\\ -> thu muc goc."""
    try:
        with io.open(LAST_DIR, encoding="utf-8") as f:
            d = f.read().strip()
        if d and os.path.isdir(d):
            return d
    except OSError:
        pass
    d = os.path.join(ROOT, "input")
    return d if os.path.isdir(d) else ROOT


def _ask_win32():
    """Hop thoai mo file cua Windows, goi thang qua comdlg32 (khong can tkinter,
    de ban Python nhung cung dung duoc)."""
    import ctypes
    from ctypes import wintypes

    class OPENFILENAMEW(ctypes.Structure):
        _fields_ = [
            ("lStructSize", wintypes.DWORD),
            ("hwndOwner", wintypes.HWND),
            ("hInstance", wintypes.HINSTANCE),
            ("lpstrFilter", wintypes.LPCWSTR),
            ("lpstrCustomFilter", wintypes.LPWSTR),
            ("nMaxCustFilter", wintypes.DWORD),
            ("nFilterIndex", wintypes.DWORD),
            ("lpstrFile", wintypes.LPWSTR),
            ("nMaxFile", wintypes.DWORD),
            ("lpstrFileTitle", wintypes.LPWSTR),
            ("nMaxFileTitle", wintypes.DWORD),
            ("lpstrInitialDir", wintypes.LPCWSTR),
            ("lpstrTitle", wintypes.LPCWSTR),
            ("Flags", wintypes.DWORD),
            ("nFileOffset", wintypes.WORD),
            ("nFileExtension", wintypes.WORD),
            ("lpstrDefExt", wintypes.LPCWSTR),
            ("lCustData", wintypes.LPARAM),
            ("lpfnHook", ctypes.c_void_p),
            ("lpTemplateName", wintypes.LPCWSTR),
            ("pvReserved", ctypes.c_void_p),
            ("dwReserved", wintypes.DWORD),
            ("FlagsEx", wintypes.DWORD),
        ]

    buf = ctypes.create_unicode_buffer(32768)
    ofn = OPENFILENAMEW()
    ofn.lStructSize = ctypes.sizeof(OPENFILENAMEW)
    ofn.lpstrFilter = DIALOG_FILTER
    ofn.lpstrFile = ctypes.cast(buf, wintypes.LPWSTR)
    ofn.nMaxFile = len(buf)
    ofn.lpstrInitialDir = _start_dir()
    ofn.lpstrTitle = DIALOG_TITLE
    # them ALLOWMULTISELECT de chon ca file VCND lan file TNCN trong mot lan mo
    ofn.Flags = (0x00081000 | 0x00000800 | 0x00000008  # EXPLORER|FILEMUSTEXIST|PATHMUSTEXIST|NOCHANGEDIR
                 | 0x00000200)                         # ALLOWMULTISELECT
    if not ctypes.windll.comdlg32.GetOpenFileNameW(ctypes.byref(ofn)):
        return []      # bam Cancel
    # chon nhieu file thi buffer la "thu muc\0ten1\0ten2\0\0", chon mot file thi la
    # ca duong dan roi \0\0 -> phai doc ca vung buffer, buf.value cat mat phan sau.
    parts = []
    for s in buf[:].split("\0"):
        if not s:
            break
        parts.append(s)
    if len(parts) <= 1:
        return parts
    return [os.path.join(parts[0], n) for n in parts[1:]]


def _ask_tk():
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    ps = filedialog.askopenfilenames(
        parent=root, title=DIALOG_TITLE, initialdir=_start_dir(),
        filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")])
    root.destroy()
    return list(ps)


def ask_files():
    """Danh sach duong dan, [] neu bam Cancel, None neu khong mo duoc hop thoai."""
    for fn in (_ask_win32, _ask_tk):
        try:
            return fn()
        except Exception:
            continue
    return None


def pick_from_console():
    cands = []
    for d in (os.path.join(ROOT, "input"), ROOT):
        cands = [p for p in sorted(glob.glob(os.path.join(d, "*.xls*")))
                 if not os.path.basename(p).startswith("~$")]
        if cands:
            break
    if not cands:
        raise ToolError("No Excel file found. Put the exports in the 'input' "
                        "folder, or drop them onto Run_PTT.bat.")
    print("Select the export files:")
    for i, p in enumerate(cands, 1):
        print("   [%d] %s" % (i, os.path.relpath(p, ROOT)))
    try:
        s = input("   Numbers, comma separated (Enter = all): ").strip()
    except EOFError:
        s = ""
    if not s:
        return cands
    picked = []
    for part in s.replace(";", ",").split(","):
        try:
            p = cands[int(part.strip()) - 1]
        except (ValueError, IndexError):
            continue
        if p not in picked:
            picked.append(p)
    return picked or cands


def pick_inputs(args):
    """Keo tha file vao .bat thi dung dung nhung file do, khong thi mo hop thoai
    (chon duoc nhieu file mot luc). Tra ve danh sach duong dan tuyet doi."""
    got = args
    if not got:
        got = ask_files()
        if got is None:                    # khong mo duoc hop thoai nao
            got = pick_from_console()
        elif not got:
            raise ToolError("Cancelled - no file selected.")
    paths = []
    for arg in got:
        p = os.path.abspath(arg)
        if os.path.isdir(p):
            raise ToolError("That is a folder (%s). Drop the Excel files, or "
                            "double-click the .bat to get the file picker." % p)
        if not os.path.isfile(p):
            raise ToolError("File not found: %s" % p)
        if p not in paths:
            paths.append(p)
    if not paths:
        raise ToolError("No file selected.")
    _remember(paths[0])
    return paths


# --------------------------------------------------------------- tien ich nho

def norm(v):
    return re.sub(r"\s+", " ", str(v)).strip().lower() if v is not None else ""


def half_up(v):
    """Lam tron nua LEN nhu ke toan. round() cua Python lam tron nua CHAN nen
    16 483 454,5 ra ...454 - phieu tuan cong phi nhieu ngay hay roi dung .5,
    lech 1 dong so voi cot VAT cua PIAS."""
    return int(Decimal(str(v)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def money(v):
    return f"{half_up(v):,}".replace(",", " ")


def dmy(d):
    return d.strftime("%d/%m/%Y")


def parse_date(s):
    for f in ("%d/%m/%Y", "%d.%m.%Y", "%Y-%m-%d", "%d/%m/%y"):
        try:
            return dt.datetime.strptime(s, f)
        except ValueError:
            pass
    raise ToolError("Cannot read the date: %s (use dd/mm/yyyy)" % s)


def parse_typed_date(s, ref):
    """Ngay go tay o cua so lenh: nhan 18/08/2026, 18-8-2026, 18.8 hay 18/8 -
    thieu nam thi lay nam cua ngay goi y."""
    parts = [p for p in re.split(r"[/.\- ]+", s.strip()) if p]
    if len(parts) == 3 and len(parts[0]) == 4:          # go kieu 2026/08/18
        parts = parts[::-1]
    if len(parts) == 2:
        parts.append(str(ref.year))
    if len(parts) != 3:
        raise ToolError("Cannot read the date: %s (use dd/mm/yyyy)" % s)
    return parse_date("%s/%s/%s" % tuple(parts))


def can_ask():
    """Co ban phim de hoi khong: chay trong build/CI hay dat PTT_NO_ASK=1 thi khong."""
    if os.environ.get("PTT_NO_ASK"):
        return False
    try:
        return bool(sys.stdin) and sys.stdin.isatty()
    except Exception:
        return False


def ask_date(label, default, known=None):
    """Hoi mot ngay ngay tren cua so lenh, Enter la lay ngay goi y."""
    if not can_ask():
        return default
    for _ in range(3):
        try:
            s = input("   %-28s [%s] : " % (label, dmy(default))).strip()
        except (EOFError, KeyboardInterrupt):
            print("")
            return default
        if not s:
            return default
        try:
            d = parse_typed_date(s, default)
        except ToolError as exc:
            print("   %s" % exc)
            continue
        if known and d not in known:
            print("   No policy issued on %s in these files." % dmy(d))
            continue
        return d
    return default


def day_list(per_day, limit=8):
    """Cac ngay co trong file, kem so don moi ngay - in ra truoc khi hoi de
    nguoi lam phieu doi chieu duoc thu Sau/thu Bay/Chu nhat cua tuan."""
    days = sorted(per_day)
    one = lambda d: "%s (%d)" % (d.strftime("%d/%m"), per_day[d])   # noqa: E731
    if len(days) > limit:
        return "%s ... %s   [%d days in the files]" % (
            ", ".join(one(d) for d in days[:2]),
            ", ".join(one(d) for d in days[-3:]), len(days))
    return ", ".join(one(d) for d in days)


def ask_period(auto_tu, auto_den, per_start):
    """Hoi ky thu cua phieu - ngay dau va ngay cuoi, Enter la nhan ngay goi y.

    Tu 24/08/2026 phieu lam MOI TUAN MOT LAN vao thu Hai nen mot phieu gom
    nhieu ngay ban (thu Sau -> Chu nhat), khong con la mot ngay nhu truoc.
    Hai ngay nay doc thang o file export: NGAY DAU cua dong dau va NGAY CUOI
    cua dong cuoi - dung dong chu in tren phieu, 22h01 ngay <dau> -> 22h00
    ngay <cuoi>. Ngay le giua ky khong co don van nhan (03/08/2026 PIAS khong
    cap don nao), chi ky RONG hoan toan moi hoi lai.
    """
    if not can_ask():
        return auto_tu, auto_den
    print("   Ngày đầu in the files: %s" % day_list(per_start))
    tu, den = auto_tu, auto_den
    for _ in range(3):
        tu = ask_date("Period from (Ngày đầu)", tu)
        den = ask_date("Period to   (Ngày cuối)",
                       den if den > tu else tu + dt.timedelta(days=1))
        if den <= tu:
            print("   The Ngày cuối has to be after the Ngày đầu.")
        else:
            n = sum(c for d, c in per_start.items() if tu <= d < den)
            if n:
                print("   -> %d policy covered by 22h01 %s -> 22h00 %s"
                      % (n, dmy(tu), dmy(den)))
                return tu, den
            print("   No policy starts inside 22h01 %s -> 22h00 %s in these files."
                  % (dmy(tu), dmy(den)))
        tu, den = auto_tu, auto_den
    return auto_tu, auto_den


def sheet_kind(header):
    if "số đơn bh/sđbs" in header:
        return "VCND"
    if "số đơn bh" in header and ("mtn-gtbh (ntệ)" in header or "số tiền phí" in header):
        return "TNCN"
    return None


# --------------------------------------------------------------------- doc du lieu

def read_rows(sources, warns, bar):
    """Do het cac file da chon vao chung mot ro. Sheet nhan dang bang tieu de nen
    dua file nao truoc cung duoc; trung don (giua sheet hay giua file) bo qua."""
    out = {"VCND": [], "TNCN": []}
    seen = {"VCND": set(), "TNCN": set()}
    dup = 0
    for path in sources:
        bar.step("reading %s" % os.path.basename(path))
        wb = load_workbook(path, data_only=True)
        hit = 0
        for ws in wb:
            head_row, header = None, {}
            for r in ws.iter_rows(min_row=1, max_row=min(ws.max_row or 1, 10)):
                names = {norm(c.value): c.column for c in r if c.value is not None}
                if "mã sp" in names:
                    head_row, header = r[0].row, names
                    break
            if head_row is None:
                continue
            kind = sheet_kind(header)
            if kind is None:
                continue
            cols = {}
            for key, cands in FIELDS[kind].items():
                cols[key] = next((header[c] for c in cands if c in header), None)
            if not all(cols[k] for k in ("policy", "ma_sp", "fee", "start", "end", "issue")):
                warns.append("%s / sheet '%s': required columns missing, skipped."
                             % (os.path.basename(path), ws.title))
                continue
            hit += 1
            for r in ws.iter_rows(min_row=head_row + 1, max_row=ws.max_row):

                def cell(key):
                    c = cols.get(key)
                    return r[c - 1].value if c else None

                pol = str(cell("policy") or "").strip()
                if not POLICY_RE.match(pol):
                    continue
                key = (pol, str(cell("ma_sp")))
                if key in seen[kind]:
                    dup += 1
                    continue
                seen[kind].add(key)
                out[kind].append({
                    "file": os.path.basename(path),
                    "sheet": ws.title,
                    "policy": pol,
                    "ma_sp": str(cell("ma_sp") or "").strip(),
                    "issue": cell("issue"),
                    "start": cell("start"),
                    "end": cell("end"),
                    "fee": float(cell("fee") or 0),
                    "vat": float(cell("vat")) if cell("vat") not in (None, "") else None,
                    "status": str(cell("status") or "").strip(),
                    "partner": str(cell("partner") or "").strip(),
                    "mtn": to_amount(cell("mtn")),
                })
        wb.close()
        if not hit:
            warns.append("%s: no PIAS policy sheet recognised, file ignored."
                         % os.path.basename(path))
    return out, dup


def to_amount(v):
    """So tien trong o Excel (so hoac chu '150,000,000'); trong/khong doc duoc -> None."""
    if v in (None, ""):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", "").replace(" ", ""))
    except ValueError:
        return None


def select_rows(data, prog):
    """Cac don len mot phieu: dung nguon, dung ma sp, dung khach hang va MTN neu co."""
    rows = [r for r in data[prog["source"]] if r["ma_sp"] in prog["ma_sp"]]
    need = prog.get("customer_contains")
    if need:
        rows = [r for r in rows if need.lower() in r["partner"].lower()]
    mtn = prog.get("mtn_equals")
    if mtn is not None:
        rows = [r for r in rows if r.get("mtn") == float(mtn)]
    return rows


# --------------------------------------------------------------------------- chay

def suggest_period(data, progs, keep):
    """Ky goi y (auto_tu, auto_den) + so don theo cot Ngay dau. Dung chung cho main
    (hoi tren cua so lenh) va web Partnership Tools (dien san vao o ngay)."""

    def status_ok(row):
        return not keep or not row["status"] or row["status"] in keep

    # Ky thu cua phieu la cua so 22h01 NGAY DAU -> 22h00 NGAY CUOI, va don len
    # phieu la don co cot Ngay dau nam trong cua so do (nua mo: dau <= x < cuoi).
    # Do theo cot Ngay dau chu khong theo Ngay cap don, dung nhu nguoi lam phieu
    # doc file. Mot phep thu tren Ngay dau lo duoc ca hai phia: don Tainan hieu
    # luc 7 ngay nen thu theo Ngay cuoi la rung sach, con don cap dem hom cuoi
    # tuan sau (Ngay dau = dung ngay cuoi ky nay) thi phai rot ra.
    # Chi dem don CO THE len phieu. Don tai nan cua CBNV (phi 20.000d) nam
    # chung file TNCN nhung khong thuoc chuong trinh nao: dem vao thi bang ngay
    # in ra co nhung ngay khong bao gio dung, nguoi doc tuong minh mat don.
    on_slip = {id(r) for p in progs for r in select_rows(data, p)}
    per_start = {}
    for rows in data.values():
        for r in rows:
            if r["start"] and status_ok(r) and id(r) in on_slip:
                per_start[r["start"]] = per_start.get(r["start"], 0) + 1
    if not per_start:
        raise ToolError("No policy of these programmes has a Ngày đầu - check "
                        "the export files.")

    # Ngay cuoi goi y lay o cot Ngay cuoi cua don VCND: don nay hieu luc dung
    # mot ngay nen ket o dung cuoi ky. Don TNCN keo them mot tuan, khong dung
    # lam moc duoc, khong co file VCND thi suy tu Ngay dau muon nhat + 1.
    ends = [r["end"] for r in data.get("VCND", [])
            if r["end"] and status_ok(r) and id(r) in on_slip]
    auto_den = max(ends) if ends else max(per_start) + dt.timedelta(days=1)
    auto_tu = min(per_start)
    if (auto_den - auto_tu).days > MAX_AUTO_SPAN:
        auto_tu = auto_den - dt.timedelta(days=1)
    return auto_tu, auto_den, per_start


def main(argv):
    ap = argparse.ArgumentParser(add_help=True)
    ap.add_argument("excel", nargs="*", help="PIAS export files (.xlsx)")
    ap.add_argument("--ngay-lap", help="slip date, default today")
    ap.add_argument("--tu", help="period start = col Ngay dau, skips the question")
    ap.add_argument("--den", help="period end = col Ngay cuoi, skips the question")
    a = ap.parse_args(argv)

    print("=== PTT generator - GSM partner ===")
    with open(os.path.join(ROOT, "config.json"), encoding="utf-8") as f:
        cfg = json.load(f)

    srcs = pick_inputs(a.excel)
    print("Data   : %s" % os.path.basename(srcs[0]))
    for p in srcs[1:]:
        print("         %s" % os.path.basename(p))

    warns, log = [], []
    progs = cfg["programs"]
    print("")
    # cac buoc: moi file 1 + moi phieu 1 + 1 ghi log
    bar = Progress(len(srcs) + len(progs) + 1)
    data, dup = read_rows(srcs, warns, bar)
    if not any(data.values()):
        raise ToolError("No policy row found - check the export files.")

    keep = cfg.get("status_include") or []

    def status_ok(row):
        return not keep or not row["status"] or row["status"] in keep

    auto_tu, auto_den, per_start = suggest_period(data, progs, keep)

    if a.tu or a.den:                                  # go san o dong lenh
        tu = parse_date(a.tu) if a.tu else auto_tu
        den = parse_date(a.den) if a.den else auto_den
    else:
        # Ky may doan CHUA CHAC la tuan minh can: don Tainan cap deu moi ngay,
        # tai file muon mot nhip la don cua tuan sau da nam trong do. Nen bay
        # ca hai dau ra hoi lai, Enter la nhan ngay may doan.
        bar.wipe()
        tu, den = ask_period(auto_tu, auto_den, per_start)
        if can_ask():
            print("")
    if den <= tu:
        raise ToolError("--den (%s) has to be after --tu (%s)." % (dmy(den), dmy(tu)))

    outside = 0
    for kind in list(data):
        inside = []
        for r in data[kind]:
            if r["start"] and not (tu <= r["start"] < den):
                outside += 1
            else:
                inside.append(r)
        data[kind] = inside
    if not any(data.values()):
        raise ToolError("No policy covered by 22h01 %s -> 22h00 %s in these files - "
                        "set the period with --tu dd/mm/yyyy --den dd/mm/yyyy."
                        % (dmy(tu), dmy(den)))

    for kind in list(data):
        good = []
        for r in data[kind]:
            if not status_ok(r):
                warns.append("policy %s skipped (status %s is not in %s)"
                             % (r["policy"], r["status"], keep))
            else:
                good.append(r)
        data[kind] = good
    kept = sum(len(v) for v in data.values())

    lap = parse_date(a.ngay_lap) if a.ngay_lap else dt.datetime.today()
    bar.note("Period : 22h01 %s -> 22h00 %s   |   slip date %s"
             % (dmy(tu), dmy(den), dmy(lap)))
    bar.note("Batch  : %d policy kept, %d row(s) outside the period"
             % (kept, outside))

    now = dt.datetime.now()
    base_dir = os.path.join(OUT_ROOT, now.strftime("%Y-%m-%d %H%M"))
    run_dir, i = base_dir, 1
    while os.path.isdir(run_dir):
        i += 1
        run_dir = "%s (%d)" % (base_dir, i)
    os.makedirs(run_dir)

    made = []
    for prog in progs:
        rows = select_rows(data, prog)
        if not rows:
            bar.step("%s: no policy" % prog["key"])
            warns.append("%s: no policy in this period, no slip produced." % prog["key"])
            continue

        fee = sum(r["fee"] for r in rows)
        rate = prog.get("vat_rate", 0)
        vat = half_up(fee * rate)
        exported = [r["vat"] for r in rows if r["vat"] is not None]
        if rate and exported and half_up(sum(exported)) != vat:
            warns.append("%s: computed VAT %s differs from the export's VAT column %s."
                         % (prog["key"], money(vat), money(sum(exported))))
        hl_from, hl_to = min(r["start"] for r in rows), max(r["end"] for r in rows)
        # Han thanh toan: mac dinh 20 ngay sau ngay cuoi hieu luc don, nhung
        # tung chuong trinh co the doi goc tinh sang ngay cuoi ky (batch_end).
        # Don VCND hieu luc het ngay cuoi dot nen hai goc cho cung ket qua; don
        # TNCN keo dai them mot tuan, ke toan van lay ngay cuoi + 20 nhu VCND.
        due_base = prog.get("payment_due_base", cfg.get("payment_due_base", "policy_end"))
        base = hl_to if due_base == "policy_end" else den
        term = int(prog.get("payment_term_days", cfg.get("payment_term_days", 20)))
        han = base + dt.timedelta(days=term)

        name = prog["filename"].format(tu=tu.strftime("%d.%m"), den=den.strftime("%d.%m"))
        bar.step("%s  ->  %s" % (prog["key"], name))

        tpl = os.path.join(ROOT, prog["template"])
        if not os.path.isfile(tpl):
            raise ToolError("Template missing: %s" % tpl)
        doc = docx.Document(tpl)
        vals = {
            "CBKT": cfg.get("cbkt", ""),
            "GCNBH": "; ".join(sorted({r["policy"] for r in rows})),
            "HL": "%s - %s" % (dmy(hl_from), dmy(hl_to)),
            "HANTT": dmy(han),
            "PHI": money(fee),
            "VAT": money(vat),
            "TONG": money(fee + vat),
            "KYTU": dmy(tu),
            "KYDEN": dmy(den),
            "D": lap.strftime("%d"),
            "M": lap.strftime("%m"),
            "Y": lap.strftime("%Y"),
        }
        for k, v in vals.items():
            replace_everywhere(doc, re.escape("{{" + k + "}}"), v.replace("\\", "\\\\"))
        doc.save(os.path.join(run_dir, name))
        made.append((prog, rows, fee, vat, vals, name))
        log.append("%-16s | %d policy | %s | premium %s + VAT %s = %s | cover %s | due %s | %s"
                   % (prog["key"], len(rows), ", ".join(sorted({r["policy"] for r in rows})),
                      money(fee), money(vat), money(fee + vat), vals["HL"],
                      vals["HANTT"], name))

    bar.step("writing log")
    with io.open(os.path.join(run_dir, "log.txt"), "w", encoding="utf-8") as f:
        f.write("Data   : %s\n" % srcs[0])
        for p in srcs[1:]:
            f.write("         %s\n" % p)
        f.write("Run    : %s\n" % now.strftime("%d/%m/%Y %H:%M"))
        f.write("Period : 22h01 %s -> 22h00 %s   |   slip date %s\n"
                % (dmy(tu), dmy(den), dmy(lap)))
        f.write("Batch  : %d policy kept, %d row(s) outside the period, "
                "%d duplicate row(s)\n\n" % (kept, outside, dup))
        f.write("\n".join(log))
        f.write("\n\nWarnings (%d):\n" % len(warns))
        f.write("\n".join(warns) if warns else "(none)")
    bar.finish("%d slip(s) done" % len(made))

    print("")
    for prog, rows, fee, vat, vals, name in made:
        print("[%s]  %d policy  |  premium %s + VAT %s = %s"
              % (prog["key"], len(rows), money(fee), money(vat), money(fee + vat)))
        print("   cover %s   |   due %s" % (vals["HL"], vals["HANTT"]))
        print("   -> %s" % name)

    if warns:
        print("")
        print("!! %d warning(s):" % len(warns))
        for w in warns[:12]:
            print("   - %s" % w)
        if len(warns) > 12:
            print("   ... %d more, see log.txt" % (len(warns) - 12))

    print("")
    print("Output folder:")
    print("   %s" % run_dir)
    if not os.environ.get("PTT_NO_OPEN"):                    # build/test dat bien nay
        try:
            os.startfile(run_dir)                            # noqa: S606
        except Exception:
            pass


if __name__ == "__main__":
    try:
        main(sys.argv[1:])
    except ToolError as e:
        if sys.stdout.isatty():
            sys.stdout.write("\r" + " " * LINE_W + "\r")
        print("")
        print("ERROR: %s" % e)
        sys.exit(1)
