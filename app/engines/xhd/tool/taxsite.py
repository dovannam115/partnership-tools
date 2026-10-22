# -*- coding: utf-8 -*-
"""Tra MST tren trang Cuc Thue (tracuunnt.gdt.gov.vn), gui form, in trang ket qua ra PDF (Edge headless)
lam ban dinh kem ho so HD cong ty.
Captcha: Lookup.auto() tu doc bang ddddocr (giong tool tra HD TGDD ben web Claim, them 08/10/2026 chieu);
doc mai khong ra thi web quay ve cho NGUOI DUNG go (Lookup.search). Captcha 5 ky tu chu thuong + so,
nen trong suot, khong nhieu: do 08/10 ddddocr dung ~1/2-1/3 so anh.

Luong cua trang (do 08/10/2026):
  GET mstdn.jsp -> GET captcha.png?uid= (anh gan voi phien JSESSIONID) -> POST mstdn.jsp:
  cm=cm, mst, fullname, address, cmt, captcha. Sai ma -> "Vui long nhap dung ma xac nhan!".

Bay (da do bang requests VA Edge that, ket qua nhu nhau):
  - Trang co ~4 may chu (duoi JSESSIONID :19ciqlppd/:19ciqlrd2/:19ciqll0q/:19ciqlo0c), bo can bang tai chia
    ngau nhien TUNG yeu cau, khong giu phien. Captcha chi dung tren may da phat ra no; POST roi sang may khac
    -> may do cap phien moi (Set-Cookie JSESSIONID moi) -> go dung van bao sai. Ca trinh duyet cung bi
    (ly do tra tay hay bao sai ma). Cach chua: nho cookie phien cua captcha; POST ma bi doi phien thi khoi
    phuc cookie cu va gui lai CUNG ma do (toi da TRIES lan) toi khi roi dung may. Do 08/10: 3/5 luot ve dung
    may trong 6 lan, 1 luot can toi lan 10. Goi day qua (vai chuc POST lien tuc) thi trang ngat ket noi.
  - Cookie TS0151a243 dat cho domain '.tracuunnt.gdt.gov.vn': http.cookiejar khong gui lai -> chep sang host.
  - SSL thieu chung chi trung gian: Python bao CERTIFICATE_VERIFY_FAILED -> dung kho chung chi Windows (truststore).
  - Gui sat nhau (< 1 giay) bi 429 -> giu cach MIN_GAP giay.
"""
import datetime as dt
import html as _html
import io
import os
import re
import shutil
import ssl
import subprocess
import tempfile
import threading
import time

import requests
from requests.adapters import HTTPAdapter

BASE = "https://tracuunnt.gdt.gov.vn/tcnnt/"
PAGE = BASE + "mstdn.jsp"
HOST = "tracuunnt.gdt.gov.vn"
MIN_GAP = 2.5
TRIES = 10
CAP_LEN = 5                 # captcha chuan 5 ky tu
AUTO_ROUNDS = 6             # so anh captcha toi da cho 1 MST truoc khi nho nguoi go
EDGE = [r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"]
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0 Safari/537.36"


class TaxSiteError(Exception):
    pass


_ocr = None
_ocr_lock = threading.Lock()


def _load_ocr():
    """Nap ddddocr 1 lan, ghim onnxruntime ve 1 luong (cung may voi web BHXH/Claim/BNO, nhu invoice_engine)."""
    global _ocr
    if _ocr is not None:
        return _ocr
    try:
        import onnxruntime as ort
        ort.set_default_logger_severity(3)
        orig = ort.InferenceSession.__init__

        def one_thread(self, *a, **k):
            so = k.get("sess_options") or ort.SessionOptions()
            so.intra_op_num_threads = 1
            so.inter_op_num_threads = 1
            k["sess_options"] = so
            return orig(self, *a, **k)

        ort.InferenceSession.__init__ = one_thread
    except Exception:
        pass
    try:
        import ddddocr
    except ImportError:
        raise TaxSiteError("Auto captcha needs ddddocr (py -3.12 -m pip install ddddocr).") from None
    _ocr = ddddocr.DdddOcr(show_ad=False)
    return _ocr


def solve(png):
    """Doc captcha: lot nen trang (anh nen trong suot), giu chu thuong + so."""
    from PIL import Image
    im = Image.open(io.BytesIO(png)).convert("RGBA")
    bg = Image.new("RGBA", im.size, (255, 255, 255, 255))
    bg.alpha_composite(im)
    buf = io.BytesIO()
    bg.convert("RGB").save(buf, "PNG")
    ocr = _load_ocr()
    with _ocr_lock:                                   # model khong thread-safe
        return re.sub(r"[^a-z0-9]", "", (ocr.classification(buf.getvalue()) or "").lower())


class _WinTrust(HTTPAdapter):
    def init_poolmanager(self, *a, **k):
        import truststore
        k["ssl_context"] = truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
        return super().init_poolmanager(*a, **k)


class Lookup:
    """Mot phien tra cuu, dung lai cho nhieu MST (moi lan tra 1 captcha moi)."""

    def __init__(self):
        self.s = requests.Session()
        self.s.mount("https://", _WinTrust())
        self.s.headers["User-Agent"] = UA
        self.last = 0.0
        self.snap = None          # cookie cua phien da phat captcha dang hien

    def close(self):
        self.s.close()

    def _sync(self):
        jar = self.s.cookies
        for c in list(jar):
            if c.domain.startswith("."):
                jar.set(c.name, c.value, domain=c.domain.lstrip("."), path=c.path or "/")

    def _sid(self):
        return next((c.value for c in self.s.cookies if c.name == "JSESSIONID"), None)

    def _gap(self):
        wait = MIN_GAP - (time.time() - self.last)
        if wait > 0:
            time.sleep(wait)

    def _req(self, method, url, **kw):
        self._gap()
        r = self.s.request(method, url, timeout=30, **kw)
        self._sync()
        self.last = time.time()
        if r.status_code == 429:
            raise TaxSiteError("The tax site says too many requests, wait a few seconds and try again.")
        r.raise_for_status()
        return r

    def captcha(self, mst=None):
        """Anh captcha (PNG) cho lan tra tiep theo; nho cookie cua phien phat ra anh."""
        if self._sid() is None:
            self._req("GET", PAGE)
        r = self._req("GET", BASE + "captcha.png?uid=")
        if not r.content.startswith(b"\x89PNG"):
            raise TaxSiteError("The tax site did not return a captcha image.")
        self.snap = {c.name: c.value for c in self.s.cookies}
        return r.content

    def _restore(self):
        for k, v in (self.snap or {}).items():
            self.s.cookies.set(k, v, domain=HOST, path="/")

    def search(self, mst, code, pdf_dir=None):
        """Gui MST + captcha nguoi dung go. POST roi sang may khac (doi JSESSIONID) thi gui lai cung ma voi
        phien cu. Tra ve dict: ok, wrong_code, not_found, rows, text, html, pdf, tries, captcha (anh moi khi sai)."""
        want = (self.snap or {}).get("JSESSIONID")
        data = {"cm": "cm", "mst": mst, "fullname": "", "address": "", "cmt": "", "captcha": code.strip()}
        r, tries = None, 0
        for tries in range(1, TRIES + 1):
            self._restore()
            try:
                r = self._req("POST", PAGE, data=data, headers={"Referer": PAGE, "Origin": "https://" + HOST})
            except requests.RequestException:          # trang hay ngat ket noi khi bi goi day: coi nhu 1 lan hong
                self.last = time.time()
                continue
            if want is None or self._sid() == want:     # roi dung may giu captcha
                break
        if r is None:
            raise TaxSiteError("The tax site is not answering, wait a minute and try again.")
        page = r.text
        text = _plain(page)
        wrong = "nhập đúng mã xác nhận" in text.lower()
        rows = [] if wrong else _rows_with(page, mst)
        nf = (not wrong) and not rows and "không tìm thấy" in text.lower()
        res = {"ok": bool(rows), "wrong_code": wrong, "not_found": nf, "rows": rows, "text": text, "html": page,
               "pdf": None, "tries": tries, "lost": want is not None and self._sid() != want}
        if wrong:
            res["captcha"] = self.captcha(mst)
        elif pdf_dir:
            res["pdf"], _ = save_result(page, mst, pdf_dir)
        return res

    def auto(self, mst, pdf_dir=None, rounds=AUTO_ROUNDS):
        """Tra 1 MST, captcha do ddddocr doc. Ket qua nhu search() them 'rounds' (so anh da dung).
        Het `rounds` van sai: wrong_code=True kem 'captcha' (anh moi cung phien) de nguoi go tiep."""
        img, bad = self.captcha(mst), 0
        for n in range(1, rounds + 1):
            code = solve(img)
            bad = bad + 1 if len(code) != CAP_LEN else 0
            # doc hut 3 anh lien thi chap nhan 4-6 ky tu de khoi ket
            if len(code) != CAP_LEN and not (bad >= 3 and CAP_LEN - 1 <= len(code) <= CAP_LEN + 1):
                img = self.captcha(mst)
                continue
            res = self.search(mst, code, pdf_dir)
            res["rounds"] = n
            if not res["wrong_code"]:
                return res
            img = res["captcha"]
        return {"ok": False, "wrong_code": True, "not_found": False, "rows": [], "text": "", "html": "",
                "pdf": None, "tries": 0, "lost": False, "rounds": rounds, "captcha": img}


def _plain(page):
    t = re.sub(r"<script.*?</script>|<style.*?</style>", " ", page, flags=re.S | re.I)
    t = re.sub(r"<[^>]+>", " ", t)
    return re.sub(r"\s+", " ", _html.unescape(t)).strip()


def _rows_with(page, mst):
    """Cac dong bang ket qua co chua MST (moi dong = list chu tung o)."""
    out = []
    for tr in re.findall(r"<tr[^>]*>(.*?)</tr>", page, flags=re.S | re.I):
        cells = [_plain(c) for c in re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", tr, flags=re.S | re.I)]
        if any(mst in c for c in cells) and len(cells) >= 3:
            out.append(cells)
    # tra MST me ra ca chi nhanh (-001, -026...): dong dung MST len dau
    return sorted(out, key=lambda r: mst not in r)


def save_result(page, mst, out_dir, when=None):
    """In trang ket qua ra PDF (nhu bam In, phong dang lam tay vay) bang Edge headless.
    Trang luu kem <base> de nap CSS/anh cua trang that; them 1 dong nguon + gio tra o dau trang vi
    ban in tu file luu khong co URL/gio o le giay; bo anh captcha (nap lai se ra ma khac).
    Tra ve (duong dan pdf | None, duong dan html)."""
    os.makedirs(out_dir, exist_ok=True)
    when = when or dt.datetime.now()
    base = f"TaxID_{re.sub(r'[^0-9-]', '', mst)}"
    note = ('<div style="font:11px Arial;color:#555;margin:4px 8px">Nguồn: %s - tra cứu lúc %s</div>'
            % (PAGE, when.strftime("%H:%M %d/%m/%Y")))
    page = re.sub(r"(<head[^>]*>)", r'\1<base href="%s"><meta charset="utf-8">' % BASE, page, count=1, flags=re.I)
    page = re.sub(r"(<body[^>]*>)", lambda m: m.group(1) + note, page, count=1, flags=re.I)
    page = re.sub(r"<img[^>]*captcha\.png[^>]*>", "", page, flags=re.I)
    # script theo doi EMC chen nut "Da ket noi EMC" fixed goc duoi -> PDF du 1 trang xam trong
    page = re.sub(r"<script[^>]*>(?:(?!</script>).)*_govaq.*?</script>", "", page, flags=re.S | re.I)
    hp = os.path.join(out_dir, base + ".html")
    with open(hp, "w", encoding="utf-8") as fh:
        fh.write(page)
    pdf = os.path.join(out_dir, base + ".pdf")
    edge = next((p for p in EDGE if os.path.isfile(p)), None)
    if not edge:
        return None, hp
    prof = tempfile.mkdtemp(prefix="edge_pdf_")
    try:
        subprocess.run([edge, "--headless=new", "--disable-gpu", f"--user-data-dir={prof}", "--no-pdf-header-footer",
                        f"--print-to-pdf={pdf}", "file:///" + hp.replace("\\", "/")],
                       timeout=90, capture_output=True, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except Exception:
        pass
    finally:
        shutil.rmtree(prof, ignore_errors=True)
    return (pdf if os.path.isfile(pdf) else None), hp
