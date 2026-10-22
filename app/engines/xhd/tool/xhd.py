# -*- coding: utf-8 -*-
"""Engine xuat hoa don (XHD) cho phong Phat trien doi tac, KHONG dung streamlit (web goi qua runners).

Hai luong, cung mot bo file dau vao (nhan loai file theo dong tieu de, ten file tuy y):
  - File yeu cau XHD (export "Danh sach yeu cau ho tro - Xuat hoa don")
  - File HOA DON THU PHI (so cai hoa don, co cot Ky hieu HD: DG = he thong tu xuat, BA = phong xuat tay)
  - File PIAS cap don (DLCN, TSTH...), co the nhieu file

1) THEO YEU CAU: moi yeu cau -> tra so cai hoa don. Da co hoa don thi bo qua (DG tu xuat / BA da xuat).
   Chua co -> 1 phieu PTT: cong ty (doi chieu ten theo MST, sua typo) hoac ca nhan.
   Don xe (AUTO/MOTO) ma nguoi mua KHAC ten cong ty -> khong du dieu kien, lap PTT ca nhan cho nguoi mua
   + dua vao danh sach "khong du dieu kien" (can bo sung hop dong).
   Don xe: tach phi tu file GCNBH (PDF o cot Link GCNBH), vi file yeu cau de phi = 0.
2) GOM KHACH LE: don PIAS chua co trong so cai hoa don, KHONG nam trong danh sach yeu cau -> gom theo
   san pham + kenh (DLCN: Ma loai TS 030615 Online / 030612 Rabbit Care), trong ky ngay cap don nguoi dung chon.
   Ra 1 PTT + 1 Excel danh sach dinh kem cho moi nhom.

Luat suy tu bo de 06/10/2026 (reference\\), da doi chieu khop 4 PTT gom mau (xem CLAUDE.md).
"""
import copy
import datetime as dt
import io
import json
import os
import re
import unicodedata
from concurrent.futures import ThreadPoolExecutor
from decimal import Decimal, ROUND_HALF_UP
from difflib import SequenceMatcher

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
TEMPLATE = os.path.join(ROOT, "templates", "PTT_Template.docx")
CONFIG = os.path.join(ROOT, "config.json")

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


class ToolError(Exception):
    pass


def load_config(path=CONFIG):
    with io.open(path, encoding="utf-8") as fh:
        return json.load(fh)


# ============================================================================ helpers

def nfc(s):
    return unicodedata.normalize("NFC", s)


def txt(v):
    if v is None or (isinstance(v, float) and pd.isna(v)) or v is pd.NaT:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return nfc(str(v)).strip()


def money(v):
    """'1,063,200' / '1.063.200' / 454000.0 -> int (VND khong co so le)."""
    s = txt(v)
    if not s:
        return 0
    neg = s.startswith("-")
    d = re.sub(r"[^\d]", "", s)
    return (-int(d) if neg else int(d)) if d else 0


def half_up(x):
    return int(Decimal(str(x)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def vnd(n):
    """454000 -> '454.000' (cach ghi tren phieu mau 30/09). 0/None -> ''."""
    if not n:
        return ""
    return f"{int(n):,}".replace(",", ".")


def to_date(v):
    if v is None or v == "" or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    s = txt(v)
    for fmt in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def dmy(d):
    return d.strftime("%d/%m/%Y") if d else ""


def strip_accents(s):
    s = unicodedata.normalize("NFD", s.replace("đ", "d").replace("Đ", "D"))
    return "".join(c for c in s if unicodedata.category(c) != "Mn")


# viet tat ten doanh nghiep -> dang ngan chung, de so ten khach go voi ten theo MST
_ABBR = [
    (r"\bcong ty\b", "cty"), (r"\bct\b", "cty"),
    (r"\btrach nhiem huu han\b", "tnhh"),
    (r"\bmot thanh vien\b", "mtv"),
    (r"\bco phan\b", "cp"), (r"\bctcp\b", "cty cp"),
    (r"\bthuong mai\b", "tm"), (r"\bdich vu\b", "dv"), (r"\btmdv\b", "tm dv"),
    (r"\bxuat nhap khau\b", "xnk"), (r"\bsan xuat\b", "sx"), (r"\bdau tu\b", "dt"),
    (r"\bxay dung\b", "xd"), (r"\bthanh pho\b", "tp"), (r"\bviet nam\b", "vn"),
]


def name_key(s):
    s = strip_accents(txt(s)).lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    for pat, rep in _ABBR:
        s = re.sub(pat, rep, s)
    return re.sub(r"\s+", " ", s).strip()


def similar(a, b):
    ka, kb = name_key(a), name_key(b)
    if not ka or not kb:
        return 0.0
    if ka == kb:
        return 1.0
    return SequenceMatcher(None, ka, kb).ratio()


_ORG = re.compile(r"\b(cty|tnhh|cp|jsc|ltd|llc|co ltd|corp|corporation|company|group|international|inc|"
                  r"truong|dai hoc|hoc vien|vien|benh vien|hop tac xa|htx|dntn|chi nhanh|van phong|"
                  r"ngan hang|tong cty|ubnd|uy ban|so|trung tam|hiep hoi|doanh nghiep|tap doan)\b")


def looks_org(name):
    return bool(_ORG.search(name_key(name)))


def clean_mst(v, warns=None, who=""):
    """MST doanh nghiep 10 so (chi nhanh 10-3), CCCD 12 so. Excel hay an so 0 dau -> bu lai."""
    s = re.sub(r"[\s.]", "", txt(v))
    m = re.fullmatch(r"(\d+)(-\d{3})?", s)
    if not m:
        return s
    body = m.group(1)
    if len(body) in (9, 11):
        body = "0" + body
        if warns is not None:
            warns.append(f"{who}: tax ID/ID number lost its leading 0 in Excel, restored to {body}{m.group(2) or ''}.")
    return body + (m.group(2) or "")


def product_of(policy):
    p = txt(policy).split("/")
    return p[3] if len(p) > 3 else ""


def unit_of(policy):
    p = txt(policy).split("/")
    return p[2] if len(p) > 2 else ""


def period_label(a, b):
    if a == b:
        return a.strftime("%d.%m.%Y")
    if a.year == b.year and a.month == b.month:
        return f"{a:%d}-{b:%d.%m.%Y}"
    if a.year == b.year:
        return f"{a:%d.%m}-{b:%d.%m.%Y}"
    return f"{a:%d.%m.%Y}-{b:%d.%m.%Y}"


def period_tag(a, b):
    return f"{a:%d%m}-{b:%d%m}" if a != b else f"{a:%d%m}"


def safe(s):
    return re.sub(r"[^A-Za-z0-9._-]+", "_", strip_accents(s)).strip("_")


# ============================================================================ doc file

KINDS = {
    "requests": ("Mã SV / Mã yêu cầu", "PVI Policy GCN"),
    "invoices": ("Số ctừ", "Ký hiệu HĐ", "Số đơn Bh"),
    "pias": ("Số đơn BH/SĐBS", "Mã sp", "Phí bảo hiểm(VNĐ)"),
}
KIND_LABEL = {"requests": "Invoice requests", "invoices": "Invoice ledger (HÓA ĐƠN THU PHÍ)", "pias": "PIAS policy export"}


def _header(path):
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True)
    try:
        ws = wb.worksheets[0]
        for row in ws.iter_rows(min_row=1, max_row=1, values_only=True):
            return [txt(v) for v in row]
        return []
    finally:
        wb.close()


def detect(path):
    try:
        head = set(_header(path))
    except Exception:
        return None
    for k, cols in KINDS.items():
        if all(c in head for c in cols):
            return k
    return None


def _read(path):
    df = pd.read_excel(path, sheet_name=0, dtype=object)
    df.columns = [txt(c) for c in df.columns]
    return df


def load_inputs(paths, warns):
    """paths -> dict(requests=df|None, invoices=df|None, pias=df|None, files=[(name, kind)])."""
    got = {"requests": [], "invoices": [], "pias": []}
    files = []
    for p in paths:
        k = detect(p)
        files.append((os.path.basename(p), k))
        if k:
            got[k].append(p)
        else:
            warns.append(f"{os.path.basename(p)}: not recognised (not a request list, invoice ledger or PIAS export), skipped.")
    out = {"files": files}
    for k in ("requests", "invoices"):
        if len(got[k]) > 1:
            raise ToolError(f"Upload only one {KIND_LABEL[k]} file ({len(got[k])} found).")
        out[k] = _read(got[k][0]) if got[k] else None
    if got["pias"]:
        out["pias"] = pd.concat([_read(p).assign(_src=os.path.basename(p)) for p in got["pias"]], ignore_index=True)
    else:
        out["pias"] = None
    return out


# ============================================================================ so cai hoa don

def invoice_index(inv):
    """{so don: {"net": tong phi rong, "docs": [(so ctu, ky hieu, ngay)], "auto": co dong DG duong,
                 "positives": so chung tu duong}}"""
    idx = {}
    if inv is None:
        return idx
    cols = list(inv.columns)
    ip, ic, ik, idt, ifee = (cols.index("Số đơn Bh"), cols.index("Số ctừ"), cols.index("Ký hiệu HĐ"),
                             cols.index("Ngày ctừ"), cols.index("Tiền phí"))
    for row in inv.itertuples(index=False, name=None):
        pol = txt(row[ip])
        if not pol:
            continue
        fee = money(row[ifee])
        e = idx.setdefault(pol, {"net": 0, "docs": [], "auto": False, "pos": set(), "neg": 0})
        e["net"] += fee
        doc = (txt(row[ic]), txt(row[ik]), to_date(row[idt]))
        if doc not in e["docs"]:
            e["docs"].append(doc)
        if fee > 0:
            e["pos"].add(doc[0])
            if doc[1] == "DG":
                e["auto"] = True
        elif fee < 0:
            e["neg"] += 1
    return idx


def invoiced(e):
    return e is not None and e["net"] > 0


def doc_text(e):
    if not e:
        return ""
    return ", ".join(f"{c} ({k}, {dmy(d)})" for c, k, d in e["docs"])


def ledger_range(inv):
    if inv is None:
        return None, None
    ds = [to_date(v) for v in inv["Ngày ctừ"].tolist()]
    ds = [d for d in ds if d]
    return (min(ds), max(ds)) if ds else (None, None)


# ============================================================================ tra cuu MST

def lookup_tax(msts, cfg, timeout=12, workers=4):
    """{mst: {"name", "address", "status"} | {"error": "..."} } qua API cong khai (du lieu Cuc Thue).
    Chi tra MST doanh nghiep (10 hoac 10-3 so). Loi mang -> {"error"} de bao 'not checked'."""
    import requests
    url = cfg.get("tax_api") or "https://api.vietqr.io/v2/business/{mst}"
    todo = sorted({m for m in msts if re.fullmatch(r"\d{10}(-\d{3})?", m or "")})
    out = {}

    def one(m):
        last = None
        for _ in range(2):
            try:
                r = requests.get(url.format(mst=m), timeout=timeout,
                                 headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"})
                j = r.json()
                if j.get("code") == "00" and j.get("data"):
                    d = j["data"]
                    return m, {"name": nfc(d.get("name") or ""), "address": nfc(d.get("address") or ""),
                               "status": nfc(d.get("status") or "")}
                return m, {"error": "not found", "desc": j.get("desc", "")}
            except Exception as e:
                last = e
        return m, {"error": f"lookup failed ({type(last).__name__})"}

    if todo:
        with ThreadPoolExecutor(workers) as ex:
            for m, res in ex.map(one, todo):
                out[m] = res
    return out


# ============================================================================ GCNBH (PDF) cho don xe

def fetch_gcn(url, timeout=45):
    import requests
    r = requests.get(url, timeout=timeout, headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"})
    r.raise_for_status()
    if r.content[:4] != b"%PDF":
        raise ToolError("the link did not return a PDF")
    return r.content


def parse_gcn(pdf_bytes):
    """Doc phi tu GCNBH xe: phi chinh + VAT, phi bo sung (gom VAT), tong.
    Tra ve dict a1, v1, a2, v2, total. Phi bo sung tach VAT 10%: net = lam tron(gom/1.1), VAT = gom - net
    (khop phieu mau: 60.000 -> 54.545 + 5.455, 50.000 -> 45.455 + 4.545)."""
    from pypdf import PdfReader
    rd = PdfReader(io.BytesIO(pdf_bytes))
    text = nfc(" ".join((p.extract_text() or "") for p in rd.pages[:2]))
    t = re.sub(r"\s+", "", text)
    num = r"([\d.,]+)"
    m1 = re.search(r"Phíbảohiểm" + num + r"đồng", t)
    mv = re.search(r"VAT" + num + r"đồng", t)
    m2 = re.search(r"Phíbảohiểm\(gồmVAT\)" + num + r"đồng", t)
    mt = re.search(r"TỔNGPHÍBẢOHIỂM" + num + r"ĐỒNG", t, re.I)
    if not (m1 and mt):
        raise ToolError("fee lines not found in the GCNBH")
    a1, v1 = money(m1.group(1)), money(mv.group(1)) if mv else 0
    gross2 = money(m2.group(1)) if m2 else 0
    a2 = half_up(Decimal(gross2) / Decimal("1.1")) if gross2 else 0
    v2 = gross2 - a2
    total = money(mt.group(1))
    if a1 + v1 + gross2 != total:
        raise ToolError(f"GCNBH fees do not add up ({a1}+{v1}+{gross2} != {total})")
    return {"a1": a1, "v1": v1, "a2": a2, "v2": v2, "total": total}


# ============================================================================ 1) theo yeu cau

REQ_COLS = ["STT", "Mã SV / Mã yêu cầu", "Tên người liên hệ", "Số điện thoại", "Email", "Trạng thái", "Đối tác",
            "Tên sản phẩm", "Số GCNBH", "Link GCNBH", "PVI Serial", "PVI Policy GCN", "PVI Number",
            "Phí bảo hiểm chính", "Phí bảo hiểm phụ", "VAT", "Tổng tiền", "Tên NĐBH/Người mua bảo hiểm",
            "Tên công ty", "Mã số thuế", "Địa chỉ", "Ngày gửi"]

ACT_COMPANY = "PTT - company invoice"
ACT_PERSONAL = "PTT - personal invoice"
ACT_INELIGIBLE = "PTT - personal (not eligible for company)"
ACT_AUTO = "Skip - auto-invoiced (DG)"
ACT_DONE = "Skip - already invoiced"
ACT_MISSING = "Check - missing policy number"
ACT_DUP = "Skip - duplicate request"
MAKES_PTT = (ACT_COMPANY, ACT_PERSONAL, ACT_INELIGIBLE)


def plan_requests(inp, cfg, warns, tax=True, gcn=True, gcn_dir=None, log=print):
    """Tra ve list item (dict) cho moi dong yeu cau, theo thu tu Ngay gui tang dan."""
    req = inp.get("requests")
    if req is None:
        return []
    miss = [c for c in REQ_COLS if c not in req.columns]
    if miss:
        raise ToolError("Request list is missing column(s): " + ", ".join(miss))
    idx = invoice_index(inp.get("invoices"))
    lo, _hi = ledger_range(inp.get("invoices"))
    vehicles = set(cfg.get("vehicle_products") or [])
    partners = {k.upper(): v for k, v in (cfg.get("partners") or {}).items()}
    items, seen = [], set()
    for r in req.to_dict("records"):
        if not txt(r.get("STT")) and not txt(r.get("Mã SV / Mã yêu cầu")):
            continue
        who = f"Request {txt(r['Mã SV / Mã yêu cầu']) or txt(r['STT'])}"
        it = {
            "stt": txt(r["STT"]), "code": txt(r["Mã SV / Mã yêu cầu"]), "sent": to_date(r["Ngày gửi"]),
            "sent_raw": txt(r["Ngày gửi"]), "status": txt(r["Trạng thái"]),
            "partner": txt(r["Đối tác"]), "channel": partners.get(txt(r["Đối tác"]).upper(), txt(r["Đối tác"])),
            "product_name": txt(r["Tên sản phẩm"]), "gcn_no": txt(r["Số GCNBH"]), "gcn_link": txt(r["Link GCNBH"]),
            "serial": txt(r["PVI Serial"]), "policy": txt(r["PVI Policy GCN"]), "pvi_number": txt(r["PVI Number"]),
            "buyer": txt(r["Tên NĐBH/Người mua bảo hiểm"]), "company": txt(r["Tên công ty"]),
            "mst": clean_mst(r["Mã số thuế"], warns, who), "address": txt(r["Địa chỉ"]),
            "email": txt(r["Email"]), "phone": txt(r["Số điện thoại"]), "contact": txt(r["Tên người liên hệ"]),
            "req_a1": money(r["Phí bảo hiểm chính"]), "req_a2": money(r["Phí bảo hiểm phụ"]),
            "req_vat": money(r["VAT"]), "req_total": money(r["Tổng tiền"]),
            "notes": [], "a1": 0, "v1": 0, "a2": 0, "v2": 0,
        }
        it["product"] = product_of(it["policy"])
        it["vehicle"] = it["product"] in vehicles
        e = idx.get(it["policy"])
        it["ledger"] = doc_text(e)
        if not it["policy"]:
            it["action"] = ACT_MISSING
            it["notes"].append("No policy number in the request (PVI Policy GCN is empty).")
        elif it["policy"] in seen:
            it["action"] = ACT_DUP
        elif invoiced(e):
            it["action"] = ACT_AUTO if e["auto"] else ACT_DONE
            if len(e["pos"]) - e["neg"] > 1:
                it["notes"].append("Ledger shows more than one live invoice for this policy, check for a duplicate.")
        else:
            it["action"] = None                       # quyet dinh sau khi tra MST
            if lo and it["sent"] and it["sent"] < lo:
                it["notes"].append(f"Request is older than the invoice ledger (starts {dmy(lo)}), "
                                   "an invoice may exist outside the file.")
        if it["policy"]:
            seen.add(it["policy"])
        items.append(it)

    todo = [it for it in items if it["action"] is None]
    taxres = {}
    if tax and todo:
        log(f"Checking {len({it['mst'] for it in todo})} tax ID(s) online")
        taxres = lookup_tax([it["mst"] for it in todo], cfg)

    for it in todo:
        t = taxres.get(it["mst"]) or {}
        it["tax_name"], it["tax_addr"], it["tax_status"] = t.get("name", ""), t.get("address", ""), t.get("status", "")
        it["tax_err"] = t.get("error", "") if tax else "not checked"
        is_biz_id = bool(re.fullmatch(r"\d{10}(-\d{3})?", it["mst"]))
        company = bool(it["company"]) and is_biz_id and (
            looks_org(it["company"]) or bool(it["tax_name"])
            or similar(it["company"], it["buyer"]) < cfg.get("same_entity_min", 0.85))
        it["inv_type"] = "Company" if company else "Personal"
        it["name_check"] = ""
        it["inv_name"] = it["company"] if company else (it["company"] if it["company"] and not looks_org(it["company"]) else it["buyer"])
        if company:
            if it["tax_name"]:
                sc = similar(it["company"], it["tax_name"])
                if sc >= 0.999:
                    it["name_check"] = "OK"
                    it["inv_name"] = it["tax_name"]
                elif sc >= cfg.get("name_fix_min", 0.6):
                    it["name_check"] = "Fixed by tax ID"
                    it["inv_name"] = it["tax_name"]
                    it["notes"].append(f"Company name corrected from tax ID: '{it['company']}' -> '{it['tax_name']}'.")
                else:
                    it["name_check"] = "MISMATCH"
                    it["notes"].append(f"Tax ID {it['mst']} belongs to '{it['tax_name']}', not '{it['company']}'. "
                                       "Check the tax ID with the customer before issuing.")
                if it["tax_status"] and "đang hoạt động" not in it["tax_status"].lower():
                    it["notes"].append(f"Tax ID status: {it['tax_status']}.")
            else:
                it["name_check"] = "Not checked" if it["tax_err"] != "not found" else "Tax ID not found"
                if it["tax_err"] == "not found":
                    it["notes"].append(f"Tax ID {it['mst']} not found in the tax data, check it before issuing.")
        same = (similar(it["buyer"], it["company"]) >= cfg.get("same_entity_min", 0.85)
                or (it["tax_name"] and similar(it["buyer"], it["tax_name"]) >= cfg.get("same_entity_min", 0.85)))
        if company and it["vehicle"] and not same:
            it["action"] = ACT_INELIGIBLE
            it["inv_type"] = "Personal"
            it["inv_name"] = it["buyer"]
            it["notes"].append("Vehicle policy bought by a person other than the invoice company: needs a contract "
                               "(e.g. vehicle lease) to invoice the company, personal invoice made for the buyer.")
        else:
            it["action"] = ACT_COMPANY if company else ACT_PERSONAL
            if it["vehicle"]:
                it["notes"].append("Vehicle policy in the invoice name is normally auto-invoiced (DG) on the issue day: "
                                   "make sure the invoice ledger is up to date before issuing, or it will be invoiced twice.")

    # phi: don thuong lay tu file yeu cau, don xe doc GCNBH
    for it in items:
        if it["action"] not in MAKES_PTT:
            continue
        pdf = None
        if gcn and it["gcn_link"].lower().startswith("http"):
            try:
                pdf = fetch_gcn(it["gcn_link"])
                if gcn_dir:
                    os.makedirs(gcn_dir, exist_ok=True)
                    with open(os.path.join(gcn_dir, f"GCN_{safe(it['serial'] or it['policy'])}.pdf"), "wb") as fh:
                        fh.write(pdf)
                it["gcn_file"] = True
            except Exception as ex:
                it["notes"].append(f"GCNBH download failed ({ex}).")
        pc = (cfg["products"].get(it["product"]) or {})
        if it["vehicle"] or (not it["req_a1"] and not it["req_a2"]):
            fees = None
            if pdf is not None:
                try:
                    fees = parse_gcn(pdf)
                except Exception as ex:
                    it["notes"].append(f"Could not read fees from the GCNBH ({ex}).")
            if fees:
                it.update({k: fees[k] for k in ("a1", "v1", "a2", "v2")})
            else:
                it["notes"].append(f"Fill the fee lines by hand, total {vnd(it['req_total'])}.")
        else:
            it["a1"], it["a2"] = it["req_a1"], it["req_a2"]
            if pc.get("request_vat_row") == 1 or (not pc.get("request_vat_row") and not it["req_a2"]):
                it["v1"] = it["req_vat"]
            else:
                it["v2"] = it["req_vat"]
        tot = it["a1"] + it["v1"] + it["a2"] + it["v2"]
        if it["req_total"] and tot and tot != it["req_total"]:
            it["notes"].append(f"Fee lines add up to {vnd(tot)} but the request total is {vnd(it['req_total'])}.")
    items.sort(key=lambda it: (it["sent"] or dt.date.max, it["stt"]))
    return items


# ============================================================================ 2) gom khach le

def plan_bulk(inp, cfg, warns, held=None):
    """Don PIAS chua co hoa don va khong nam trong danh sach yeu cau.
    Tra ve DataFrame (moi dong 1 dong phi) cot: policy, product, channel, issued, start, end, code, si, fee,
    rate, row, vat_rate, vat, buyer, serial, source, pname."""
    pias = inp.get("pias")
    if pias is None:
        return None
    need = ["Số đơn BH/SĐBS", "Mã sp", "Phí bảo hiểm(VNĐ)", "Ngày cấp đơn", "Mã t.thái"]
    miss = [c for c in need if c not in pias.columns]
    if miss:
        raise ToolError("PIAS export is missing column(s): " + ", ".join(miss))
    idx = invoice_index(inp.get("invoices"))
    req = inp.get("requests")
    req_pols = set(txt(v) for v in req["PVI Policy GCN"].tolist()) if req is not None else set()
    req_pols.discard("")
    keep_status = set(cfg.get("status_include") or [])
    units = set(cfg.get("units") or [])
    prods = {k: v for k, v in cfg["products"].items() if v.get("bulk")}

    def col(r, c):
        return r.get(c) if c in r else None

    rows, other_unit, other_prod, other_code, bad_status = [], 0, set(), {}, 0
    for r in pias.to_dict("records"):
        pol = txt(r["Số đơn BH/SĐBS"])
        if not pol or pol.count("/") < 3:
            continue
        if keep_status and txt(r["Mã t.thái"]) not in keep_status:
            bad_status += 1
            continue
        prod = product_of(pol)
        if prod not in prods:
            other_prod.add(prod)
            continue
        if units and unit_of(pol) not in units:
            other_unit += 1
            continue
        pc = prods[prod]
        code = txt(r["Mã sp"])
        rate = txt(col(r, "Tỷ lệ phí BH(%)"))
        if code not in pc["codes"]:
            other_code[code] = other_code.get(code, 0) + 1
            continue
        rule = pc["codes"][code]
        if pc.get("accident_rate") is not None and rate:
            try:
                if abs(float(rate) - float(pc["accident_rate"])) < 1e-9:
                    rule = pc["accident_row"]
            except ValueError:
                pass
        if pc.get("channel_col"):
            ccode = txt(col(r, pc["channel_col"]))
            channel = pc["channels"].get(ccode) or f"Other {ccode or 'blank'}"
        else:
            channel = pc.get("channel_fixed", "Online")
        rows.append({
            "policy": pol, "product": prod, "pname": pc["name"], "channel": channel,
            "issued": to_date(r["Ngày cấp đơn"]), "start": to_date(col(r, "Ngày đầu")), "end": to_date(col(r, "Ngày cuối")),
            "code": code, "si": money(col(r, "Số tiền BH")), "fee": money(r["Phí bảo hiểm(VNĐ)"]), "rate": rate,
            "row": rule["row"], "vat_rate": rule["vat"], "buyer": txt(col(r, "Tên đơn vị TGBH")),
            "serial": txt(col(r, "Số thẻ")), "source": txt(col(r, "Nguồn cấp đơn")), "file": txt(col(r, "_src")),
        })
    if bad_status:
        warns.append(f"PIAS: {bad_status} row(s) with a status other than {', '.join(sorted(keep_status))} left out.")
    if other_unit:
        warns.append(f"PIAS: {other_unit} row(s) of other units (not {', '.join(sorted(units))}) left out.")
    if other_prod - {""}:
        warns.append("PIAS: products without a bulk rule left out: " + ", ".join(sorted(other_prod - {""})) + ".")
    for c, n in other_code.items():
        warns.append(f"PIAS: {n} row(s) with product code {c} have no bulk rule, left out.")
    df = pd.DataFrame(rows)
    if df.empty:
        return df
    # export PIAS lap lai phi cua don theo tung nguoi duoc BH -> khu trung theo dong phi
    n0 = len(df)
    df = df.drop_duplicates(["policy", "code", "si", "fee", "rate"]).reset_index(drop=True)
    if len(df) < n0:
        warns.append(f"PIAS: {n0 - len(df)} repeated row(s) (one per insured person) counted once.")
    oth = df[df["channel"].str.startswith("Other")]
    if len(oth):
        warns.append(f"PIAS: {oth['policy'].nunique()} policy(ies) with an unknown channel code, grouped as "
                     + ", ".join(sorted(oth["channel"].unique())) + ".")
    df["invoiced"] = df["policy"].map(lambda p: invoiced(idx.get(p)))
    df["requested"] = df["policy"].isin(req_pols)
    df["vat"] = [half_up(Decimal(f) * Decimal(str(v))) for f, v in zip(df["fee"], df["vat_rate"])]
    if held is not None:
        held.extend(sorted(df.loc[~df["invoiced"] & df["requested"], "policy"].unique()))
    return df


def bulk_candidates(df):
    if df is None or df.empty:
        return df
    return df[~df["invoiced"] & ~df["requested"]].copy()


def suggest_period(cand):
    if cand is None or cand.empty:
        return None, None
    ds = [d for d in cand["issued"].tolist() if d]
    return (min(ds), max(ds)) if ds else (None, None)


def per_day(cand):
    if cand is None or cand.empty:
        return {}
    return cand.groupby("issued")["policy"].nunique().to_dict()


# ============================================================================ Word

def _fill_xml(el, values):
    from docx.oxml.ns import qn
    for t in el.iter(qn("w:t")):
        if t.text and "{{" in t.text:
            t.text = re.sub(r"\{\{(\w+)\}\}", lambda m: str(values.get(m.group(1), "")), t.text)
            if t.text != t.text.strip():
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def _restart_numbering(doc, els):
    """Muc 1. 2. 3. cua phieu dung list tu dong cua Word: phieu chep sau se danh tiep 4. 5. 6.
    -> moi phieu mot w:num moi (cung abstractNum, startOverride = 1)."""
    from docx.oxml import parse_xml
    from docx.oxml.ns import qn, nsdecls
    numbering = doc.part.numbering_part.element
    nums = {n.get(qn("w:numId")): n for n in numbering.findall(qn("w:num"))}
    used = {}
    for e in els:
        for nid in e.iter(qn("w:numId")):
            used.setdefault(nid.get(qn("w:val")), []).append(nid)
    for old, refs in used.items():
        src = nums.get(old)
        if src is None:
            continue
        new = str(max(int(k) for k in nums) + 1)
        absid = src.find(qn("w:abstractNumId")).get(qn("w:val"))
        el = parse_xml(f'<w:num {nsdecls("w")} w:numId="{new}"><w:abstractNumId w:val="{absid}"/>'
                       + "".join(f'<w:lvlOverride w:ilvl="{l}"><w:startOverride w:val="1"/></w:lvlOverride>'
                                 for l in range(9)) + "</w:num>")
        numbering.append(el)
        nums[new] = el
        for ref in refs:
            ref.set(qn("w:val"), new)


def write_ptt(slips, path):
    """slips = list dict gia tri placeholder. Moi phieu 1 trang, gop vao 1 file (giong file mau)."""
    import docx
    from docx.oxml.ns import qn
    doc = docx.Document(TEMPLATE)
    body = doc.element.body
    sect = body.find(qn("w:sectPr"))
    els = [e for e in body.iterchildren() if e is not sect]
    block, sep = els[:-1], els[-1]
    for e in els:
        body.remove(e)
    for i, vals in enumerate(slips):
        if i:
            body.insert(len(body) - 1, copy.deepcopy(sep))
        new = []
        for e in block:
            ne = copy.deepcopy(e)
            _fill_xml(ne, vals)
            body.insert(len(body) - 1, ne)
            new.append(ne)
        if i:
            _restart_numbering(doc, new)
    doc.save(path)


def slip_date_vals(d):
    return {"D": f"{d:%d}", "M": f"{d:%m}", "Y": f"{d:%Y}"} if d else {"D": "   ", "M": "   ", "Y": str(dt.date.today().year)}


def request_slip(it, cfg, slip_date):
    pc = cfg["products"].get(it["product"]) or {}
    labels = pc.get("labels") or cfg["default_labels"]
    company = it["action"] == ACT_COMPANY
    name = it["inv_name"] if company else (it["inv_name"] + cfg["personal_suffix"])
    email = "; ".join(x for x in (it["email"], cfg["cbkt_email"]) if x)
    tail = f"Seri số {it['serial']}" if it["serial"] else f"số {it['policy']}"
    tot = it["a1"] + it["v1"] + it["a2"] + it["v2"]
    v = {
        "CBKT": cfg["cbkt"], "GCN": it["policy"], "HANTT": dmy(it["sent"]) or it["sent_raw"],
        "TENKH": name, "DIACHI": it["address"], "MSTLBL": pc.get("mst_label", "Mã số thuế"),
        "MST": it["mst"] if company else "", "EMAIL": email, "DBH1": pc.get("dbh1", "ĐBH 1"),
        "L1": labels[0], "L2": labels[1], "LT": labels[2],
        "A1": vnd(it["a1"]), "V1": vnd(it["v1"]), "A2": vnd(it["a2"]), "V2": vnd(it["v2"]),
        "TOT": vnd(tot) or vnd(it["req_total"]), "NOTE": it["product_name"] or pc.get("name", ""),
        "CONTENT": f"Phí {it['product_name'] or pc.get('name', '')} cấp qua website {it['channel']} đơn bảo hiểm {tail}",
    }
    v.update(slip_date_vals(slip_date))
    return v


def bulk_slip(g, prod, channel, a, b, cfg, slip_date):
    pc = cfg["products"][prod]
    labels = pc.get("labels") or cfg["default_labels"]
    s = {}
    for row in (1, 2):
        part = g[g["row"] == row]
        s[f"A{row}"], s[f"V{row}"] = int(part["fee"].sum()), int(part["vat"].sum())
    tot = s["A1"] + s["V1"] + s["A2"] + s["V2"]
    v = {
        "CBKT": cfg["cbkt"], "GCN": "Theo danh sách đính kèm", "HANTT": "Theo danh sách đính kèm",
        "TENKH": cfg["retail_customer"], "DIACHI": "", "MSTLBL": pc.get("mst_label", "Mã số thuế"), "MST": "",
        "EMAIL": cfg["cbkt_email"], "DBH1": pc.get("dbh1", "ĐBH 1"),
        "L1": labels[0], "L2": labels[1], "LT": labels[2],
        "A1": vnd(s["A1"]), "V1": vnd(s["V1"]), "A2": vnd(s["A2"]), "V2": vnd(s["V2"]),
        "TOT": vnd(tot), "NOTE": pc["name"],
        "CONTENT": pc["content"].format(channel=channel, period=period_label(a, b)),
    }
    v.update(slip_date_vals(slip_date))
    return v, s, tot


# ============================================================================ Excel

def _xl_base():
    from openpyxl import Workbook
    from openpyxl.styles import Font
    wb = Workbook()
    ws = wb.active
    ws.sheet_view.showGridLines = False
    return wb, ws, Font


def _style_header(ws, row, ncol):
    from openpyxl.styles import Alignment, Border, Font, Side
    for c in range(1, ncol + 1):
        cl = ws.cell(row, c)
        cl.font = Font(name="Arial", size=10, bold=True)
        cl.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cl.border = Border(bottom=Side(style="thin", color="808080"))
    ws.row_dimensions[row].height = 30


def _font_all(ws, top, bottom, ncol, skip_rows=()):
    from openpyxl.styles import Font
    for r in range(top, bottom + 1):
        if r in skip_rows:
            continue
        for c in range(1, ncol + 1):
            cl = ws.cell(r, c)
            f = cl.font
            cl.font = Font(name="Arial", size=10, bold=f.bold, color=f.color)


def write_bulk_list(g, path, cfg):
    """Danh sach dinh kem PTT gom: 1 dong / dong phi, VAT va thanh tien la cong thuc, dong tong + bang
    doi chieu 2 dong phi cua phieu (SUMIF theo cot Dong PTT)."""
    from openpyxl.styles import Alignment, Border, Font, Side
    from openpyxl.utils import get_column_letter as L
    wb, ws, _ = _xl_base()
    ws.title = "Danh sách"
    heads = ["STT", "Số đơn BH/SĐBS", "Số thẻ", "Tên đơn vị TGBH", "Ngày cấp đơn", "Ngày đầu", "Ngày cuối",
             "Mã sp", "Số tiền BH", "Phí bảo hiểm (VNĐ)", "Thuế suất VAT", "VAT", "Thành tiền", "Dòng PTT", "Nguồn cấp đơn"]
    ws.cell(1, 1, cfg["retail_customer"]).font = Font(name="Arial", size=11, bold=True)
    for i, h in enumerate(heads, 1):
        ws.cell(2, i, h)
    _style_header(ws, 2, len(heads))
    g = g.sort_values(["issued", "policy", "row", "code"]).reset_index(drop=True)
    stt, last = 0, None
    for i, r in enumerate(g.itertuples(index=False), start=3):
        if r.policy != last:
            stt, last = stt + 1, r.policy
        vals = [stt, r.policy, r.serial, r.buyer, r.issued, r.start, r.end, r.code, r.si or None, r.fee,
                float(r.vat_rate), f"=ROUND(J{i}*K{i},0)", f"=J{i}+L{i}", int(r.row), r.source]
        for c, v in enumerate(vals, 1):
            ws.cell(i, c, v)
        for c in (5, 6, 7):
            ws.cell(i, c).number_format = "dd/mm/yyyy"
        for c in (9, 10, 12, 13):
            ws.cell(i, c).number_format = "#,##0;-#,##0;-"
        ws.cell(i, 11).number_format = "0%"
        ws.cell(i, 11).font = Font(name="Arial", size=10, color="0000FF")
    last_row = 2 + len(g)
    t = last_row + 1
    ws.cell(t, 2, "Tổng cộng")
    ws.cell(t, 3, f'=SUMPRODUCT(1/COUNTIF(B3:B{last_row},B3:B{last_row}))')
    ws.cell(t, 3).number_format = '0" đơn"'
    for c in (10, 12, 13):
        ws.cell(t, c, f"=SUM({L(c)}3:{L(c)}{last_row})")
        ws.cell(t, c).number_format = "#,##0;-#,##0;-"
    for c in range(1, len(heads) + 1):
        ws.cell(t, c).border = Border(top=Side(style="thin", color="808080"))
        ws.cell(t, c).font = Font(name="Arial", size=10, bold=True)
    pc = cfg["products"][g["product"].iloc[0]]
    labels = pc.get("labels") or cfg["default_labels"]
    s = t + 2
    ws.cell(s, 2, "Đối chiếu PTT").font = Font(name="Arial", size=10, bold=True)
    ws.cell(s, 10, "Số tiền").font = Font(name="Arial", size=10, bold=True)
    ws.cell(s, 12, "VAT").font = Font(name="Arial", size=10, bold=True)
    for k, row in enumerate((1, 2), start=1):
        rr = s + k
        ws.cell(rr, 2, labels[k - 1])
        ws.cell(rr, 10, f"=SUMIF(N3:N{last_row},{row},J3:J{last_row})")
        ws.cell(rr, 12, f"=SUMIF(N3:N{last_row},{row},L3:L{last_row})")
    rr = s + 3
    ws.cell(rr, 2, labels[2]).font = Font(name="Arial", size=10, bold=True)
    ws.cell(rr, 13, f"=J{s+1}+L{s+1}+J{s+2}+L{s+2}")
    ws.cell(rr, 13).font = Font(name="Arial", size=10, bold=True)
    for r_ in range(s + 1, rr + 1):
        for c in (10, 12, 13):
            ws.cell(r_, c).number_format = "#,##0;-#,##0;-"
    _font_all(ws, 3, rr, len(heads), skip_rows=(t, s, rr))
    for r_ in range(3, last_row + 1):
        ws.cell(r_, 11).font = Font(name="Arial", size=10, color="0000FF")
    widths = [6, 26, 12, 30, 12, 12, 12, 9, 15, 15, 9, 12, 15, 8, 12]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[L(i)].width = w
    for r_ in range(3, last_row + 1):
        ws.cell(r_, 1).alignment = Alignment(horizontal="center")
        ws.cell(r_, 14).alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:{L(len(heads))}{last_row}"
    wb.save(path)


def write_request_check(items, path):
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter as L
    wb, ws, _ = _xl_base()
    ws.title = "Yêu cầu"
    heads = ["STT", "Mã yêu cầu", "Ngày gửi", "Đối tác", "Sản phẩm", "Số GCNBH/SĐBS", "Seri", "Người mua",
             "Tên công ty (khách nhập)", "MST / CCCD", "Tên theo MST", "Tình trạng MST", "Kiểm tra tên",
             "Loại hóa đơn", "Tên trên PTT", "Hóa đơn đã có", "Xử lý", "Phí chính", "VAT phí chính",
             "Phí bổ sung", "VAT phí bổ sung", "Tổng PTT", "Tổng theo yêu cầu", "Chênh lệch", "Ghi chú"]
    for i, h in enumerate(heads, 1):
        ws.cell(1, i, h)
    _style_header(ws, 1, len(heads))
    order = {ACT_COMPANY: 0, ACT_PERSONAL: 1, ACT_INELIGIBLE: 2, ACT_MISSING: 3, ACT_DONE: 4, ACT_AUTO: 5, ACT_DUP: 6}
    rows = sorted(items, key=lambda it: (order.get(it["action"], 9), it["sent"] or dt.date.max))
    warn_fill = PatternFill("solid", start_color="FFF2CC")
    for i, it in enumerate(rows, start=2):
        mk = it["action"] in MAKES_PTT
        name = it.get("inv_name", "") if mk else ""
        vals = [it["stt"], it["code"], it["sent"], it["partner"], it["product_name"], it["policy"], it["serial"],
                it["buyer"], it["company"], it["mst"], it.get("tax_name", ""),
                it.get("tax_status", "") or it.get("tax_err", ""), it.get("name_check", ""),
                it.get("inv_type", "") if mk else "", name, it["ledger"], it["action"],
                it["a1"] if mk else None, it["v1"] if mk else None, it["a2"] if mk else None, it["v2"] if mk else None,
                f"=SUM(R{i}:U{i})" if mk else None, it["req_total"] or None,
                f"=V{i}-W{i}" if mk and it["req_total"] else None, " ".join(it["notes"])]
        for c, v in enumerate(vals, 1):
            ws.cell(i, c, v)
        ws.cell(i, 3).number_format = "dd/mm/yyyy"
        ws.cell(i, 10).number_format = "@"
        for c in range(18, 25):
            ws.cell(i, c).number_format = "#,##0;-#,##0;-"
        if it["notes"] and (it.get("name_check") == "MISMATCH" or it["action"] in (ACT_MISSING, ACT_INELIGIBLE)
                            or any("by hand" in n or "duplicate" in n for n in it["notes"])):
            ws.cell(i, 17).fill = warn_fill
    n = len(rows)
    _font_all(ws, 2, n + 1, len(heads))
    widths = [5, 13, 11, 11, 26, 26, 11, 24, 32, 14, 32, 18, 14, 10, 32, 30, 30, 11, 11, 11, 11, 12, 12, 10, 60]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[L(i)].width = w
    for r_ in range(2, n + 2):
        ws.cell(r_, 25).alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{L(len(heads))}{n + 1}"

    # sheet 2: don khong du dieu kien xuat HD cong ty (gui khach bo sung hop dong)
    w2 = wb.create_sheet("Không đủ điều kiện")
    w2.sheet_view.showGridLines = False
    h2 = ["STT", "Mã yêu cầu", "Ngày gửi", "Sản phẩm", "Số GCNBH/SĐBS", "Seri", "Người mua",
          "Công ty yêu cầu xuất HĐ", "MST", "Người liên hệ", "Email", "Số điện thoại", "Tổng tiền", "Lý do", "Đã xử lý"]
    for i, h in enumerate(h2, 1):
        w2.cell(1, i, h)
    _style_header(w2, 1, len(h2))
    bad = [it for it in rows if it["action"] == ACT_INELIGIBLE]
    for i, it in enumerate(bad, start=2):
        vals = [it["stt"], it["code"], it["sent"], it["product_name"], it["policy"], it["serial"], it["buyer"],
                it.get("tax_name") or it["company"], it["mst"], it["contact"], it["email"], it["phone"],
                it["req_total"] or None,
                "Đơn xe: người mua khác tên công ty yêu cầu xuất hóa đơn, cần bổ sung hợp đồng (thuê xe, ...)",
                "Đã lập PTT xuất hóa đơn cá nhân cho người mua"]
        for c, v in enumerate(vals, 1):
            w2.cell(i, c, v)
        w2.cell(i, 3).number_format = "dd/mm/yyyy"
        w2.cell(i, 9).number_format = "@"
        w2.cell(i, 13).number_format = "#,##0"
    _font_all(w2, 2, len(bad) + 1, len(h2))
    for i, w in enumerate([5, 13, 11, 26, 26, 11, 24, 40, 13, 22, 28, 13, 12, 60, 36], 1):
        w2.column_dimensions[L(i)].width = w
    w2.freeze_panes = "A2"
    wb.save(path)
    return len(bad)


# ============================================================================ chay tron goi

def run(paths, out_dir, cfg=None, period=None, slip_date=None, tax=True, gcn=True, log=print):
    """Chay ca 2 luong tren bo file. period = (tu, den) ngay cap don cho luong gom (None = tu goi y).
    Tra ve dict: outputs [path], requests (tom tat), bulk [tom tat nhom], warns, held, period."""
    warns = []
    inp = load_inputs(paths, warns)
    return run_loaded(inp, out_dir, cfg, period, slip_date, tax, gcn, log, warns)


def run_loaded(inp, out_dir, cfg=None, period=None, slip_date=None, tax=True, gcn=True, log=print, warns=None):
    cfg = cfg or load_config()
    warns = [] if warns is None else warns
    os.makedirs(out_dir, exist_ok=True)
    if inp["invoices"] is None:
        raise ToolError("Missing the invoice ledger (HÓA ĐƠN THU PHÍ export). Without it the tool cannot tell "
                        "which policies already have an invoice.")
    if inp["requests"] is None and inp["pias"] is None:
        raise ToolError("Upload the request list and/or the PIAS exports together with the invoice ledger.")
    slip_date = slip_date or dt.date.today()
    outputs, res = [], {"warns": warns, "held": [], "bulk": [], "requests": None}
    tag = slip_date.strftime("%Y%m%d")

    if inp["requests"] is not None:
        log("Reading the request list")
        items = plan_requests(inp, cfg, warns, tax=tax, gcn=gcn, gcn_dir=os.path.join(out_dir, "GCN"), log=log)
        mk = [it for it in items if it["action"] in MAKES_PTT]
        if mk:
            p = os.path.join(out_dir, f"PTT_Requests_{tag}.docx")
            write_ptt([request_slip(it, cfg, slip_date) for it in mk], p)
            outputs.append(p)
        p = os.path.join(out_dir, f"Request_Check_{tag}.xlsx")
        n_bad = write_request_check(items, p)
        outputs.append(p)
        gdir = os.path.join(out_dir, "GCN")
        if os.path.isdir(gdir):
            outputs += sorted(os.path.join(gdir, f) for f in os.listdir(gdir))
        cnt = {}
        for it in items:
            cnt[it["action"]] = cnt.get(it["action"], 0) + 1
        res["requests"] = {"items": items, "count": cnt, "ptt": len(mk), "ineligible": n_bad}
        for it in items:
            for n in it["notes"]:
                if it["action"] in MAKES_PTT or it["action"] == ACT_MISSING:
                    warns.append(f"{it['policy'] or it['code']}: {n}")
    else:
        warns.append("No request list uploaded: requested policies cannot be held out of the bulk lists.")

    if inp["pias"] is not None:
        log("Grouping uninvoiced policies")
        df = plan_bulk(inp, cfg, warns, held=res["held"])
        cand = bulk_candidates(df)
        if cand is None or cand.empty:
            warns.append("Bulk: every policy in the PIAS export already has an invoice or a request.")
        else:
            a, b = period or suggest_period(cand)
            res["period"] = (a, b)
            inside = cand[cand["issued"].map(lambda d: d is not None and a <= d <= b)]
            left = cand["policy"].nunique() - inside["policy"].nunique()
            if left:
                warns.append(f"Bulk: {left} uninvoiced policy(ies) issued outside {dmy(a)} - {dmy(b)} left for another batch.")
            for (prod, ch), g in inside.groupby(["product", "channel"], sort=True):
                ptag = period_tag(a, b)
                base = f"{prod}_{safe(ch)}_{ptag}"
                vals, s, tot = bulk_slip(g, prod, ch, a, b, cfg, slip_date)
                p1 = os.path.join(out_dir, f"PTT_{base}.docx")
                write_ptt([vals], p1)
                p2 = os.path.join(out_dir, f"List_{base}.xlsx")
                write_bulk_list(g, p2, cfg)
                outputs += [p1, p2]
                res["bulk"].append({"product": prod, "channel": ch, "policies": int(g["policy"].nunique()),
                                    "main": s["A1"], "main_vat": s["V1"], "extra": s["A2"], "extra_vat": s["V2"],
                                    "total": tot, "ptt": p1, "list": p2})
            if res["held"]:
                warns.append(f"Bulk: {len(res['held'])} uninvoiced policy(ies) held out because the customer asked for "
                             "their own invoice: " + ", ".join(res["held"]) + ".")
    res["outputs"] = outputs
    return res
