"""
Doc masterfile KBank (file Excel co mat khau) -> xuat giay chung nhan bao hiem
dang .docx, moi dong mot file.

Ghi chu ve file dau vao:
  - File .xlsx nhung thuc chat la OLE2 da ma hoa (magic D0CF11E0), phai giai ma
    bang mat khau truoc khi openpyxl doc duoc.
  - Header CO HAI TANG: dong tren (STT, So GCNBH, cac nhom cot) va dong duoi
    (ten cot chi tiet). Chu "So GCNBH" xuat hien o CA HAI dong, nhung du lieu
    that chi nam o cot cua dong TREN - dong duoi la nhan thua, cot rong.
    Vi vay phai do cot theo ten tren CA HAI dong.
  - Cot so tien la cong thuc (=Building*1.1) nen doc bang data_only va lam tron.
"""

from __future__ import annotations

import copy
import datetime as _dt
import io
import os
import re
import zipfile
from dataclasses import dataclass, field

import msoffcrypto
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from openpyxl import load_workbook

import docx_fill

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "template.docx")


class KbankError(Exception):
    pass


@dataclass
class Row:
    stt: int
    so_gcn: str
    ten: str
    cccd: str
    dia_chi: str
    dien_thoai: str
    dia_diem: str
    stbh: float
    phi: float
    tu: object
    den: object
    excel_row: int


@dataclass
class Result:
    rows: list = field(default_factory=list)
    files: list = field(default_factory=list)     # (ten file, bytes)
    warnings: list = field(default_factory=list)
    header_row: int = 0
    merged: object = None                         # (ten file, bytes) khi gop 1 file


def decrypt(data: bytes, password: str) -> bytes:
    """Tra ve noi dung .xlsx thuong. File khong ma hoa thi tra ve nguyen ban."""
    if data[:4] != b"\xd0\xcf\x11\xe0":
        return data
    try:
        of = msoffcrypto.OfficeFile(io.BytesIO(data))
        of.load_key(password=password)
        out = io.BytesIO()
        of.decrypt(out)
        return out.getvalue()
    except Exception as exc:
        raise KbankError(
            "Khong mo duoc file: sai mat khau hoac file hong (%s)." % exc
        ) from exc


def _txt(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def _money(v) -> str:
    """1043900000.0000001 -> '1,043,900,000'"""
    try:
        return f"{round(float(v)):,}"
    except (TypeError, ValueError):
        return _txt(v)


def _date(v) -> str:
    if isinstance(v, (_dt.datetime, _dt.date)):
        return v.strftime("%d/%m/%Y")
    s = _txt(v)
    m = re.match(r"(\d{1,2})[/-](\d{1,2})[/-](\d{4})", s)
    if m:
        return "%02d/%02d/%s" % (int(m.group(1)), int(m.group(2)), m.group(3))
    return s


def read_rows(data: bytes, password: str) -> tuple[list, int, list]:
    plain = decrypt(data, password)
    wb = load_workbook(io.BytesIO(plain), data_only=True)
    ws = wb[wb.sheetnames[0]]
    warnings = []

    # dong header = dong chua "Ho va ten"
    hdr = 0
    for r in range(1, 30):
        for c in range(1, ws.max_column + 1):
            if "họ và tên" in str(ws.cell(r, c).value or "").lower():
                hdr = r
                break
        if hdr:
            break
    if not hdr:
        raise KbankError("Khong tim thay dong tieu de (khong co cot 'Ho va ten').")

    def col(*keys):
        """Do cot theo ten tren dong header VA dong ngay tren no.

        Mot nhan co the xuat hien o CA HAI dong header: vd "So GCNBH" nam o
        dong tren (cot co du lieu that) lan dong duoi (nhan thua, cot rong).
        Nen gom het ung vien roi chon cot NAO CO DU LIEU ben duoi.
        """
        cands = []
        for rr in (hdr - 1, hdr):
            if rr < 1:
                continue
            for c in range(1, ws.max_column + 1):
                h = str(ws.cell(rr, c).value or "").lower()
                if h and all(k.lower() in h for k in keys) and c not in cands:
                    cands.append(c)
        if not cands:
            return 0
        for c in cands:
            for r in range(hdr + 1, min(hdr + 30, ws.max_row) + 1):
                if ws.cell(r, c).value not in (None, ""):
                    return c
        return cands[0]

    C = {
        "stt": col("stt"),
        "gcn": col("số gcnbh"),
        "ten": col("họ và tên"),
        "cccd": col("cmnd"),
        "dia_chi": col("địa chỉ"),
        "dien_thoai": col("điện thoại"),
        "dia_diem": col("địa điểm được bảo hiểm"),
        "stbh": col("số tiền bảo hiểm"),
        "phi": col("phí bảo hiểm (bao gồm"),
        "tu": col("ngày bắt đầu"),
        "den": col("ngày kết thúc"),
    }
    missing = [k for k, v in C.items() if not v and k != "stt"]
    if missing:
        raise KbankError("Khong tim thay cot: " + ", ".join(missing))

    rows, n = [], 0
    for r in range(hdr + 1, ws.max_row + 1):
        ten = _txt(ws.cell(r, C["ten"]).value)
        if not ten:
            continue
        n += 1
        stt_raw = ws.cell(r, C["stt"]).value if C["stt"] else None
        try:
            stt = int(stt_raw)
        except (TypeError, ValueError):
            stt = n
        gcn_base = _txt(ws.cell(r, C["gcn"]).value)
        if not gcn_base:
            warnings.append(f"Dong {r}: thieu So GCNBH.")
        rows.append(
            Row(
                stt=stt,
                so_gcn=f"{gcn_base}/{stt:02d}" if gcn_base else "",
                ten=ten,
                cccd=_txt(ws.cell(r, C["cccd"]).value),
                dia_chi=_txt(ws.cell(r, C["dia_chi"]).value),
                dien_thoai=_txt(ws.cell(r, C["dien_thoai"]).value),
                dia_diem=_txt(ws.cell(r, C["dia_diem"]).value),
                stbh=ws.cell(r, C["stbh"]).value,
                phi=ws.cell(r, C["phi"]).value,
                tu=ws.cell(r, C["tu"]).value,
                den=ws.cell(r, C["den"]).value,
                excel_row=r,
            )
        )
        for label, val in (("dia chi", rows[-1].dia_chi),
                           ("dien thoai", rows[-1].dien_thoai),
                           ("dia diem", rows[-1].dia_diem)):
            if not val:
                warnings.append(f"Dong {r} ({ten[:24]}): thieu {label}.")
    wb.close()
    if not rows:
        raise KbankError("Khong co dong du lieu nao duoi dong tieu de.")
    return rows, hdr, warnings


def _fill(row: Row, ngay_cap: _dt.date):
    """Mo template, dien du lieu mot dong, tra ve doi tuong Document."""
    doc = docx_fill.load(TEMPLATE)
    mapping = {
        "{{SO_GCN}}": row.so_gcn,
        "{{TEN}}": row.ten,
        "{{CCCD}}": row.cccd,
        "{{DIA_CHI}}": row.dia_chi,
        "{{DIEN_THOAI}}": row.dien_thoai,
        "{{DIA_DIEM}}": row.dia_diem,
        "{{STBH}}": _money(row.stbh),
        "{{PHI}}": _money(row.phi),
        "{{TU}}": _date(row.tu),
        "{{DEN}}": _date(row.den),
        "{{NGAY}}": "%02d" % ngay_cap.day,
        "{{THANG}}": "%02d" % ngay_cap.month,
        "{{NAM}}": str(ngay_cap.year),
    }
    hits = docx_fill.replace_all(doc, mapping)
    missed = [k for k, v in hits.items() if v == 0]
    if missed:
        raise KbankError(
            "Template thieu cho danh dau: %s. Chay lai _make_template.py."
            % ", ".join(missed)
        )
    return doc


def build(row: Row, ngay_cap: _dt.date) -> bytes:
    buf = io.BytesIO()
    _fill(row, ngay_cap).save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------- gop 1 file
#
# Template khong co anh, khong dung numbering trong than bai, chi mot section
# (header/footer rong) - nen gop chi la chep cac phan tu cua <w:body> sang,
# khong phai anh xa lai quan he (r:id) nhu cac thu vien gop docx tong quat.
# Chi hai thu phai xu ly: sectPr luon nam cuoi body (chen truoc no) va bookmark
# OLE_LINK cua template se trung ten khi lap lai (bo di cho sach).


def _strip_bookmarks(el) -> None:
    for tag in ("w:bookmarkStart", "w:bookmarkEnd"):
        for b in el.findall(".//" + qn(tag)):
            b.getparent().remove(b)


def _page_break(body) -> None:
    """Chen ngat trang o CUOI doan cuoi cung - khong sinh doan rong dau trang moi."""
    sect = body.find(qn("w:sectPr"))
    last = None
    for el in body.iterchildren():
        if el is sect:
            break
        last = el
    if last is None or last.tag != qn("w:p"):
        last = OxmlElement("w:p")
        if sect is None:
            body.append(last)
        else:
            sect.addprevious(last)
    run = OxmlElement("w:r")
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    run.append(br)
    last.append(run)


def build_merged(rows: list, ngay_cap: _dt.date) -> bytes:
    """Moi giay chung nhan mot trang, noi tiep nhau trong MOT file .docx."""
    if not rows:
        raise KbankError("Khong co dong nao de gop.")
    merged = _fill(rows[0], ngay_cap)
    body = merged.element.body
    for row in rows[1:]:
        doc = _fill(row, ngay_cap)
        _page_break(body)
        sect = body.find(qn("w:sectPr"))
        src_sect = doc.element.body.find(qn("w:sectPr"))
        for el in doc.element.body.iterchildren():
            if el is src_sect:
                continue
            new = copy.deepcopy(el)
            _strip_bookmarks(new)
            if sect is None:
                body.append(new)
            else:
                sect.addprevious(new)
    buf = io.BytesIO()
    merged.save(buf)
    return buf.getvalue()


def run(data: bytes, password: str, ngay_cap: _dt.date, merge: bool = False) -> Result:
    rows, hdr, warnings = read_rows(data, password)
    res = Result(rows=rows, header_row=hdr, warnings=warnings)
    if merge:
        res.merged = (
            f"GCNBH KBank {ngay_cap:%Y-%m-%d} - merged {len(rows)} certificates.docx",
            build_merged(rows, ngay_cap),
        )
    else:
        for row in rows:
            res.files.append((f"GCNBH KBank {row.stt:02d}.docx", build(row, ngay_cap)))
    return res


def zip_bytes(files: list) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for name, blob in files:
            z.writestr(name, blob)
    return buf.getvalue()
