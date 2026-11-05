# -*- coding: utf-8 -*-
"""Bien ban doi soat Rabbit Care (Chau, phong Phat trien doi tac).

Dau vao: file raw PIAS cua thang (don cap moi) + file raw SDBS, xuat thang tu PIAS.
Dau ra: template.xlsx dien du 2 sheet "BBĐS - Reconcile" va "Danh sách - List".

Luat:
  - PIAS lap 1 dong cho moi nguoi duoc BH (cung so don, cung ma sp, cung phi) -> giu 1 dong
    theo (so don, ma sp), so the lay o dau tien co gia tri.
  - Dong la SDBS khi so don co duoi /E01.. hoac co Ma sdbs. Chi lay SDBS co phi khac 0
    (Tang phi / Giam phi); SDBS "Thay doi khac" phi 0 bo qua.
  - Phi PIAS chua gom VAT; VAT theo ma sp trong config.json (010605 0%, 010612 10%).
  - Phi dich vu = ty le (config fee_rate) x tong phi khong gom VAT.

    py rabbit_recon.py <file raw> [<file SDBS>] [--tu dd/mm/yyyy --den dd/mm/yyyy] [--out file.xlsx]
"""
import calendar
import copy
import datetime as dt
import io
import json
import os
import re
import sys

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "template.xlsx")
CONFIG = os.path.join(HERE, "config.json")

SH_BB, SH_LIST = "BBĐS - Reconcile", "Danh sách - List"
COLS = {"so": "Số đơn BH/SĐBS", "ten": "Tên đơn vị TGBH", "the": "Số thẻ", "cap": "Ngày cấp đơn",
        "dau": "Ngày đầu", "cuoi": "Ngày cuối", "sp": "Mã sp", "stbh": "Số tiền BH",
        "phi": "Phí bảo hiểm(VNĐ)", "ma_sdbs": "Mã sđbs", "ten_sdbs": "Tên sđbs"}
NEED = ("so", "sp", "phi")
END_RE = re.compile(r"/E\d+$", re.I)


class RabbitError(Exception):
    """Loi dau vao nguoi dung sua duoc (sai file, thieu cot, ma sp la)."""


def config():
    return json.load(open(CONFIG, encoding="utf-8"))


def _txt(v):
    if v is None: return ""
    if isinstance(v, float) and v.is_integer(): v = int(v)
    return str(v).strip()


def _num(v):
    if v in (None, ""): return 0.0
    if isinstance(v, (int, float)): return float(v)
    try: return float(str(v).replace(",", "").replace(" ", ""))
    except ValueError: return 0.0


def _date(v):
    if isinstance(v, dt.datetime): return v.date()
    if isinstance(v, dt.date): return v
    s = _txt(v)
    for f in ("%d/%m/%Y", "%Y-%m-%d", "%d/%m/%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S"):
        try: return dt.datetime.strptime(s, f).date()
        except ValueError: pass
    return None


def read_rows(name, data):
    """Mot file PIAS -> list dict theo COLS. Dong tieu de la dong dau co cot 'Số đơn BH/SĐBS'."""
    try:
        wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    except Exception as e:
        raise RabbitError("%s: cannot open as Excel (%s)." % (name, e))
    rows = []
    for ws in wb.worksheets:
        it = ws.iter_rows(values_only=True)
        head = None
        for _ in range(15):
            r = next(it, None)
            if r is None: break
            names = [_txt(c) for c in r]
            if COLS["so"] in names:
                head = {k: names.index(v) for k, v in COLS.items() if v in names}
                break
        if head is None: continue
        miss = [COLS[k] for k in NEED if k not in head]
        if miss: raise RabbitError("%s: missing column %s." % (name, ", ".join(miss)))
        for r in it:
            get = lambda k: r[head[k]] if k in head and head[k] < len(r) else None
            so = _txt(get("so"))
            if not so: continue                       # dong SUBTOTAL cuoi file
            rows.append({"file": name, "so": so, "ten": _txt(get("ten")), "the": _txt(get("the")),
                         "cap": _date(get("cap")), "dau": _date(get("dau")), "cuoi": _date(get("cuoi")),
                         "sp": _txt(get("sp")).zfill(6), "stbh": _num(get("stbh")), "phi": _num(get("phi")),
                         "ma_sdbs": _txt(get("ma_sdbs")), "ten_sdbs": _txt(get("ten_sdbs"))})
    wb.close()
    if not rows: raise RabbitError("%s: no PIAS rows found (column '%s' missing)." % (name, COLS["so"]))
    return rows


def is_endorsement(r):
    return bool(END_RE.search(r["so"]) or r["ma_sdbs"])


def dedup(rows, warns):
    """Gop dong lap theo (so don, ma sp), giu thu tu xuat hien."""
    out = {}
    for r in rows:
        k = (r["so"], r["sp"])
        if k not in out:
            out[k] = dict(r, n=1)
            continue
        o = out[k]; o["n"] += 1
        if not o["the"] and r["the"]: o["the"] = r["the"]
        if abs(o["phi"] - r["phi"]) > 0.5:
            warns.append("%s %s: duplicate rows have different premiums (%s, %s), the first is used."
                         % (k[0], k[1], format(int(o["phi"]), ","), format(int(r["phi"]), ",")))
    return list(out.values())


def month_of(d):
    return dt.date(d.year, d.month, 1), dt.date(d.year, d.month, calendar.monthrange(d.year, d.month)[1])


def preview(files):
    """files = [(ten, bytes)] -> dict: policies, endorsements, skipped, auto_tu, auto_den, files, warns."""
    cfg, warns, kinds = config(), [], []
    pol, end = [], []
    for name, data in files:
        rows = read_rows(name, data)
        e = [r for r in rows if is_endorsement(r)]
        p = [r for r in rows if not is_endorsement(r)]
        pol += p; end += e
        kinds.append((name, "Endorsements (SĐBS)" if e and not p else "Policies" if p and not e else "Policies + endorsements",
                      len({r["so"] for r in rows})))
    pol, end = dedup(pol, warns), dedup(end, warns)
    by_end = {}
    for r in end: by_end.setdefault(r["so"], []).append(r)
    keep = [r for r in end if any(abs(x["phi"]) > 0.5 for x in by_end[r["so"]])]
    skipped = sorted(s for s, rs in by_end.items() if not any(abs(x["phi"]) > 0.5 for x in rs))
    unknown = sorted({r["sp"] for r in pol + keep if r["sp"] not in cfg["vat"]})
    if unknown:
        raise RabbitError("Product code %s has no VAT rate in the tool. Ask Nam to add it." % ", ".join(unknown))
    days = sorted(r["cap"] for r in pol if r["cap"]) or sorted(r["cap"] for r in keep if r["cap"])
    tu, den = month_of(days[0]) if days else (None, None)
    if days and (days[0].year, days[0].month) != (days[-1].year, days[-1].month):
        warns.append("Ngày cấp đơn runs from %s to %s, more than one month. Check the period."
                     % (days[0].strftime("%d/%m/%Y"), days[-1].strftime("%d/%m/%Y")))
    if not pol and not keep:
        raise RabbitError("No policy and no endorsement with a premium change in these files.")
    return {"policies": pol, "endorsements": keep, "skipped": skipped, "auto_tu": tu, "auto_den": den,
            "files": kinds, "warns": warns}


def _styles(ws, row, last_col=13):
    return [copy.copy(ws.cell(row, c)._style) for c in range(1, last_col + 1)]


def _put(ws, row, styles, values):
    for c, st in enumerate(styles, 1):
        cell = ws.cell(row, c)
        cell._style = copy.copy(st)
        if c in values: cell.value = values[c]


def build(prev, tu, den):
    """-> (bytes cua file xlsx, tong hop dict)."""
    cfg = config()
    wb = openpyxl.load_workbook(TEMPLATE)
    wb._external_links = []                          # mau goc tro sang file 'Chi tiết' ben ngoai
    bb, ls = wb[SH_BB], wb[SH_LIST]

    # ---- Danh sach: lay kieu dong mau roi xoa het dong cu
    s_data, s_tot, s_gap = _styles(ls, 3), _styles(ls, 18), _styles(ls, 19)
    s_blk = [_styles(ls, r) for r in (20, 21, 22, 23)]
    h_blk = [ls.row_dimensions[r].height for r in (20, 21, 22, 23)]
    t_blk = [ls.cell(r, 2).value for r in (20, 21, 22, 23)]
    t_hdr = (ls["J20"].value, ls["L20"].value)
    ls.delete_rows(2, ls.max_row)
    for r in range(2, 40): ls.row_dimensions[r].height = None

    # 2 phan: Cap moi roi Dieu chinh (SDBS), moi phan 1 dong tieu de, STT rieng, 1 dong cong
    TITLE = {"pol": "Cấp mới / Issued", "end": "Điều chỉnh (SĐBS) / Adjustment"}
    SUBT = {"pol": "Cộng cấp mới", "end": "Cộng điều chỉnh"}
    r, subs = 2, {}
    for part, items in (("pol", prev["policies"]), ("end", prev["endorsements"])):
        if not items: continue
        _put(ls, r, s_tot, {2: TITLE[part]}); r += 1
        first, stt, seen = r, 0, {}
        for it in items:
            if it["so"] not in seen:
                stt += 1; seen[it["so"]] = stt
            _put(ls, r, s_data, {1: seen[it["so"]], 2: it["so"], 3: it["the"] or None, 4: it["ten"],
                                 5: it["cap"], 6: it["dau"], 7: it["cuoi"], 8: it["sp"],
                                 9: int(round(it["stbh"])), 10: int(round(it["phi"])), 11: cfg["vat"][it["sp"]],
                                 12: "=ROUND(J%d*K%d,0)" % (r, r), 13: "=J%d+L%d" % (r, r)})
            r += 1
        f = lambda col: "%s%d:%s%d" % (col, first, col, r - 1)
        _put(ls, r, s_tot, {2: SUBT[part], 3: "=SUMPRODUCT(1/COUNTIF(%s,%s))" % (f("B"), f("B")),
                            10: "=SUM(%s)" % f("J"), 12: "=SUM(%s)" % f("L"), 13: "=SUM(%s)" % f("M")})
        subs[part] = r
        r += 1
        _put(ls, r, s_gap, {}); r += 1
    last = r - 1
    tot = r
    plus = lambda col: "=" + "+".join("%s%d" % (col, x) for x in subs.values())
    _put(ls, tot, s_tot, {2: "Tổng cộng", 3: plus("C"), 10: plus("J"), 12: plus("L"), 13: plus("M")})
    _put(ls, tot + 1, s_gap, {})
    b0 = tot + 2
    rng = lambda col: "%s2:%s%d" % (col, col, last)

    def part_sum(codes, col):
        terms = "+".join('(%s="%s")' % (rng("H"), c) for c in codes)
        return "=SUMPRODUCT((%s)*%s)" % (terms, rng(col))

    vals = [{2: t_blk[0], 10: t_hdr[0], 12: t_hdr[1]},
            {2: t_blk[1], 10: part_sum(cfg["main_codes"], "J"), 12: part_sum(cfg["main_codes"], "L")},
            {2: t_blk[2], 10: part_sum(cfg["extra_codes"], "J"), 12: part_sum(cfg["extra_codes"], "L")},
            {2: t_blk[3], 13: "=J%d+L%d+J%d+L%d" % (b0 + 1, b0 + 1, b0 + 2, b0 + 2)}]
    for i in range(4):
        _put(ls, b0 + i, s_blk[i], vals[i])
        ls.row_dimensions[b0 + i].height = h_blk[i]
    ls.auto_filter.ref = "A1:M%d" % last

    # ---- BBDS
    L = "'%s'!" % SH_LIST
    ky = "%02d/%d" % (tu.month, tu.year)
    bb["C5"] = ("BIÊN BẢN ĐỐI SOÁT CHƯƠNG TRÌNH BẢO HIỂM THÁNG %s\nMinutes of Acceptance of insurance program %s" % (ky, ky))
    f1, f2 = tu.strftime("%d/%m/%Y"), den.strftime("%d/%m/%Y")
    bb["C6"] = "Kỳ thanh toán: Từ ngày %s đến ngày %s\nPayment period: From %s to %s" % (f1, f2, f1, f2)
    for row, part in ((9, "pol"), (10, "end")):
        sub = subs.get(part)
        bb.cell(row, 6)._style = copy.copy(bb.cell(row, 7)._style)
        if sub:
            bb.cell(row, 5).value = "=%sC%d" % (L, sub)
            bb.cell(row, 6).value = "=%sM%d" % (L, sub)
            bb.cell(row, 7).value = "=%sJ%d" % (L, sub)
        else:
            bb.cell(row, 5).value, bb.cell(row, 6).value, bb.cell(row, 7).value = 0, 0, 0
        bb.cell(row, 5).number_format = "00"
    bb["F11"]._style = copy.copy(bb["G11"]._style)
    bb["G12"] = cfg["fee_rate"]

    buf = io.BytesIO(); wb.save(buf)

    def total(items, vat):
        return sum(int(round(it["phi"])) if not vat else int(round(round(it["phi"]) * cfg["vat"][it["sp"]]))
                   for it in items)
    pol, end = prev["policies"], prev["endorsements"]
    fee = total(pol, False) + total(end, False)
    vat = total(pol, True) + total(end, True)
    summary = {"n_pol": len({x["so"] for x in pol}), "n_end": len({x["so"] for x in end}),
               "premium": fee, "vat": vat, "total": fee + vat, "service_fee": fee * cfg["fee_rate"],
               "period": ky,
               "endorsements": [[x["so"], x["sp"], x["ten_sdbs"], int(round(x["phi"]))] for x in end]}
    return buf.getvalue(), summary


def out_name(tu):
    return "Rabbit_Care_Recon_%d%02d.xlsx" % (tu.year, tu.month)


def main(argv):
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("files", nargs="+")
    ap.add_argument("--tu"); ap.add_argument("--den"); ap.add_argument("--out")
    a = ap.parse_args(argv)
    prev = preview([(os.path.basename(p), open(p, "rb").read()) for p in a.files])
    tu = _date(a.tu) if a.tu else prev["auto_tu"]
    den = _date(a.den) if a.den else prev["auto_den"]
    data, s = build(prev, tu, den)
    out = a.out or os.path.join(HERE, "output", out_name(tu))
    os.makedirs(os.path.dirname(out), exist_ok=True)
    open(out, "wb").write(data)
    for w in prev["warns"]: print("[WARN]", w)
    print("Policies %d, endorsements %d (skipped no-fee: %d)" % (s["n_pol"], s["n_end"], len(prev["skipped"])))
    print("Premium %s + VAT %s = %s, service fee %s" % tuple(format(int(v), ",") for v in
          (s["premium"], s["vat"], s["total"], s["service_fee"])))
    print(out)


if __name__ == "__main__":
    try: sys.stdout.reconfigure(encoding="utf-8")
    except Exception: pass
    main(sys.argv[1:])
