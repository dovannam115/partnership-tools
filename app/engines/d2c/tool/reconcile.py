"""Dựng file 'D2C - VNPAY - Bóc phí MM.YYYY.xlsx' từ 4 nguồn: VNPAY, Admin, PIAS, sổ hoá đơn.

    python reconcile.py <thư_mục_input> <file_template> <file_output>

Thư mục input cần các file (khớp theo tên, không phân biệt hoa thường):
    vnpay*.xlsx | admin*.csv (1 file/kênh) | pias*.xlsx (4 file) | *hóa đơn*.xlsx
"""
import sys, os, glob, re, warnings
from copy import copy
warnings.filterwarnings("ignore", module="openpyxl")
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from engine import (read_vnpay, read_admin, read_pias, read_hoadon, phi_dich_vu,
                    tach_sp, nghiep_vu_tu_don, RATE_QUOC_TE, PCT_QUOC_TE, FIX_QUOC_TE,
                    PCT_NOI_DIA, FIX_NOI_DIA, KENH, _s)

TRANG_THAI = {"1": "Thành công", "2": "Không thành công", "3": "Nghi vấn",
              "4": "Hủy không thành công", "5": "Hủy thành công",
              "6": "Đang thực hiện", "8": "Hết hạn thanh toán"}
COL = {"stt": 1, "tg": 2, "magd": 3, "cn": 4, "madt": 5, "dt": 6, "shd": 7, "shd2": 8,
       "mdh": 9, "mtt": 10, "mkm": 11, "sdt": 12, "tenkh": 13, "sotk": 14, "loaigd": 15,
       "kttkts": 16, "phibh": 17, "truockm": 18, "phidv": 19, "conlai": 20, "sodon": 21,
       "hoadon": 22, "ttdh": 23, "nh": 24, "loaithe": 25, "loaiph": 26, "vnpaytt": 27,
       "nguontien": 28, "dichvu": 29, "soct": 30, "kenhtt": 31, "trangthai": 32,
       "tgds": 33, "ghichu": 34, "ttbs": 35, "maloi": 36, "sp1": 37, "sp2": 38,
       "phidvsp1": 39, "phidvsp2": 40, "chkphi": 41, "chkdv": 42, "kenh": 43}
HDR_ROW, FIRST_ROW = 3, 4
# các cột tiền -> mượn style (Times New Roman + #,##0 + viền) của cột "Phí Bảo hiểm"
MONEY_COLS = {COL[c] for c in ("phibh", "truockm", "phidv", "conlai", "sp1", "sp2",
                               "phidvsp1", "phidvsp2", "chkphi", "chkdv")}
# 2 cột double check, khác 0 là lệch (giống cột phòng tự thêm ở file làm tay)
CHECK_HDR = {"chkphi": "Check phí BH\n(SP1 + SP2 - Phí BH)",
             "chkdv": "Check phí DV\n(DV SP1 + DV SP2 - Phí DV)"}


def _find(folder, *pats):
    for p in pats:
        hits = [f for f in glob.glob(os.path.join(folder, "**", "*"), recursive=True)
                if re.search(p, os.path.basename(f), re.I) and not os.path.basename(f).startswith("~$")]
        if hits:
            return sorted(hits)
    return []


def _check_ghi_duoc(output):
    """Excel khoá file đang mở -> báo rõ thay vì ném PermissionError khó hiểu."""
    if not os.path.exists(output):
        return
    try:
        with open(output, "r+b"):
            pass
    except PermissionError:
        raise SystemExit(
            f"\n[LỖI] File đang mở trong Excel:\n    {output}\n"
            f"Đóng file đó rồi chạy lại.\n")


def build(folder, template, output, progress=None, log=None):
    step = progress or (lambda pct, msg: None)
    say = log or print
    _check_ghi_duoc(output)

    say(f"Input folder: {folder}")
    vn_f = _find(folder, r"^vnpay.*\.xlsx$", r"danhsachgiaodich.*\.xlsx$")
    adm_f = _find(folder, r"^admin.*\.csv$")
    pias_f = _find(folder, r"^pias.*\.xlsx$")
    hd_f = _find(folder, r"hóa đơn thu phí.*\.xlsx$", r"hoa don thu phi.*\.xlsx$")
    for name, fs, pat in [("VNPAY statement", vn_f, "VNPAY*.xlsx"),
                          ("Admin export", adm_f, "ADMIN*.csv"),
                          ("PIAS export", pias_f, "PIAS*.xlsx"),
                          ("Invoice ledger", hd_f, "*HOA DON THU PHI*.xlsx")]:
        if not fs:
            raise SystemExit(f"\n[ERROR] Missing {name} ({pat}) in:\n    {folder}\n")

    step(10, "Reading input files")
    say(f"VNPAY   {os.path.basename(vn_f[0])}")
    vn = read_vnpay(vn_f[0])
    say(f"        {len(vn)} transactions")
    for f in adm_f:
        say(f"Admin   {os.path.basename(f)}")
    adm = read_admin(adm_f)
    say(f"        {len(adm)} orders")
    for f in pias_f:
        say(f"PIAS    {os.path.basename(f)}")
    tra, bang_phi = read_pias(pias_f)
    say(f"        {len(tra)} policy keys")
    say(f"Invoice {os.path.basename(hd_f[0])}")
    hd = read_hoadon(hd_f[0])
    say(f"        {len(hd)} invoices")

    tong_gd = len(vn)
    vn["_ok"] = vn.get("_trang_thai").eq("1")
    # sao kê VNPAY xuất theo thứ tự giảm dần; file làm tay sắp tăng dần theo thời gian
    vn["_ts"] = pd.to_datetime(vn["Thời gian GD"], dayfirst=True, errors="coerce")
    vn = vn.sort_values("_ts", kind="stable").reset_index(drop=True)
    df = vn[vn["_ok"]].copy().reset_index(drop=True)
    # 0 dòng thành công = đọc sai cột Trạng thái, đừng ghi ra file rỗng rồi để người
    # dùng tự đoán
    if tong_gd and df.empty:
        raise SystemExit(
            f"\n[LỖI] {tong_gd} giao dịch nhưng không đọc được dòng nào có "
            f"Trạng thái = Thành công.\nKiểm tra cột Trạng thái của file VNPAY:\n"
            f"    {vn_f[0]}\n")

    a = adm.drop_duplicates("_pay").set_index("_pay")
    df = df.join(a[["Tổng phí bảo hiểm", "Tên sản phẩm", "_serial", "_gcn",
                    "Phí BH chính", "Phí BH bổ sung", "Phí VAT"]], on="_mdh")
    df["_khop_admin"] = df["Tổng phí bảo hiểm"].notna()

    tra_map = tra.set_index("_key")["_don"]
    key = df["_serial"].where(df["_serial"].notna() & (df["_serial"] != "nan"), df["_gcn"])
    df["_key"] = key
    df["_don"] = key.map(tra_map).fillna(df["_gcn"].map(tra_map))   # dự phòng: đơn DLCN loại P
    df["_hoadon"] = df["_don"].map(hd.set_index("_don")["_hd"])
    df["_phibh"] = df["Tổng phí bảo hiểm"].fillna(df["Số tiền trước KM"])
    df["_nv"] = df["_don"].map(nghiep_vu_tu_don)

    step(55, "Matching transactions")
    say(f"Successful {int(df['_ok'].sum() if '_ok' in df else len(df))}"
        f" | matched to Admin {int(df['_khop_admin'].sum())}"
        f" | policy no. {int(df['_don'].notna().sum())}"
        f" | invoice no. {int(df['_hoadon'].notna().sum())}")
    sp1, sp2 = [], []
    for _, r in df.iterrows():
        if pd.notna(r["Phí BH chính"]):
            # Admin có sẵn tách phí du lịch, chuẩn hơn PIAS (PIAS ghi ở cấp đơn gom);
            # dùng cả khi PIAS chưa cấp đơn (T7/2026: 32 dòng)
            a1, a2 = round(r["Phí BH chính"]), round(r["Phí BH bổ sung"] + r["Phí VAT"])
            sp1.append(a1 if a1 else None); sp2.append(a2 if a2 else None)
        elif pd.notna(r["_nv"]) and pd.notna(r["_phibh"]):
            x, y = tach_sp(r["_key"], r["_nv"], r["_phibh"], bang_phi)
            sp1.append(x); sp2.append(y)
        else:
            sp1.append(None); sp2.append(None)
    df["_sp1"], df["_sp2"] = sp1, sp2

    step(75, "Writing workbook")
    wb = load_workbook(template, data_only=False)
    ws = wb["Sheet1"]
    ncol = ws.max_column
    # chép style dòng dữ liệu của template TRƯỚC khi xoá, nếu không sẽ mất
    # font/viền/định dạng số và file ra trông khác hẳn bản làm tay
    proto = {c: copy(ws.cell(FIRST_ROW, c)._style) for c in range(1, ncol + 1)}
    proto_money = copy(ws.cell(FIRST_ROW, COL["phibh"])._style)
    for c in MONEY_COLS:
        proto[c] = copy(proto_money)
    row_h = ws.row_dimensions[FIRST_ROW].height
    if ws.max_row >= FIRST_ROW:
        ws.delete_rows(FIRST_ROW, ws.max_row - FIRST_ROW + 1)
    for lnk in list(getattr(wb, "_external_links", [])):
        wb._external_links.remove(lnk)

    g = lambda r, c: r[c] if c in r and pd.notna(r[c]) else None
    for i, (_, r) in enumerate(df.iterrows()):
        x = FIRST_ROW + i
        put = lambda c, v: ws.cell(x, COL[c], v)
        put("stt", i + 1)
        for c, src in [("tg", "Thời gian GD"), ("magd", "Mã giao dịch"), ("cn", "Chi nhánh"),
                       ("madt", "Mã điểm thu"), ("dt", "Điểm thu"), ("shd", "Số hóa đơn"),
                       ("mdh", "Mã đơn hàng"), ("mtt", "Mã trừ tiền/Mã chuẩn chi"),
                       ("mkm", "Mã khuyến mại"), ("sdt", "Số điện thoại"),
                       ("tenkh", "Tên khách hàng"), ("sotk", "Số TK/Thẻ"),
                       ("truockm", "Số tiền trước KM"), ("ttdh", "Thông tin đặt hàng"),
                       ("nh", "Ngân hàng"), ("loaiph", "Loại phát hành"),
                       ("tgds", "Thời gian đối soát"), ("ghichu", "Ghi chú"),
                       ("ttbs", "Thông tin bổ sung"), ("maloi", "Mã lỗi bank")]:
            put(c, g(r, src))
        # chữ đã giải mã qua sheet Config, hiển thị giống hệt file làm tay
        for c, src in [("loaigd", "TXT Loại giao dịch"), ("kttkts", "TXT Kênh thanh toán KTS"),
                       ("loaithe", "TXT Loại thẻ/Tài khoản"), ("nguontien", "TXT Nguồn tiền"),
                       ("dichvu", "TXT Dịch vụ"), ("kenhtt", "TXT Kênh thanh toán"),
                       ("trangthai", "TXT Trạng thái")]:
            put(c, g(r, src))
        if ws.cell(x, COL["trangthai"]).value is None:
            put("trangthai", TRANG_THAI.get(str(g(r, "_trang_thai")), g(r, "_trang_thai")))
        put("kenh", r["_kenh"])
        put("phibh", None if pd.isna(r["_phibh"]) else round(r["_phibh"]))
        pct, fix = ((PCT_QUOC_TE, FIX_QUOC_TE) if r.get("_loai_the") in RATE_QUOC_TE
                    else (PCT_NOI_DIA, FIX_NOI_DIA))
        put("phidv", f"=ROUND(Q{x}*{pct*100:.1f}%+{fix},0)")
        put("conlai", f"=Q{x}-S{x}")
        put("sodon", r["_don"] if pd.notna(r["_don"]) else None)
        put("hoadon", r["_hoadon"] if pd.notna(r["_hoadon"]) else None)
        put("sp1", r["_sp1"]); put("sp2", r["_sp2"])
        # công thức Partnership chốt 07/10: Phí DV SP2 tính theo tỷ trọng SP2/Phí BH rồi
        # làm tròn, Phí DV SP1 = Phí DV - Phí DV SP2 (tổng luôn khớp phí VNPAY thu)
        put("phidvsp2", f'=IF(AL{x}="",0,IFERROR(ROUND(($Q{x}*{pct*100:.1f}%)*(AL{x}/$Q{x})'
                        f'+({fix}*(AL{x}/$Q{x})),0),0))')
        # SP1 trống thì Phí DV SP1 = 0, để dòng chưa tách phí lộ ra ở cột check
        put("phidvsp1", f'=IF(AK{x}="",0,$S{x}-AN{x})')
        put("chkphi", f"=AK{x}+AL{x}-Q{x}")
        put("chkdv", f"=AM{x}+AN{x}-S{x}")
        for c in range(1, ncol + 1):
            ws.cell(x, c)._style = copy(proto[c])
        if row_h:
            ws.row_dimensions[x].height = row_h

    n = len(df)
    last = FIRST_ROW + n - 1
    ws.auto_filter.ref = f"A{HDR_ROW}:{ws.cell(HDR_ROW, ncol).column_letter}{last}"
    ws.freeze_panes = f"A{FIRST_ROW}"
    ws.cell(2, 5, n); ws.cell(2, 7, f"=SUM(R{FIRST_ROW}:R{last})")
    # tiêu đề + số dòng lệch của 2 cột check
    for k, txt in CHECK_HDR.items():
        c = COL[k]
        ws.cell(HDR_ROW, c, txt)._style = copy(ws.cell(HDR_ROW, COL["phidvsp2"])._style)
        L = ws.cell(HDR_ROW, c).column_letter
        ws.cell(2, c, f'=COUNTIF({L}{FIRST_ROW}:{L}{last},"<>0")')._style = copy(proto_money)
        ws.column_dimensions[L].width = 18

    _sheet_canhbao(wb, vn, df, tong_gd)
    _sheet_treo(wb, df)
    _sheet_tonghop(wb, df, vn, tong_gd, n)
    step(88, "Saving")
    say(f"Sheets: Sheet1 ({n} rows), Cảnh báo, Treo, Tổng hợp")
    wb.save(output)
    say(f"Saved {os.path.basename(output)}")
    return df, vn, tong_gd, n


def _style(ws, ncol):
    """Header vàng + Times New Roman, khớp quy ước của template."""
    hdr = Font(bold=True, color="FF110102", name="Times New Roman", size=12)
    fill = PatternFill("solid", start_color="FFFFFF00")
    body = Font(name="Times New Roman", size=12, color="FF110102")
    for c in range(1, ncol + 1):
        cell = ws.cell(1, c)
        cell.font = hdr; cell.fill = fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ncol):
        for cell in row:
            cell.font = body
            if isinstance(cell.value, (int, float)):
                cell.number_format = "#,##0"
    ws.freeze_panes = "A2"


def _sheet_canhbao(wb, vn, df, tong_gd):
    ws = wb.create_sheet("Cảnh báo")
    ws.append(["Nhóm", "Mã giao dịch", "Thời gian GD", "Kênh", "Số tiền", "Ngân hàng",
               "Số đơn BH", "Ghi chú xử lý"])
    for _, r in vn[~vn["_ok"]].iterrows():
        ws.append(["GD không thành công", r["Mã giao dịch"], r["Thời gian GD"], r["_kenh"],
                   r["Số tiền trước KM"], r["Ngân hàng"], None,
                   f"Trạng thái = {TRANG_THAI.get(str(r.get('_trang_thai')), '?')}, loại khỏi đối soát"])
    for _, r in df[~df["_khop_admin"]].iterrows():
        ws.append(["GD thành công, không có trong Admin", r["Mã giao dịch"], r["Thời gian GD"],
                   r["_kenh"], r["Số tiền trước KM"], r["Ngân hàng"], None,
                   "Tra portal thủ công: sản phẩm/kênh nào?"])
    for _, r in df[df["_khop_admin"] & df["_don"].isna()].iterrows():
        ws.append(["Có Admin, chưa có số đơn BH", r["Mã giao dịch"], r["Thời gian GD"],
                   r["_kenh"], r["Số tiền trước KM"], r["Ngân hàng"], None,
                   "PIAS chưa cấp đơn, xem sheet Treo"])
    for _, r in df[df["_don"].notna() & df["_hoadon"].isna()].iterrows():
        ws.append(["Có số đơn, chưa có hoá đơn", r["Mã giao dịch"], r["Thời gian GD"],
                   r["_kenh"], r["Số tiền trước KM"], r["Ngân hàng"], r["_don"],
                   "Chưa xuất hoá đơn hoặc chưa có trong sổ"])
    for _, r in df[df["_khop_admin"] & df["_sp1"].isna() & df["_sp2"].isna()].iterrows():
        ws.append(["Không tách được SP1/SP2", r["Mã giao dịch"], r["Thời gian GD"],
                   r["_kenh"], r["Số tiền trước KM"], r["Ngân hàng"], r["_don"],
                   "Tổng phí không khớp tổ hợp mã sp trong PIAS"])
    _style(ws, 8)
    for c, w in zip("ABCDEFGH", (34, 14, 20, 14, 14, 14, 26, 46)):
        ws.column_dimensions[c].width = w


def _sheet_treo(wb, df):
    ws = wb.create_sheet("Treo")
    ws.append(["Mã giao dịch", "Thời gian GD", "Kênh", "Sản phẩm", "PVI Serial", "Số GCNBH",
               "Số tiền", "Lý do", "Đã xử lý (điền tay)"])
    for _, r in df[df["_khop_admin"] & df["_don"].isna()].iterrows():
        ws.append([r["Mã giao dịch"], r["Thời gian GD"], r["_kenh"], r.get("Tên sản phẩm"),
                   r.get("_serial"), r.get("_gcn"), r["Số tiền trước KM"],
                   "Chưa có trong PIAS kỳ này, nạp lại sheet Treo ở kỳ sau", None])
    _style(ws, 9)
    for c, w in zip("ABCDEFGHI", (14, 20, 14, 30, 14, 24, 14, 46, 20)):
        ws.column_dimensions[c].width = w


def _sheet_tonghop(wb, df, vn, tong_gd, n):
    ws = wb.create_sheet("Tổng hợp")
    ws.append(["Chỉ tiêu", "Số GD", "Số tiền", "Ghi chú"])
    add = lambda *v: ws.append(list(v))
    add("VNPAY tải về", tong_gd, float(vn["Số tiền trước KM"].sum()), "toàn bộ dòng trong sao kê")
    bad = vn[~vn["_ok"]]
    add("GD không thành công", len(bad), float(bad["Số tiền trước KM"].sum()), "Trạng thái ≠ 1, loại bỏ")
    add("GD thành công", n, float(df["Số tiền trước KM"].sum()), "Trạng thái = 1")
    k = df[df["_khop_admin"]]
    add("khớp Admin", len(k), float(k["Số tiền trước KM"].sum()), "nối Mã đơn hàng = Payment ID")
    nk = df[~df["_khop_admin"]]
    add("KHÔNG khớp Admin", len(nk), float(nk["Số tiền trước KM"].sum()), "cần tra tay")
    add("Có số đơn BH", int(df["_don"].notna().sum()), float(df.loc[df["_don"].notna(), "Số tiền trước KM"].sum()), "")
    add("Có số hoá đơn", int(df["_hoadon"].notna().sum()), float(df.loc[df["_hoadon"].notna(), "Số tiền trước KM"].sum()), "")
    add("Tách được SP1/SP2", int((df["_sp1"].notna() | df["_sp2"].notna()).sum()), None, "")
    ws.append([])
    ws.append(["Theo kênh", "Số GD", "Số tiền", ""])
    for kenh, sub in df.groupby("_kenh"):
        add(kenh, len(sub), float(sub["Số tiền trước KM"].sum()), "")
    ws.append([])
    ws.append(["Chênh lệch tiền VNPAY vs Admin", None, None, ""])
    d = (k["Số tiền trước KM"] - k["Tổng phí bảo hiểm"]).abs()
    add("Dòng khớp tuyệt đối", int((d < 1).sum()), None, f"trên tổng {len(k)} dòng đối chiếu được")
    add("Dòng lệch tiền", int((d >= 1).sum()), float(d[d >= 1].sum()), "")
    _style(ws, 4)
    for c, w in zip("ABCD", (40, 12, 18, 44)):
        ws.column_dimensions[c].width = w


def summary(df, tong, n):
    return (f"{tong} transactions | {n} successful | {int(df['_khop_admin'].sum())} matched\n"
            f"Policy no. {int(df['_don'].notna().sum())} | "
            f"Invoice no. {int(df['_hoadon'].notna().sum())} | "
            f"Premium split {int((df['_sp1'].notna() | df['_sp2'].notna()).sum())}")


if __name__ == "__main__":
    try:                                  # console Windows mặc định cp1252/cp437,
        sys.stdout.reconfigure(encoding="utf-8")   # tên file có dấu sẽ nổ UnicodeEncodeError
    except Exception:
        pass
    if len(sys.argv) < 4:
        raise SystemExit(__doc__)
    df, vn, tong, n = build(*sys.argv[1:4])
    print(summary(df, tong, n))
    print("Saved:", sys.argv[3])
