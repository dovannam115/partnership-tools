"""Đọc & chuẩn hoá 4 nguồn dữ liệu của quy trình bóc phí VNPAY D2C."""
import os, re, glob
from itertools import product
import pandas as pd
from openpyxl import load_workbook

VLOOKUP_LIT = re.compile(r'=(?:IFERROR\()?VLOOKUP\("([^"]*)"')
# VLOOKUP("<mã>", Config!$G$2:$H$24, 2, 0) — vùng tra nằm ngay trong công thức
VLOOKUP_FULL = re.compile(
    r'VLOOKUP\(\s*"([^"]*)"\s*,\s*\'?[Cc][Oo][Nn][Ff][Ii][Gg]\'?!'
    r'\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+)')
IF_EQ = re.compile(r'^=\s*IF\(\s*"?([^,="]*)"?\s*=\s*"?([^,)"]*)"?\s*,')

# mã trong cột "Loại thẻ/Tài khoản" của VNPAY -> biểu phí
RATE_NOI_DIA = {"ACC", "ATM", "EW"}
RATE_QUOC_TE = {"VISA", "MASTERCARD", "JCB", "AMEX", "UNIONPAY"}
PCT_NOI_DIA, FIX_NOI_DIA = 0.011, 1650
PCT_QUOC_TE, FIX_QUOC_TE = 0.024, 2200

VAT = 1.1
# mã sp: (miễn VAT, chịu VAT)
MA_SP = {
    "oto":    {"chinh": 50101, "phu": 50103, "mien_vat": set()},
    "moto":   {"chinh": 50201, "phu": 50202, "mien_vat": set()},
    "dlcn":   {"chinh": 10605, "phu": 10612, "mien_vat": {10605}},
}
KENH = {"ONLPVI01": "WEBSITE ONL", "DIGRCV10": "RABBIT CARE"}

_s = lambda x: x.astype(str).str.strip().str.replace(r'^"|"$', "", regex=True).str.replace(r"\.0$", "", regex=True)


def _formula_codes(ws, col, first_row=4):
    """Bóc literal trong =VLOOKUP("<mã>",Config!..) — VNPAY giấu dữ liệu ở đây."""
    out = {}
    for r in range(first_row, ws.max_row + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and v.startswith("="):
            m = VLOOKUP_LIT.match(v)
            if m:
                out[r] = m.group(1)
    return out


def _resolve(formula, cfg_ws, cache):
    """Tự tra bảng Config theo đúng vùng ghi trong công thức -> ra chữ như Excel hiện.

    Cột `Loại thẻ/Tài khoản` tra G2:H24: Config chứa CẢ mã số (01,02...) lẫn mã chữ
    (ACC, VISA, MASTERCARD...) trong cùng cột G, nên "ACC" -> "Tài khoản nội địa".
    """
    if not isinstance(formula, str) or not formula.startswith("="):
        return None
    hits = VLOOKUP_FULL.findall(formula)
    if not hits:
        return None
    m = IF_EQ.match(formula)
    lit, c1, r1, c2, r2 = hits[0] if (m and m.group(1).strip() == m.group(2).strip()) else hits[-1]
    key = (c1, r1, c2, r2)
    if key not in cache:
        from openpyxl.utils import column_index_from_string as CI
        cache[key] = {str(cfg_ws.cell(r, CI(c1)).value).strip(): cfg_ws.cell(r, CI(c2)).value
                      for r in range(int(r1), int(r2) + 1)
                      if cfg_ws.cell(r, CI(c1)).value is not None}
    return cache[key].get(str(lit).strip())


def _config_nguoc(cfg_ws):
    """Sheet Config -> {tên cột: {chữ hiển thị: mã}}, khối mã/chữ đặt cạnh nhau, tiêu đề ở dòng 1.

    Một chữ có thể ứng với nhiều mã (Visa = "02" và "VISA"); giữ mã chữ vì biểu phí
    RATE_QUOC_TE/RATE_NOI_DIA tra theo mã chữ.
    """
    out = {}
    if cfg_ws is None:
        return out
    for c in range(1, cfg_ws.max_column + 1):
        head = cfg_ws.cell(1, c).value
        if not head:
            continue
        m = {}
        for r in range(2, cfg_ws.max_row + 1):
            ma, chu = cfg_ws.cell(r, c).value, cfg_ws.cell(r, c + 1).value
            if ma is None or chu is None:
                continue
            k = str(chu).strip().casefold()
            if k not in m or str(m[k]).isdigit():
                m[k] = str(ma).strip()
        out[str(head).strip()] = m
    return out


def read_vnpay(path):
    """Sao kê VNPAY. Header ở dòng 3. Trả về cả cột ẩn trong công thức."""
    df = pd.read_excel(path, sheet_name="Sheet1", header=2)
    thieu = [c for c in ("Mã giao dịch", "Mã đơn hàng", "Số tiền trước KM", "Mã điểm thu")
             if c not in df.columns]
    if thieu:
        raise SystemExit(f"\n[LỖI] File VNPAY không đúng mẫu sao kê, thiếu cột: "
                         f"{', '.join(thieu)}\n    {path}\n")
    df = df[df["Mã giao dịch"].notna()].reset_index(drop=True)
    # sao kê gốc ghi giờ dạng chữ dd/mm/yyyy hh:mm:ss; bản dán giá trị ra kiểu ngày,
    # để nguyên thì ô trong file kết quả hiện số serial vì style cột là text
    for c in ("Thời gian GD", "Thời gian đối soát"):
        if c in df.columns and str(df[c].dtype).startswith("datetime"):
            df[c] = df[c].dt.strftime("%d/%m/%Y %H:%M:%S")

    wb = load_workbook(path, data_only=False, read_only=False)
    ws = wb["Sheet1"]
    cfg = next((wb[s] for s in wb.sheetnames if s.lower() == "config"), None)
    hdr = {str(ws.cell(3, c).value).strip(): c for c in range(1, ws.max_column + 1) if ws.cell(3, c).value}
    rows = [r for r in range(4, ws.max_row + 1) if ws.cell(r, 3).value is not None]

    # Bản tải thẳng từ VNPAY giấu mã trong =VLOOKUP("<mã>",Config!..). Bản lọc tay
    # (file "đơn thiếu") đã dán thành chữ, không còn công thức -> tra ngược Config
    # để lấy lại mã, nếu không thì mã trạng thái rỗng và mọi dòng bị coi là hỏng.
    cache, rev = {}, _config_nguoc(cfg)
    MA_COT = {"Loại thẻ/Tài khoản": "_loai_the", "Trạng thái": "_trang_thai",
              "Nguồn tiền": "_nguon_tien_ma", "Loại giao dịch": "_loai_gd_ma"}
    for name in ("Loại giao dịch", "Kênh thanh toán KTS", "Loại thẻ/Tài khoản",
                 "Nguồn tiền", "Dịch vụ", "Kênh thanh toán", "Trạng thái"):
        if name not in hdr:
            continue
        c, ma, chu = hdr[name], [], []
        r_map = rev.get(name, {})
        for r in rows:
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.startswith("="):
                m = VLOOKUP_LIT.match(v)
                ma.append(m.group(1) if m else None)
                chu.append(_resolve(v, cfg, cache) if cfg is not None else None)
            else:
                t = str(v).strip() if v is not None else None
                k = r_map.get(t.casefold()) if t else None
                # Config không có mọi mã (VD "AMEX" chỉ có dòng chữ "Thẻ Amex"),
                # ô dán nguyên mã thì lấy luôn nó, kẻo tính nhầm biểu phí nội địa
                if k is None and t and re.fullmatch(r"[A-Za-z0-9_]+", t):
                    k = t.upper()
                ma.append(k)
                chu.append(t)
        df["TXT " + name] = chu
        if name in MA_COT:
            df[MA_COT[name]] = ma
    wb.close()

    df["_mdh"] = _s(df["Mã đơn hàng"])
    df["_kenh"] = df["Mã điểm thu"].map(KENH).fillna(df["Mã điểm thu"])
    return df


def read_admin(paths):
    """Export Admin website, 1 file/kênh."""
    out = []
    for p in paths:
        d = pd.read_csv(p, encoding="utf-8-sig")
        d["_file"] = os.path.basename(p)
        out.append(d)
    a = pd.concat(out, ignore_index=True)
    a["_pay"] = _s(a["Payment ID"])
    a["_serial"] = _s(a["PVI Serial"])
    a["_gcn"] = _s(a["Số GCNBH"])
    return a


def read_pias(paths):
    """4 file PIAS -> 1 bảng tra (khoá, số đơn BH) + 1 bảng phí (số seri, mã sp -> phí chưa VAT).

    Khoá khác nhau theo nghiệp vụ:
      xe ô tô / xe máy : Số seri
      du lịch          : Số thẻ      (PIAS DLCN không có cột Số seri)
      tài sản          : Số đơn BH   (nối thẳng với Số GCNBH bên Admin)
    """
    tra, phi = [], []
    for p in paths:
        d = pd.read_excel(p, sheet_name=0)
        if "Số seri" in d.columns:                       # xe ô tô / xe máy
            # PIAS xuất 2 mẫu: mẫu bảng kê (Số đơn SĐBS, Số tiền phí, có khi thiếu cột
            # Số đơn SĐBS) và mẫu chung như DLCN (Mã sđbs, Phí bảo hiểm(VNĐ))
            sdbs = next((c for c in ("Số đơn SĐBS", "Mã sđbs", "Tên sđbs") if c in d.columns), None)
            if sdbs:
                d = d[d[sdbs].isna()]                    # bỏ dòng sửa đổi bổ sung
            ph = "Số tiền phí" if "Số tiền phí" in d.columns else "Phí bảo hiểm(VNĐ)"
            k, don = "Số seri", "Số đơn BH"
        elif "Số thẻ" in d.columns:                      # du lịch
            don = "Số đơn BH/SĐBS" if "Số đơn BH/SĐBS" in d.columns else "Số đơn BH"   # 10/2026 PIAS đổi tên cột
            k, ph = "Số thẻ", "Phí bảo hiểm(VNĐ)"
        else:                                            # tài sản
            don = "Số đơn BH/SĐBS" if "Số đơn BH/SĐBS" in d.columns else "Số đơn BH"
            k, ph = don, "Phí bảo hiểm(VNĐ)"
        thieu = [c for c in dict.fromkeys((k, don, ph, "Mã sp")) if c not in d.columns]
        if thieu:
            raise SystemExit(f"\n[LỖI] File PIAS không đúng mẫu, thiếu cột: "
                             f"{', '.join(thieu)}\n    {os.path.basename(p)}\n")
        if k == "Số thẻ":
            # đơn nhóm/gia đình (…/DLCN/P…) không có Số thẻ -> khoá bằng số đơn,
            # reconcile tra dự phòng theo Số GCNBH bên Admin
            d[k] = d[k].fillna(d[don])
        d = d[d[k].notna()].copy()
        d["_key"] = _s(d[k])
        tra.append(d[["_key", don]].rename(columns={don: "_don"}).drop_duplicates("_key"))
        phi.append(d[["_key", "Mã sp", ph]].rename(columns={ph: "_phi"}))
    return (pd.concat(tra, ignore_index=True).drop_duplicates("_key"),
            pd.concat(phi, ignore_index=True))


def read_hoadon(path):
    """Sổ hoá đơn toàn công ty -> map số đơn BH sang số chứng từ (= số hoá đơn)."""
    d = pd.read_excel(path, sheet_name=0)
    # 08/2026 ghi "Số đơn Bh", 09/2026 ghi "Số đơn BH" -> dò không phân biệt hoa thường
    cols = {str(c).casefold(): c for c in d.columns}
    col_don = cols.get("số đơn bh.1") or cols.get("số đơn bh")
    if col_don is None or "Số ctừ" not in d.columns:
        raise SystemExit(f"\n[LỖI] File hoá đơn không đúng mẫu, thiếu cột Số đơn BH / Số ctừ"
                         f"\n    {os.path.basename(path)}\n")
    d = d[d[col_don].notna() & d["Số ctừ"].notna()].copy()
    d["_don"] = _s(d[col_don])
    d["_hd"] = d["Số ctừ"]
    return d.sort_values("Ngày ctừ").drop_duplicates("_don", keep="last")[["_don", "_hd"]]


def phi_dich_vu(phi_bh, loai_the):
    """Biểu phí VNPAY. Suy ngược từ 539/539 dòng file thật, sai 0."""
    if loai_the in RATE_QUOC_TE:
        return round(phi_bh * PCT_QUOC_TE + FIX_QUOC_TE)
    return round(phi_bh * PCT_NOI_DIA + FIX_NOI_DIA)


def tach_sp(key, nghiep_vu, tong_phi, bang_phi):
    """Tách tổng phí (đã VAT) thành Phí BH 1 / Phí BH 2 theo dòng sản phẩm PIAS.

    PIAS ghi phí CHƯA VAT còn khách trả ĐÃ VAT, và PIAS vẫn có dòng sp phụ kể cả
    khi khách không mua -> suy ngược từ tổng phí (số tiền thật, đã khớp VNPAY).
    Phí BH 2 luôn = tổng - Phí BH 1 để cột check (SP1 + SP2 - Phí BH) ra đúng 0.
    """
    r = bang_phi[bang_phi["_key"] == key]
    if r.empty:
        return (None, None)
    tong = round(tong_phi)
    if nghiep_vu == "taisan":
        return _tach_taisan([float(x) for x in r["_phi"] if x], tong)
    cfg = MA_SP.get(nghiep_vu)
    if cfg is None:
        return (None, None)
    g = r.groupby("Mã sp")["_phi"].sum()
    c = float(g.get(cfg["chinh"], 0) or 0)
    p = float(g.drop(cfg["chinh"], errors="ignore").sum())   # mọi mã phụ (50103, 50104...)

    if nghiep_vu == "dlcn":                       # phạm vi chính miễn VAT
        if p and abs(c + p * VAT - tong) <= 1:
            return (round(c), tong - round(c))
        if abs(c - tong) <= 1:
            return (tong, None)
        return (None, None)

    # xe: cả hai mã đều chịu VAT
    if p and abs((c + p) * VAT - tong) <= 2:
        s1 = round(c * VAT)
        return (s1, tong - s1)
    if abs(c * VAT - tong) <= 1:                  # khách chỉ mua sp chính
        return (tong, None)
    return (None, None)


def _tach_taisan(lines, tong):
    """Mọi rủi ro căn hộ: các dòng cùng mã sp 20113, có dòng chịu VAT có dòng không
    (T7/2026: dòng căn hộ tỷ lệ 0,09% chịu VAT, dòng 60 triệu tỷ lệ 0,11% = 66.000
    không VAT). PIAS không ghi dòng nào chịu thuế -> thử tổ hợp cho khớp tổng phí.
    SP1 = phần không VAT, SP2 = phần có VAT; cùng một loại thuế thì SP1 = dòng đầu.
    """
    if not lines or len(lines) > 8:
        return (None, None)
    for vat in product((True, False), repeat=len(lines)):      # thử "tất cả VAT" trước
        if abs(sum(x * VAT if v else x for x, v in zip(lines, vat)) - tong) > 1:
            continue
        khong = round(sum(x for x, v in zip(lines, vat) if not v))
        if 0 < khong < tong:
            return (khong, tong - khong)
        if len(lines) == 1:
            return (tong, None)
        s1 = round(lines[0] * (VAT if vat[0] else 1))
        return (s1, tong - s1)
    return (None, None)


def nghiep_vu_tu_don(so_don):
    if not isinstance(so_don, str):
        return None
    m = re.search(r"/(AUTO|MOTO|DLCN|TSTH)/", so_don)
    return {"AUTO": "oto", "MOTO": "moto", "DLCN": "dlcn", "TSTH": "taisan"}.get(m.group(1)) if m else None
