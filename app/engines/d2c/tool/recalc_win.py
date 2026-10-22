"""Tính lại toàn bộ công thức bằng Excel COM và quét lỗi (#REF!, #DIV/0!, ...).

    python recalc_win.py <file.xlsx>
"""
import sys, os, json

ERRS = ("#REF!", "#DIV/0!", "#VALUE!", "#N/A", "#NAME?", "#NULL!", "#NUM!")


def _ghi_duoc(path):
    try:
        with open(path, "r+b"):
            return True
    except OSError:
        return False


def recalc(path):
    import win32com.client as win32
    path = os.path.abspath(path)
    ghi_duoc = _ghi_duoc(path)  # phải hỏi trước, Excel mở lên là khoá ghi ngay
    app = win32.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.AskToUpdateLinks = False
    try:                        # chặn mọi hộp thoại; thiếu dòng này Excel ẩn có thể
        app.Interactive = False  # bật "Save As" và treo vô hạn, phải kill tay
    except Exception:
        pass
    try:
        wb = app.Workbooks.Open(path, UpdateLinks=0, ReadOnly=False)
        if wb.ReadOnly and not ghi_duoc:          # Excel khác đang giữ file
            wb.Close(SaveChanges=False)
            raise RuntimeError("File đang mở trong Excel, đóng rồi chạy lại để "
                               "tính lại công thức")
        app.CalculateFullRebuild()
        found, nfor = {}, 0
        for ws in wb.Worksheets:
            used = ws.UsedRange
            if used.Count == 1 and used.Value is None:
                continue
            try:
                nfor += ws.Cells.SpecialCells(-4123).Count      # xlCellTypeFormulas
            except Exception:
                pass
            for e in ERRS:
                try:
                    hits = used.Find(e, LookIn=-4163, LookAt=2)  # xlValues, xlWhole
                except Exception:
                    hits = None
                if hits is not None:
                    first, locs = hits.Address, []
                    while True:
                        locs.append(f"{ws.Name}!{hits.Address(False, False)}")
                        hits = used.FindNext(hits)
                        if hits is None or hits.Address == first or len(locs) > 50:
                            break
                    found.setdefault(e, {"count": 0, "locations": []})
                    found[e]["count"] += len(locs)
                    found[e]["locations"] += locs[:10]
        if wb.ReadOnly:
            # máy này Excel mở file có sẵn ở chế độ chỉ đọc, Save sẽ bật hộp thoại
            # Save As và treo; SaveCopyAs ghi được nên chép đè lại file gốc
            tmp = path + ".recalc.tmp.xlsx"
            if os.path.exists(tmp):
                os.remove(tmp)
            wb.SaveCopyAs(tmp)
            wb.Close(SaveChanges=False)
            os.replace(tmp, path)
        else:
            wb.Save()
            wb.Close(SaveChanges=True)
        return {"status": "errors_found" if found else "success",
                "total_errors": sum(v["count"] for v in found.values()),
                "total_formulas": nfor, "error_summary": found}
    finally:
        app.Quit()


if __name__ == "__main__":
    print(json.dumps(recalc(sys.argv[1]), ensure_ascii=False, indent=2))
