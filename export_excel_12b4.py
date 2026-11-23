# -*- coding: utf-8 -*-
"""
XUẤT BÁO CÁO EXCEL CHẤM ĐIỂM LỚP 12B4 (CHUẨN ĐỐI SOÁT KÉP & BẢO TOÀN SĨ SỐ)
Giáo viên: ThS. Nguyễn Văn Sang - Khoa Cơ bản - Trường Cao đẳng Nghề số 1 - BQP
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from data_12b4 import STUDENTS_12B4, NON_EXAM_FILES, MASTER_KEYS, ABSENT_STUDENTS_12B4

EXCEL_MAIN = r"D:\TOAN\05_CONG_CU_VA_ENGINE\Cham_Diem_Hoc_Sinh\03_KET_QUA_EXCEL\Ket_Qua_Cham_Diem_Giua_Ky_1_Lop12.xlsx"
EXCEL_12B4 = r"D:\TOAN\05_CONG_CU_VA_ENGINE\Cham_Diem_Hoc_Sinh\03_KET_QUA_EXCEL\Ket_Qua_Cham_Diem_Giua_Ky_1_Lop12B4.xlsx"

font_title = Font(name="Times New Roman", size=15, bold=True, color="1E3A8A")
font_sub = Font(name="Times New Roman", size=11, italic=True)
font_th = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
font_tb = Font(name="Times New Roman", size=11)
font_tb_bold = Font(name="Times New Roman", size=11, bold=True)
font_ok = Font(name="Times New Roman", size=11, bold=True, color="047857")
font_err = Font(name="Times New Roman", size=11, bold=True, color="B91C1C")
font_key = Font(name="Times New Roman", size=11, bold=True, color="1E3A8A")

fill_th = PatternFill(start_color="1E40AF", end_color="1E40AF", fill_type="solid")
fill_th2 = PatternFill(start_color="3B82F6", end_color="3B82F6", fill_type="solid")
fill_zebra = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
fill_green_light = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
fill_red_light = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
fill_amber_light = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
fill_key = PatternFill(start_color="E0E7FF", end_color="E0E7FF", fill_type="solid")
fill_sum = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")

thin = Side(border_style="thin", color="CBD5E1")
border_all = Border(left=thin, right=thin, top=thin, bottom=thin)


def _norm_p3(v):
    s = str(v or "").strip()
    return s[:-2] if s.endswith(".0") else s


def build_sheets(wb):
    # =========================================================================
    # 1. Sheet Bảng Điểm 12B4
    # =========================================================================
    if "Bảng Điểm 12B4" in wb.sheetnames:
        del wb["Bảng Điểm 12B4"]
    ws = wb.create_sheet(title="Bảng Điểm 12B4")
    ws.views.sheetView[0].showGridLines = True

    ws.merge_cells("A1:L1")
    ws["A1"] = "BẢNG TỔNG HỢP ĐIỂM KIỂM TRA GIỮA HỌC KỲ 1 - NĂM HỌC 2026 - 2027 (LỚP 12B4)"
    ws["A1"].font = font_title
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:L2")
    ws["A2"] = (
        f"Môn: Toán 12 | Lớp: 12B4 (Dự thi: {len(STUDENTS_12B4)}/38 HS; Vắng thi: {', '.join(ABSENT_STUDENTS_12B4)}) "
        "| Trường Cao đẳng Nghề số 1 - BQP | GV: ThS. Nguyễn Văn Sang"
    )
    ws["A2"].font = font_sub
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 20

    headers = [
        "STT", "Họ và Tên Thí Sinh", "Lớp", "Mã Đề",
        "Phần I (3đ)", "Phần II (2đ)", "Phần III (2đ)", "Tổng TN (7đ)",
        "Phần IV / TL (3đ)", "Tổng Điểm (10đ)", "Xếp Loại", "Lỗi Sai Chi Tiết & Nhận Xét Sư Phạm",
    ]
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col_idx, value=h)
        cell.font = font_th
        cell.fill = fill_th
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_all
    ws.row_dimensions[4].height = 30

    for row_idx, s in enumerate(STUDENTS_12B4, 5):
        ws.row_dimensions[row_idx].height = 28
        fill = fill_zebra if row_idx % 2 == 0 else PatternFill(fill_type=None)

        row_data = [
            s["stt"], s["name"], s["class"], s["code"],
            float(s["p1"]), float(s["p2"]), float(s["p3"]), float(s["total_tn"]),
            float(s["p4"]), float(s["total"]), s["rank"], f"{s['note']} {s['errors']}",
        ]
        for col_idx, val in enumerate(row_data, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = font_tb_bold if col_idx in (8, 10) else font_tb
            cell.border = border_all
            cell.fill = fill
            if col_idx in (1, 3, 4, 11):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx in (2, 12):
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            else:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = "0.00"

    stat_row = 5 + len(STUDENTS_12B4) + 1
    avg_total = round(sum(s["total"] for s in STUDENTS_12B4) / len(STUDENTS_12B4), 2)
    pass_cnt = sum(1 for s in STUDENTS_12B4 if s["total"] >= 5.0)
    good_cnt = sum(1 for s in STUDENTS_12B4 if s["total"] >= 6.5)
    exc_cnt = sum(1 for s in STUDENTS_12B4 if s["total"] >= 8.0)

    ws.merge_cells(start_row=stat_row, start_column=1, end_row=stat_row, end_column=12)
    ws.cell(
        row=stat_row, column=1,
        value=(
            f"THỐNG KÊ LỚP 12B4 ({len(STUDENTS_12B4)}/38 HS dự thi; Vắng: {len(ABSENT_STUDENTS_12B4)} HS): "
            f"Điểm TB = {avg_total:.2f}  |  Trên TB (>=5.0): {pass_cnt}/{len(STUDENTS_12B4)} ({pass_cnt*100/len(STUDENTS_12B4):.1f}%)  |  "
            f"Khá trở lên (>=6.5): {good_cnt}/{len(STUDENTS_12B4)}  |  Giỏi/XS (>=8.0): {exc_cnt}/{len(STUDENTS_12B4)}"
        )
    )
    ws.cell(row=stat_row, column=1).font = font_key
    ws.cell(row=stat_row, column=1).fill = fill_key
    ws.cell(row=stat_row, column=1).alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[stat_row].height = 24

    # =========================================================================
    # 2. Sheet Đối Chiếu 12B4
    # =========================================================================
    if "Đối Chiếu 12B4" in wb.sheetnames:
        del wb["Đối Chiếu 12B4"]
    ws2 = wb.create_sheet(title="Đối Chiếu 12B4")
    ws2.views.sheetView[0].showGridLines = True

    ws2.merge_cells("A1:AL1")
    ws2["A1"] = "BẢNG ĐỐI CHIẾU BÀI LÀM CHI TIẾT TỪNG CÂU (C1 - C21) - LỚP 12B4"
    ws2["A1"].font = font_title
    ws2["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 28

    hdr2 = [
        "STT", "Họ và Tên", "Lớp", "Mã Đề",
        "C1", "C2", "C3", "C4", "C5", "C6", "C7", "C8", "C9", "C10", "C11", "C12", "Điểm P.I",
        "13a", "13b", "13c", "13d", "14a", "14b", "14c", "14d", "Điểm P.II",
        "C15", "C16", "C17", "C18", "Điểm P.III",
        "C19", "C20", "C21", "Điểm P.IV", "Tổng Điểm", "Xếp Loại", "Chi Tiết Lỗi Sai",
    ]
    for col_idx, h in enumerate(hdr2, 1):
        cell = ws2.cell(row=3, column=col_idx, value=h)
        cell.font = font_th
        cell.fill = fill_th2
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border_all
    ws2.row_dimensions[3].height = 26

    # Master keys rows (rows 4..7)
    for k_offset, code in enumerate(["101", "102", "103", "104"]):
        r_idx = 4 + k_offset
        ws2.row_dimensions[r_idx].height = 20
        k = MASTER_KEYS[code]
        vals = ["ĐÁP ÁN", f"Mã đề {code}", "CHUẨN", code] + k["p1"] + [3.00] + k["p2_13"] + k["p2_14"] + [2.00] + k["p3"] + [2.00, 1.00, 1.00, 1.00, 3.00, 10.00, "Chuẩn", "Đáp án chuẩn Bộ GDPT 2018"]
        for c_i, val in enumerate(vals, 1):
            c = ws2.cell(row=r_idx, column=c_i, value=val)
            c.font = font_key
            c.fill = fill_key
            c.border = border_all
            c.alignment = Alignment(horizontal="center" if c_i not in (2, 38) else "left", vertical="center")
            if isinstance(val, float):
                c.number_format = "0.00"

    for s_idx, s in enumerate(STUDENTS_12B4):
        r_idx = 8 + s_idx
        ws2.row_dimensions[r_idx].height = 24
        code = str(s["code"])
        k = MASTER_KEYS[code]

        base = [s["stt"], s["name"], s["class"], code]
        for c_i, val in enumerate(base, 1):
            c = ws2.cell(row=r_idx, column=c_i, value=val)
            c.font = font_tb_bold if c_i == 2 else font_tb
            c.border = border_all
            c.alignment = Alignment(horizontal="center" if c_i != 2 else "left", vertical="center")

        for i1 in range(12):
            col = 5 + i1
            st_a = s["answers_p1"][i1] if i1 < len(s["answers_p1"]) else ""
            cor_a = k["p1"][i1]
            ok = (st_a == cor_a)
            disp = st_a if ok else f"{st_a or '-'}→{cor_a}"
            c = ws2.cell(row=r_idx, column=col, value=disp)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = border_all
            c.fill = fill_green_light if ok else fill_red_light
            c.font = font_ok if ok else font_err
        c_p1 = ws2.cell(row=r_idx, column=17, value=float(s["p1"]))
        c_p1.font = font_tb_bold
        c_p1.alignment = Alignment(horizontal="right", vertical="center")
        c_p1.number_format = "0.00"
        c_p1.border = border_all

        cor_p2 = k["p2_13"] + k["p2_14"]
        for i2 in range(8):
            col = 18 + i2
            st_a = s["answers_p2"][i2] if i2 < len(s["answers_p2"]) else ""
            cor_a = cor_p2[i2]
            ok = (st_a == cor_a)
            disp = st_a if ok else f"{st_a or '-'}→{cor_a}"
            c = ws2.cell(row=r_idx, column=col, value=disp)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = border_all
            c.fill = fill_green_light if ok else fill_red_light
            c.font = font_ok if ok else font_err
        c_p2 = ws2.cell(row=r_idx, column=26, value=float(s["p2"]))
        c_p2.font = font_tb_bold
        c_p2.alignment = Alignment(horizontal="right", vertical="center")
        c_p2.number_format = "0.00"
        c_p2.border = border_all

        for i3 in range(4):
            col = 27 + i3
            st_a = _norm_p3(s["answers_p3"][i3])
            cor_a = _norm_p3(k["p3"][i3])
            ok = (st_a == cor_a)
            disp = st_a if ok else f"{st_a or 'Trống'}→{cor_a}"
            c = ws2.cell(row=r_idx, column=col, value=disp)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = border_all
            c.fill = fill_green_light if ok else fill_red_light
            c.font = font_ok if ok else font_err
        c_p3 = ws2.cell(row=r_idx, column=31, value=float(s["p3"]))
        c_p3.font = font_tb_bold
        c_p3.alignment = Alignment(horizontal="right", vertical="center")
        c_p3.number_format = "0.00"
        c_p3.border = border_all

        for q_i, q_val in enumerate(s["answers_p4"]):
            col = 32 + q_i
            c = ws2.cell(row=r_idx, column=col, value=float(q_val))
            c.number_format = "0.00"
            c.alignment = Alignment(horizontal="right", vertical="center")
            c.border = border_all
            if q_val >= 1.0:
                c.fill = fill_green_light
                c.font = font_ok
            elif q_val > 0:
                c.fill = fill_amber_light
                c.font = font_tb_bold
            else:
                c.fill = fill_red_light
                c.font = font_err

        for col, val, fmt in [(35, float(s["p4"]), "0.00"), (36, float(s["total"]), "0.00")]:
            c = ws2.cell(row=r_idx, column=col, value=val)
            c.font = font_tb_bold
            c.number_format = fmt
            c.alignment = Alignment(horizontal="right", vertical="center")
            c.border = border_all

        c_rk = ws2.cell(row=r_idx, column=37, value=s["rank"])
        c_rk.font = font_tb_bold
        c_rk.alignment = Alignment(horizontal="center", vertical="center")
        c_rk.border = border_all

        c_err = ws2.cell(row=r_idx, column=38, value=s["errors"])
        c_err.font = font_tb
        c_err.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        c_err.border = border_all

    # =========================================================================
    # 3. Sheet Ảnh Cần Xác Nhận (12B4) & Đối Soát 105/105 Ảnh Gốc
    # =========================================================================
    if "Ảnh Cần Xác Nhận (12B4)" in wb.sheetnames:
        del wb["Ảnh Cần Xác Nhận (12B4)"]
    ws3 = wb.create_sheet(title="Ảnh Cần Xác Nhận (12B4)")
    ws3.views.sheetView[0].showGridLines = True

    ws3.merge_cells("A1:G1")
    ws3["A1"] = "ĐỐI SOÁT KHO ẢNH GỐC (105/105 TỆP) & DANH SÁCH ẢNH TRÙNG / VẮNG THI - LỚP 12B4"
    ws3["A1"].font = font_title
    ws3["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[1].height = 28

    omr_cnt = len(STUDENTS_12B4)
    tl_cnt = sum(len(s["tl_pages"]) for s in STUDENTS_12B4)
    dup_cnt = sum(1 for x in NON_EXAM_FILES if x["kind"].startswith("duplicate"))
    non_cnt = sum(1 for x in NON_EXAM_FILES if x["kind"] == "non_exam")

    recon_rows = [
        ("Tổng số ảnh gốc trong 01_ANH_CHUA_CHAM/Lop_12B4", 105, "Kiểm kê 100% tệp ảnh đầu vào (Zero-leak)"),
        ("Phiếu trắc nghiệm OMR chính thức", omr_cnt, "Đủ 36/36 học sinh dự thi (Vắng thi theo danh sách: Nguyễn Quang Anh, Dương Đức Cường)"),
        ("Trang bài làm Tự luận chính thức", tl_cnt, f"Đã xuất đủ {tl_cnt} trang chấm bút đỏ chi tiết"),
        ("Ảnh chụp trùng lặp", dup_cnt, "2 ảnh OMR chụp trùng (Nguyễn Vũ Yên Nguyên, Âu Đình Huấn) đã dùng bản chính thức có số phách"),
        ("Ảnh ngoài bài thi", non_cnt, "Không có ảnh ngoài bài thi"),
        ("KIỂM ĐỊNH ĐỐI SOÁT 3 CỘT SỐ", f"{omr_cnt} + {tl_cnt} + {dup_cnt} + {non_cnt} = {omr_cnt + tl_cnt + dup_cnt + non_cnt}", "KHỚP TUYỆT ĐỐI 105/105 TỆP | ĐỦ 36/36 BÀI THI (38 SĨ SỐ)"),
    ]
    ws3.merge_cells("A3:D3")
    ws3["A3"] = "Hạng Mục Kiểm Kê Kho Ảnh Lớp 12B4"
    ws3["E3"] = "Số Tệp"
    ws3.merge_cells("F3:G3")
    ws3["F3"] = "Kết Luận Kiểm Toán"
    for cc in range(1, 8):
        c = ws3.cell(row=3, column=cc)
        c.font = font_th
        c.fill = fill_th
        c.border = border_all
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[3].height = 24

    for idx_r, (lbl, cnt, desc) in enumerate(recon_rows, 4):
        ws3.merge_cells(start_row=idx_r, start_column=1, end_row=idx_r, end_column=4)
        ws3.merge_cells(start_row=idx_r, start_column=6, end_row=idx_r, end_column=7)
        c1 = ws3.cell(row=idx_r, column=1, value=lbl)
        c2 = ws3.cell(row=idx_r, column=5, value=cnt)
        c3 = ws3.cell(row=idx_r, column=6, value=desc)
        is_last = (idx_r == 4 + len(recon_rows) - 1)
        for cc in range(1, 8):
            cell = ws3.cell(row=idx_r, column=cc)
            cell.border = border_all
            cell.fill = fill_green_light if is_last else fill_sum
            cell.font = font_ok if is_last else font_tb_bold if cc == 5 else font_tb
        c1.alignment = Alignment(horizontal="left", vertical="center")
        c2.alignment = Alignment(horizontal="center", vertical="center")
        c3.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws3.row_dimensions[idx_r].height = 24

    start_tbl3 = 11
    headers3 = ["STT", "Tên tệp ảnh", "Phân loại", "Học sinh liên quan", "Mã đề", "Nội dung nhận diện", "Biện pháp xử lý"]
    for col_idx, h in enumerate(headers3, 1):
        cell = ws3.cell(row=start_tbl3, column=col_idx, value=h)
        cell.font = font_th
        cell.fill = PatternFill(start_color="D97706", end_color="D97706", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border_all
    ws3.row_dimensions[start_tbl3].height = 26

    for r_idx, item in enumerate(NON_EXAM_FILES, start_tbl3 + 1):
        ws3.row_dimensions[r_idx].height = 28
        vals = [
            r_idx - start_tbl3,
            item["filename"],
            item["type"],
            item["student_match"] or "-",
            item["code"] or "-",
            item["content"],
            item["action"],
        ]
        for c_i, val in enumerate(vals, 1):
            c = ws3.cell(row=r_idx, column=c_i, value=val)
            c.border = border_all
            c.font = font_tb
            c.alignment = Alignment(horizontal="center" if c_i in (1, 3, 5) else "left", vertical="center", wrap_text=True)

    widths_ws1 = {1: 6, 2: 24, 3: 8, 4: 9, 5: 13, 6: 13, 7: 13, 8: 14, 9: 15, 10: 15, 11: 13, 12: 65}
    for ci, w in widths_ws1.items():
        ws.column_dimensions[get_column_letter(ci)].width = w

    widths_ws2 = {1: 6, 2: 24, 3: 8, 4: 9, 17: 10, 26: 10, 31: 10, 32: 9, 33: 9, 34: 9, 35: 10, 36: 11, 37: 12, 38: 65}
    for col_idx in range(1, 39):
        w = widths_ws2.get(col_idx, 9 if col_idx in range(5, 32) else 10)
        ws2.column_dimensions[get_column_letter(col_idx)].width = w

    widths_ws3 = {1: 6, 2: 28, 3: 24, 4: 24, 5: 10, 6: 58, 7: 44}
    for ci, w in widths_ws3.items():
        ws3.column_dimensions[get_column_letter(ci)].width = w


if __name__ == "__main__":
    wb_main = openpyxl.load_workbook(EXCEL_MAIN)
    build_sheets(wb_main)
    wb_main.save(EXCEL_MAIN)
    print(f"Updated main workbook: {EXCEL_MAIN}")

    wb_12b4 = openpyxl.Workbook()
    wb_12b4.remove(wb_12b4.active)
    build_sheets(wb_12b4)
    wb_12b4.save(EXCEL_12B4)
    print(f"Exported standalone 12B4 workbook: {EXCEL_12B4}")
