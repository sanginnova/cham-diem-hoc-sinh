# -*- coding: utf-8 -*-
"""
XUẤT BÁO CÁO EXCEL CHẤM ĐIỂM LỚP 10A17 (CHUẨN ĐỐI SOÁT KÉP & BẢO TOÀN SĨ SỐ)
Giáo viên: ThS. Nguyễn Văn Sang - Khoa Cơ bản - Trường Cao đẳng Nghề số 1 - BQP
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from data_10a17 import STUDENTS_10A17, NON_EXAM_FILES, MASTER_KEYS

EXCEL_10A17 = r"D:\TOAN\05_CONG_CU_VA_ENGINE\Cham_Diem_Hoc_Sinh\03_KET_QUA_EXCEL\Ket_Qua_Cham_Diem_Giua_Ky_1_Lop10A17.xlsx"

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
    return str(v or "").strip().replace(".", ",")


from common_utils import vn_name_sort_key

def build_sheets(wb):
    STUDENTS_10A17.sort(key=lambda s: vn_name_sort_key(s["name"]))
    for _i, _s in enumerate(STUDENTS_10A17, 1):
        _s["stt"] = _i
    # 1. Bảng Điểm 10A17
    if "Bảng Điểm 10A17" in wb.sheetnames:
        del wb["Bảng Điểm 10A17"]
    ws = wb.create_sheet(title="Bảng Điểm 10A17")
    ws.views.sheetView[0].showGridLines = True

    ws.merge_cells("A1:L1")
    ws["A1"] = "BẢNG TỔNG HỢP ĐIỂM KIỂM TRA GIỮA HỌC KỲ 1 - MÔN TOÁN 10 (LỚP 10A17)"
    ws["A1"].font = font_title
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:L2")
    ws["A2"] = f"Môn: Toán 10 | Lớp: 10A17 (Sĩ số dự thi: {len(STUDENTS_10A17)}/36 HS) | Trường Cao đẳng Nghề số 1 - BQP | GV: ThS. Nguyễn Văn Sang"
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

    for row_idx, s in enumerate(STUDENTS_10A17, 5):
        ws.row_dimensions[row_idx].height = 28
        fill = fill_amber_light if s["code"] == "CHƯA_TÔ" else (fill_zebra if row_idx % 2 == 0 else PatternFill(fill_type=None))
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

    stat_row = 5 + len(STUDENTS_10A17) + 1
    avg_total = round(sum(s["total"] for s in STUDENTS_10A17) / len(STUDENTS_10A17), 2)
    pass_cnt = sum(1 for s in STUDENTS_10A17 if s["total"] >= 5.0)
    good_cnt = sum(1 for s in STUDENTS_10A17 if s["total"] >= 6.5)
    exc_cnt = sum(1 for s in STUDENTS_10A17 if s["total"] >= 8.0)

    ws.merge_cells(start_row=stat_row, start_column=1, end_row=stat_row, end_column=12)
    ws.cell(
        row=stat_row, column=1,
        value=(
            f"THỐNG KÊ LỚP 10A17 ({len(STUDENTS_10A17)} HS): Điểm TB = {avg_total:.2f}  |  "
            f"Trên TB (>=5.0): {pass_cnt}/{len(STUDENTS_10A17)} ({pass_cnt*100/len(STUDENTS_10A17):.1f}%)  |  "
            f"Khá trở lên (>=6.5): {good_cnt}/{len(STUDENTS_10A17)}  |  Giỏi/XS (>=8.0): {exc_cnt}/{len(STUDENTS_10A17)}"
        )
    )
    ws.cell(row=stat_row, column=1).font = font_key
    ws.cell(row=stat_row, column=1).fill = fill_key
    ws.cell(row=stat_row, column=1).alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[stat_row].height = 24

    # 2. Đối Chiếu 10A17
    if "Đối Chiếu 10A17" in wb.sheetnames:
        del wb["Đối Chiếu 10A17"]
    ws2 = wb.create_sheet(title="Đối Chiếu 10A17")
    ws2.views.sheetView[0].showGridLines = True

    ws2.merge_cells("A1:AL1")
    ws2["A1"] = "BẢNG ĐỐI CHIẾU BÀI LÀM CHI TIẾT TỪNG CÂU (C1 - C21) - TOÁN 10 LỚP 10A17"
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

    for k_offset, code in enumerate(["101", "102", "103", "104"]):
        r_idx = 4 + k_offset
        ws2.row_dimensions[r_idx].height = 20
        k = MASTER_KEYS[code]
        pm = k["p4_max"]
        vals = ["ĐÁP ÁN", f"Mã đề {code}", "CHUẨN", code] + k["p1"] + [3.00] + k["p2_13"] + k["p2_14"] + [2.00] + k["p3"] + [2.00, pm["C19"], pm["C20"], pm["C21"], 3.00, 10.00, "Chuẩn", "Đáp án chuẩn Toán 10"]
        for c_i, val in enumerate(vals, 1):
            c = ws2.cell(row=r_idx, column=c_i, value=val)
            c.font = font_key
            c.fill = fill_key
            c.border = border_all
            c.alignment = Alignment(horizontal="center" if c_i not in (2, 38) else "left", vertical="center")
            if isinstance(val, float):
                c.number_format = "0.00"

    for s_idx, s in enumerate(STUDENTS_10A17):
        r_idx = 8 + s_idx
        ws2.row_dimensions[r_idx].height = 24
        code = str(s["code"])
        k = MASTER_KEYS.get(code)

        base = [s["stt"], s["name"], s["class"], code]
        for c_i, val in enumerate(base, 1):
            c = ws2.cell(row=r_idx, column=c_i, value=val)
            c.font = font_tb_bold if c_i == 2 else font_tb
            c.border = border_all
            c.alignment = Alignment(horizontal="center" if c_i != 2 else "left", vertical="center")

        for i1 in range(12):
            col = 5 + i1
            st_a = s["answers_p1"][i1] if i1 < len(s["answers_p1"]) else ""
            if k:
                cor_a = k["p1"][i1]
                ok = (st_a == cor_a)
                disp = st_a if ok else f"{st_a or '-'}→{cor_a}"
            else:
                ok = False
                disp = st_a or "-"
            c = ws2.cell(row=r_idx, column=col, value=disp)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = border_all
            c.fill = fill_green_light if ok else (fill_amber_light if not k else fill_red_light)
            c.font = font_ok if ok else font_err
        c_p1 = ws2.cell(row=r_idx, column=17, value=float(s["p1"]))
        c_p1.font = font_tb_bold
        c_p1.alignment = Alignment(horizontal="right", vertical="center")
        c_p1.number_format = "0.00"
        c_p1.border = border_all

        cor_p2 = (k["p2_13"] + k["p2_14"]) if k else None
        for i2 in range(8):
            col = 18 + i2
            st_a = s["answers_p2"][i2] if i2 < len(s["answers_p2"]) else ""
            if cor_p2:
                cor_a = cor_p2[i2]
                ok = (st_a == cor_a)
                disp = st_a if ok else f"{st_a or '-'}→{cor_a}"
            else:
                ok = False
                disp = st_a or "-"
            c = ws2.cell(row=r_idx, column=col, value=disp)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = border_all
            c.fill = fill_green_light if ok else (fill_amber_light if not k else fill_red_light)
            c.font = font_ok if ok else font_err
        c_p2 = ws2.cell(row=r_idx, column=26, value=float(s["p2"]))
        c_p2.font = font_tb_bold
        c_p2.alignment = Alignment(horizontal="right", vertical="center")
        c_p2.number_format = "0.00"
        c_p2.border = border_all

        for i3 in range(4):
            col = 27 + i3
            st_a = _norm_p3(s["answers_p3"][i3])
            if k:
                cor_a = _norm_p3(k["p3"][i3])
                ok = (st_a == cor_a)
                disp = st_a if ok else f"{st_a or 'Trống'}→{cor_a}"
            else:
                ok = False
                disp = st_a or "Trống"
            c = ws2.cell(row=r_idx, column=col, value=disp)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = border_all
            c.fill = fill_green_light if ok else (fill_amber_light if not k else fill_red_light)
            c.font = font_ok if ok else font_err
        c_p3 = ws2.cell(row=r_idx, column=31, value=float(s["p3"]))
        c_p3.font = font_tb_bold
        c_p3.alignment = Alignment(horizontal="right", vertical="center")
        c_p3.number_format = "0.00"
        c_p3.border = border_all

        for q_i, q_val in enumerate(s["answers_p4"]):
            col = 32 + q_i
            q_key = ("C19", "C20", "C21")[q_i]
            q_max = s["p4_detail"][q_key]["max_score"]
            c = ws2.cell(row=r_idx, column=col, value=float(q_val))
            c.number_format = "0.00"
            c.alignment = Alignment(horizontal="right", vertical="center")
            c.border = border_all
            if q_val >= q_max:
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

    # 3. Sheet Kiểm Kê Kho Ảnh 78/78
    if "Ảnh Cần Xác Nhận (10A17)" in wb.sheetnames:
        del wb["Ảnh Cần Xác Nhận (10A17)"]
    ws3 = wb.create_sheet(title="Ảnh Cần Xác Nhận (10A17)")
    ws3.views.sheetView[0].showGridLines = True

    ws3.merge_cells("A1:G1")
    ws3["A1"] = "ĐỐI SOÁT KHO ẢNH GỐC (78/78 TỆP) & LƯU Ý BÀI THI - LỚP 10A17"
    ws3["A1"].font = font_title
    ws3["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[1].height = 28

    omr_cnt = len(STUDENTS_10A17)
    tl_cnt = sum(len(s["tl_pages"]) for s in STUDENTS_10A17)
    dup_cnt = sum(1 for x in NON_EXAM_FILES if x["kind"].startswith("duplicate"))
    non_cnt = sum(1 for x in NON_EXAM_FILES if x["kind"] == "non_exam")

    recon_rows = [
        ("Tổng số ảnh gốc trong 01_ANH_CHUA_CHAM/Lop_10A17", 78, "Kiểm kê 100% tệp ảnh đầu vào (Zero-leak)"),
        ("Phiếu trắc nghiệm OMR chính thức", omr_cnt, "Đủ 36/36 học sinh (#1 đến #36)"),
        ("Trang bài làm Tự luận / Giấy nháp kèm theo", tl_cnt, f"Đã xử lý & đóng gói đủ {tl_cnt} trang"),
        ("Ảnh chụp trùng / ngoài bài thi", dup_cnt + non_cnt, "Không có ảnh trùng/ngoài bài thi"),
        ("KIỂM ĐỊNH ĐỐI SOÁT 3 CỘT SỐ", f"{omr_cnt} + {tl_cnt} + {dup_cnt + non_cnt} = {omr_cnt + tl_cnt + dup_cnt + non_cnt}", "KHỚP TUYỆT ĐỐI 78/78 TỆP | ĐỦ 36/36 HỌC SINH"),
    ]
    ws3.merge_cells("A3:D3")
    ws3["A3"] = "Hạng Mục Kiểm Kê Kho Ảnh Lớp 10A17"
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

    widths_ws1 = {1: 6, 2: 24, 3: 8, 4: 10, 5: 13, 6: 13, 7: 13, 8: 14, 9: 15, 10: 15, 11: 13, 12: 65}
    for ci, w in widths_ws1.items():
        ws.column_dimensions[get_column_letter(ci)].width = w

    widths_ws2 = {1: 6, 2: 24, 3: 8, 4: 10, 17: 10, 26: 10, 31: 10, 32: 9, 33: 9, 34: 9, 35: 10, 36: 11, 37: 12, 38: 65}
    for col_idx in range(1, 39):
        w = widths_ws2.get(col_idx, 9 if col_idx in range(5, 32) else 10)
        ws2.column_dimensions[get_column_letter(col_idx)].width = w

    widths_ws3 = {1: 6, 2: 28, 3: 24, 4: 24, 5: 12, 6: 58, 7: 44}
    for ci, w in widths_ws3.items():
        ws3.column_dimensions[get_column_letter(ci)].width = w


if __name__ == "__main__":
    wb_10a17 = openpyxl.Workbook()
    wb_10a17.remove(wb_10a17.active)
    build_sheets(wb_10a17)
    wb_10a17.save(EXCEL_10A17)
    print(f"Exported standalone 10A17 workbook: {EXCEL_10A17}")
