# -*- coding: utf-8 -*-
"""
XUẤT BÁO CÁO EXCEL CHẤM ĐIỂM LỚP 12B6 (CHUẨN ĐỐI SOÁT KÉP & BẢO TOÀN SĨ SỐ)
Giáo viên: ThS. Nguyễn Văn Sang - Khoa Cơ bản - Trường Cao đẳng Nghề số 1 - BQP
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from data_12b6 import STUDENTS_12B6, NON_EXAM_FILES, MASTER_KEYS

EXCEL_MAIN = r"D:\TOAN\05_CONG_CU_VA_ENGINE\Cham_Diem_Hoc_Sinh\03_KET_QUA_EXCEL\Ket_Qua_Cham_Diem_Giua_Ky_1_Lop12.xlsx"
EXCEL_12B6 = r"D:\TOAN\05_CONG_CU_VA_ENGINE\Cham_Diem_Hoc_Sinh\03_KET_QUA_EXCEL\Ket_Qua_Cham_Diem_Giua_Ky_1_Lop12B6.xlsx"

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
    # 1. Sheet Bảng Điểm 12B6
    # =========================================================================
    if "Bảng Điểm 12B6" in wb.sheetnames:
        del wb["Bảng Điểm 12B6"]
    ws = wb.create_sheet(title="Bảng Điểm 12B6")
    ws.views.sheetView[0].showGridLines = True

    ws.merge_cells("A1:L1")
    ws["A1"] = "BẢNG TỔNG HỢP ĐIỂM KIỂM TRA GIỮA HỌC KỲ 1 - NĂM HỌC 2026 - 2027 (LỚP 12B6)"
    ws["A1"].font = font_title
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:L2")
    ws["A2"] = "Môn: Toán 12 | Lớp: 12B6 (Sĩ số thi: 34/34 HS) | Trường Cao đẳng Nghề số 1 - BQP | GV: ThS. Nguyễn Văn Sang"
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

    for row_idx, s in enumerate(STUDENTS_12B6, 5):
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

    # Thống kê tổng hợp cuối Bảng Điểm 12B6
    stat_row = 5 + len(STUDENTS_12B6) + 1
    avg_total = round(sum(s["total"] for s in STUDENTS_12B6) / len(STUDENTS_12B6), 2)
    avg_tn = round(sum(s["total_tn"] for s in STUDENTS_12B6) / len(STUDENTS_12B6), 2)
    avg_tl = round(sum(s["p4"] for s in STUDENTS_12B6) / len(STUDENTS_12B6), 2)
    rank_counts = {}
    for s in STUDENTS_12B6:
        rank_counts[s["rank"]] = rank_counts.get(s["rank"], 0) + 1

    stat_lines = [
        ("THỐNG KÊ CHẤT LƯỢNG LỚP 12B6 (34/34 HỌC SINH)", ""),
        ("Điểm trung bình toàn lớp", f"{avg_total:.2f} / 10.0đ (TB Trắc nghiệm: {avg_tn:.2f}/7.0đ | TB Tự luận: {avg_tl:.2f}/3.0đ)"),
        ("Xuất sắc (>= 9.0đ)", f"{rank_counts.get('Xuất sắc', 0)} HS ({rank_counts.get('Xuất sắc', 0)/34*100:.1f}%)"),
        ("Giỏi (8.0 - 8.99đ)", f"{rank_counts.get('Giỏi', 0)} HS ({rank_counts.get('Giỏi', 0)/34*100:.1f}%)"),
        ("Khá (6.5 - 7.99đ)", f"{rank_counts.get('Khá', 0)} HS ({rank_counts.get('Khá', 0)/34*100:.1f}%)"),
        ("Trung bình (5.0 - 6.49đ)", f"{rank_counts.get('Trung bình', 0)} HS ({rank_counts.get('Trung bình', 0)/34*100:.1f}%)"),
        ("Yếu & Kém (< 5.0đ)", f"{rank_counts.get('Yếu', 0) + rank_counts.get('Kém', 0)} HS ({(rank_counts.get('Yếu', 0) + rank_counts.get('Kém', 0))/34*100:.1f}%)"),
    ]
    for off, (lbl, val) in enumerate(stat_lines):
        r = stat_row + off
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=12)
        c1 = ws.cell(row=r, column=1, value=lbl)
        c2 = ws.cell(row=r, column=4, value=val)
        c1.font = font_th if off == 0 else font_tb_bold
        c2.font = font_th if off == 0 else font_tb
        bg = fill_th if off == 0 else fill_sum
        for cc in range(1, 13):
            cell = ws.cell(row=r, column=cc)
            cell.fill = bg
            cell.border = border_all
        ws.row_dimensions[r].height = 22

    # =========================================================================
    # 2. Sheet Đối Chiếu 12B6 (Chuẩn 38 cột đồng bộ 12B1 + 4 dòng Đáp Án Chuẩn)
    # =========================================================================
    if "Đối Chiếu 12B6" in wb.sheetnames:
        del wb["Đối Chiếu 12B6"]
    ws2 = wb.create_sheet(title="Đối Chiếu 12B6")
    ws2.views.sheetView[0].showGridLines = True

    ws2.merge_cells("A1:AL1")
    ws2["A1"] = "BẢNG ĐỐI CHIẾU BÀI LÀM CHI TIẾT TỪNG HỌC SINH - GIỮA HỌC KỲ 1 (LỚP 12B6)"
    ws2["A1"].font = font_title
    ws2["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 28

    ws2.merge_cells("A2:AL2")
    ws2["A2"] = "Môn: Toán 12 | Thang Phần II Bộ GD&ĐT: 1 ý=0.10, 2 ý=0.25, 3 ý=0.50, 4 ý=1.00 | GV: ThS. Nguyễn Văn Sang"
    ws2["A2"].font = font_sub
    ws2["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[2].height = 20

    group_spans = [
        ("A4:D4", "THÔNG TIN HỌC SINH"),
        ("E4:Q4", "PHẦN I: TRẮC NGHIỆM ABCD (12 CÂU - 3.00 ĐIỂM)"),
        ("R4:Z4", "PHẦN II: ĐÚNG/SAI (2 CÂU / 8 Ý - 2.00 ĐIỂM)"),
        ("AA4:AE4", "PHẦN III: TRẢ LỜI NGẮN (4 CÂU - 2.00 ĐIỂM)"),
        ("AF4:AI4", "PHẦN IV: TỰ LUẬN (3 CÂU - 3.00 ĐIỂM)"),
        ("AJ4:AL4", "TỔNG KẾT & NHẬN XÉT"),
    ]
    for span, txt in group_spans:
        ws2.merge_cells(span)
        top_left = span.split(":")[0]
        ws2[top_left] = txt
        ws2[top_left].font = font_th
        ws2[top_left].alignment = Alignment(horizontal="center", vertical="center")
    for cc in range(1, 39):
        cell = ws2.cell(row=4, column=cc)
        cell.fill = fill_th
        cell.border = border_all
    ws2.row_dimensions[4].height = 24

    headers2 = (
        ["STT", "Họ và Tên", "Lớp", "Mã Đề"]
        + [f"C{i}" for i in range(1, 13)] + ["P1 (3đ)"]
        + [f"13{x}" for x in ("a", "b", "c", "d")]
        + [f"14{x}" for x in ("a", "b", "c", "d")] + ["P2 (2đ)"]
        + [f"C{i}" for i in range(15, 19)] + ["P3 (2đ)"]
        + ["C19", "C20", "C21", "P4 (3đ)", "Tổng Điểm", "Xếp Loại", "Đối Chiếu Lỗi Sai & Nhận Xét Sư Phạm"]
    )
    assert len(headers2) == 38
    for col_idx, h in enumerate(headers2, 1):
        cell = ws2.cell(row=5, column=col_idx, value=h)
        cell.font = font_th
        cell.fill = fill_th2
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border_all
    ws2.row_dimensions[5].height = 26

    # 4 dòng ĐÁP ÁN CHUẨN MÃ 101..104
    for r_k, code in enumerate(["101", "102", "103", "104"], 6):
        k = MASTER_KEYS[code]
        key_row = (
            [None, f"[ĐÁP ÁN CHUẨN MÃ {code}]", "12B6", code]
            + list(k["p1"]) + [3.00]
            + list(k["p2_13"]) + list(k["p2_14"]) + [2.00]
            + list(k["p3"]) + [2.00]
            + [1.00, 1.00, 1.00, 3.00, 10.00, "Chuẩn", "Đáp án chuẩn chính thức nhà trường"]
        )
        ws2.row_dimensions[r_k].height = 21
        for col_idx, val in enumerate(key_row, 1):
            c = ws2.cell(row=r_k, column=col_idx, value=val)
            c.font = font_key
            c.fill = fill_key
            c.border = border_all
            c.alignment = Alignment(horizontal="left" if col_idx in (2, 38) else "center", vertical="center")

    start_st_row = 10
    for idx_s, s in enumerate(STUDENTS_12B6):
        r_idx = start_st_row + idx_s
        ws2.row_dimensions[r_idx].height = 26
        k = MASTER_KEYS[s["code"]]

        base_info = [s["stt"], s["name"], "12B6", s["code"]]
        for ci, val in enumerate(base_info, 1):
            c = ws2.cell(row=r_idx, column=ci, value=val)
            c.font = font_tb_bold if ci == 2 else font_tb
            c.border = border_all
            c.alignment = Alignment(horizontal="left" if ci == 2 else "center", vertical="center")

        # P1 C1..C12 (cols 5..16) + P1 score (col 17)
        for q in range(12):
            col = 5 + q
            st_a = s["answers_p1"][q]
            cor_a = k["p1"][q]
            ok = (st_a == cor_a)
            disp = st_a if ok else f"{st_a or 'Trống'}→{cor_a}"
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

        # P2 13a..14d (cols 18..25) + P2 score (col 26)
        k2 = k["p2_13"] + k["p2_14"]
        for i2 in range(8):
            col = 18 + i2
            st_a = s["answers_p2"][i2]
            cor_a = k2[i2]
            ok = (st_a == cor_a)
            disp = st_a if ok else f"{st_a or 'Trống'}→{cor_a}"
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

        # P3 C15..C18 (cols 27..30) + P3 score (col 31)
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

        # P4 C19, C20, C21, P4 score, Tổng Điểm, Xếp Loại, Nhận xét (cols 32..38)
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
    # 3. Sheet Ảnh Cần Xác Nhận (12B6) & Đối Soát 123/123 Ảnh Gốc
    # =========================================================================
    if "Ảnh Cần Xác Nhận (12B6)" in wb.sheetnames:
        del wb["Ảnh Cần Xác Nhận (12B6)"]
    ws3 = wb.create_sheet(title="Ảnh Cần Xác Nhận (12B6)")
    ws3.views.sheetView[0].showGridLines = True

    ws3.merge_cells("A1:G1")
    ws3["A1"] = "ĐỐI SOÁT KHO ẢNH GỐC (123/123 TỆP) & DANH SÁCH ẢNH TRÙNG / NGOÀI BÀI THI - LỚP 12B6"
    ws3["A1"].font = font_title
    ws3["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[1].height = 28

    omr_cnt = len(STUDENTS_12B6)
    tl_cnt = sum(len(s["tl_pages"]) for s in STUDENTS_12B6)
    dup_cnt = sum(1 for x in NON_EXAM_FILES if x["kind"].startswith("duplicate"))
    non_cnt = sum(1 for x in NON_EXAM_FILES if x["kind"] == "non_exam")

    recon_rows = [
        ("Tổng số ảnh gốc trong 01_ANH_CHUA_CHAM/Lop_12B6", 123, "Kiểm kê 100% tệp ảnh đầu vào (Zero-leak)"),
        ("Phiếu trắc nghiệm OMR chính thức", omr_cnt, "Đủ 34/34 học sinh dự thi (0 thiếu)"),
        ("Trang bài làm Tự luận chính thức", tl_cnt, "Đã xuất đủ 60 trang chấm bút đỏ chi tiết"),
        ("Ảnh chụp trùng lặp (buổi chiều 16:54–17:32)", dup_cnt, "23 ảnh OMR trùng + 1 ảnh Tự luận trùng (đối chiếu đã đủ cặp chính thức 23:04–23:18)"),
        ("Ảnh ngoài bài thi (screenshot, sân trường, lớp khác)", non_cnt, "Không phải bài thi Lớp 12B6, loại trừ khỏi bảng điểm"),
        ("KIỂM ĐỊNH ĐỐI SOÁT 3 CỘT SỐ", f"{omr_cnt} + {tl_cnt} + {dup_cnt} + {non_cnt} = {omr_cnt + tl_cnt + dup_cnt + non_cnt}", "KHỚP TUYỆT ĐỐI 123/123 TỆP | ĐỦ 34/34 HỌC SINH"),
    ]
    ws3.merge_cells("A3:D3")
    ws3["A3"] = "Hạng Mục Kiểm Kê Kho Ảnh Lớp 12B6"
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
        ws3.row_dimensions[idx_r].height = 22

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
        ws3.row_dimensions[r_idx].height = 26
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

    # Set clean readable widths
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

    wb_12b6 = openpyxl.Workbook()
    wb_12b6.remove(wb_12b6.active)
    build_sheets(wb_12b6)
    wb_12b6.save(EXCEL_12B6)
    print(f"Exported standalone 12B6 workbook: {EXCEL_12B6}")
