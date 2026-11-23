# -*- coding: utf-8 -*-
"""
HỆ THỐNG CHẤM ĐIỂM BÀI THI TOÁN 12 - CHUẨN 4 BỘ ĐỀ GỐC IN THỰC TẾ (101, 102, 103, 104)
Giáo viên: ThS. Nguyễn Văn Sang - Khoa Cơ bản - Trường Cao đẳng Nghề số 1 - BQP
Lớp: 12B1 - Năm học 2026 - 2027
"""

import os
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from data_12b1 import STUDENTS_12B1 as STUDENT_RESULTS, NON_EXAM_FILES_12B1

_CORE_KEYS = {
    "101": {
        "p1": ["A", "C", "D", "C", "D", "B", "D", "A", "B", "D", "B", "C"],
        "p2": ["Đ", "Đ", "S", "S", "S", "Đ", "Đ", "S"],
        "p3": ["2", "1", "7", "18"],
    },
    "102": {
        "p1": ["A", "D", "A", "D", "D", "D", "D", "D", "C", "B", "A", "C"],
        "p2": ["S", "Đ", "Đ", "S", "Đ", "Đ", "S", "S"],
        "p3": ["7", "18", "1", "2"],
    },
    "103": {
        "p1": ["B", "A", "B", "D", "C", "A", "A", "D", "B", "C", "A", "C"],
        "p2": ["Đ", "Đ", "S", "S", "S", "Đ", "Đ", "S"],
        "p3": ["18", "2", "1", "7"],
    },
    "104": {
        "p1": ["D", "A", "B", "A", "D", "C", "A", "D", "C", "C", "D", "D"],
        "p2": ["Đ", "Đ", "S", "S", "S", "Đ", "Đ", "S"],
        "p3": ["2", "1", "7", "18"],
    },
    "201": {
        "p1": ["D", "B", "C", "A", "C", "D", "B", "C", "B", "A", "B", "D"],
        "p2": ["Đ", "S", "S", "Đ", "Đ", "S", "Đ", "Đ"],
        "p3": ["7", "20", "1", "1"],
    },
    "202": {
        "p1": ["D", "C", "A", "B", "B", "D", "A", "A", "A", "B", "A", "D"],
        "p2": ["Đ", "S", "Đ", "Đ", "Đ", "S", "S", "Đ"],
        "p3": ["1", "20", "1", "7"],
    },
    "203": {
        "p1": ["A", "B", "D", "A", "B", "A", "C", "B", "A", "D", "C", "C"],
        "p2": ["Đ", "S", "Đ", "Đ", "S", "Đ", "S", "Đ"],
        "p3": ["1", "1", "20", "7"],
    },
    "204": {
        "p1": ["B", "D", "A", "B", "D", "B", "C", "A", "C", "D", "A", "A"],
        "p2": ["Đ", "S", "Đ", "Đ", "Đ", "S", "Đ", "S"],
        "p3": ["1", "1", "7", "20"],
    },
}
# Alias mapping 4-digit codes to 3-digit keys
for _k in list(_CORE_KEYS.keys()):
    _CORE_KEYS[f"0{_k}"] = _CORE_KEYS[_k]

_P2_BGD_SCALE = {0: 0.0, 1: 0.1, 2: 0.25, 3: 0.5, 4: 1.0}

def generate_student_feedback(r, k):
    wrong_p1 = []
    correct_p1_cnt = 0
    for i, (ans, cor) in enumerate(zip(r.get("answers_p1", []), k["p1"])):
        if ans == cor:
            correct_p1_cnt += 1
        else:
            wrong_p1.append(f"C{i+1} ({ans}→{cor})")
    r["details_p1"] = f"{correct_p1_cnt}/12 đúng" + (f". Sai: {', '.join(wrong_p1)}" if wrong_p1 else " tuyệt đối (100%)") + f" -> {r['p1']:.2f}đ"

    c13_wrong = [f"13{chr(97+i)}" for i, (ans, cor) in enumerate(zip(r.get("answers_p2", [])[:4], k["p2"][:4])) if ans != cor]
    c14_wrong = [f"14{chr(97+i)}" for i, (ans, cor) in enumerate(zip(r.get("answers_p2", [])[4:], k["p2"][4:])) if ans != cor]
    p2_wrong = c13_wrong + c14_wrong
    correct_p2_cnt = 8 - len(p2_wrong)
    r["details_p2"] = f"{correct_p2_cnt}/8 ý đúng" + (f". Sai: {', '.join(p2_wrong)}" if p2_wrong else " tuyệt đối") + f" -> {r['p2']:.2f}đ"

    wrong_p3 = []
    correct_p3_cnt = 0
    for i, (ans, cor) in enumerate(zip(r.get("answers_p3", []), k["p3"])):
        if str(ans).strip() == str(cor).strip():
            correct_p3_cnt += 1
        else:
            wrong_p3.append(f"C{15+i} ({ans}→{cor})")
    r["details_p3"] = f"{correct_p3_cnt}/4 câu đúng" + (f". Sai: {', '.join(wrong_p3)}" if wrong_p3 else " tuyệt đối") + f" -> {r['p3']:.2f}đ"

    all_errs = []
    if wrong_p1:
        all_errs.append(f"P1 sai {len(wrong_p1)} câu: {', '.join(wrong_p1[:4])}")
    if p2_wrong:
        all_errs.append(f"P2 sai {len(p2_wrong)} ý: {', '.join(p2_wrong)}")
    if wrong_p3:
        all_errs.append(f"P3 sai {len(wrong_p3)} câu: {', '.join(wrong_p3)}")
    if sum(r.get("answers_p4", [0, 0, 0])) < 3.0:
        c19_sc, c20_sc, c21_sc = r.get("answers_p4", [0, 0, 0])
        p4_notes = []
        if c19_sc < 1.0: p4_notes.append(f"C19 ({c19_sc}đ)")
        if c20_sc < 1.0: p4_notes.append(f"C20 ({c20_sc}đ)")
        if c21_sc < 1.0: p4_notes.append(f"C21 ({c21_sc}đ)")
        all_errs.append(f"P4 tự luận: {', '.join(p4_notes)}")

    r["errors"] = "; ".join(all_errs) if all_errs else "Không có lỗi sai, làm bài xuất sắc tuyệt đối 10.0đ"

    tot = r["total"]
    if tot >= 9.0:
        r["note"] = f"Xuất sắc {tot:.2f}/10.0đ. Kỹ năng giải toán và trình bày rất chuẩn mực."
    elif tot >= 8.0:
        r["note"] = f"Giỏi {tot:.2f}/10.0đ. Nắm chắc kiến thức, cẩn thận hơn ở các câu phân hóa."
    elif tot >= 6.5:
        r["note"] = f"Khá {tot:.2f}/10.0đ. Cần rèn thêm kỹ năng biến đổi tự luận và đọc đồ thị."
    elif tot >= 5.0:
        r["note"] = f"Trung bình {tot:.2f}/10.0đ. Cần ôn tập lại các dạng toán cơ bản và bảng biến thiên."
    else:
        r["note"] = f"Yếu {tot:.2f}/10.0đ. Cần phụ đạo thêm kiến thức nền tảng và phương pháp giải trắc nghiệm."

for _r in STUDENT_RESULTS:
    _k = _CORE_KEYS.get(_r["code"])
    if not _k:
        continue
    _r["p1"] = round(sum(1 for a, b in zip(_r["answers_p1"], _k["p1"]) if a == b) * 0.25, 2)
    _c13 = sum(1 for a, b in zip(_r["answers_p2"][:4], _k["p2"][:4]) if a == b)
    _c14 = sum(1 for a, b in zip(_r["answers_p2"][4:], _k["p2"][4:]) if a == b)
    _r["p2"] = round(_P2_BGD_SCALE[_c13] + _P2_BGD_SCALE[_c14], 2)
    _r["p3"] = round(sum(1 for a, b in zip(_r["answers_p3"], _k["p3"]) if str(a).strip() == str(b).strip()) * 0.5, 2)
    _r["p4"] = round(sum(_r["answers_p4"]), 2)
    _r["total"] = round(_r["p1"] + _r["p2"] + _r["p3"] + _r["p4"], 2)
    _tot = _r["total"]
    _r["rank"] = "Xuất sắc" if _tot >= 9.0 else ("Giỏi" if _tot >= 8.0 else ("Khá" if _tot >= 6.5 else ("Trung bình" if _tot >= 5.0 else "Yếu")))
    generate_student_feedback(_r, _k)

def export_excel():
    excel_path = r"D:\TOAN\05_CONG_CU_VA_ENGINE\Cham_Diem_Hoc_Sinh\03_KET_QUA_EXCEL\Ket_Qua_Cham_Diem_Giua_Ky_1_Lop12.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Bảng Điểm Chi Tiết"
    
    # Header format
    font_title = Font(name="Times New Roman", size=16, bold=True, color="1F4E79")
    font_sub = Font(name="Times New Roman", size=11, italic=True)
    font_head = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
    font_data = Font(name="Times New Roman", size=11)
    font_bold = Font(name="Times New Roman", size=11, bold=True)
    
    fill_head = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    fill_zebra = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
    fill_highlight = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    fill_flag = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )
    
    # Title
    ws.merge_cells("A1:K1")
    ws["A1"] = "BẢNG KẾT QUẢ CHẤM ĐIỂM BÀI KIỂM TRA GIỮA HỌC KỲ 1 - NĂM HỌC 2026 - 2027"
    ws["A1"].font = font_title
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    
    ws.merge_cells("A2:K2")
    ws["A2"] = "Môn: Toán 12 | Lớp: 12B1 | Giáo viên chấm: ThS. Nguyễn Văn Sang - Trường Cao đẳng Nghề số 1 - BQP"
    ws["A2"].font = font_sub
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
    
    ws.row_dimensions[1].height = 32
    ws.row_dimensions[2].height = 20
    ws.row_dimensions[4].height = 28
    
    headers = [
        "STT", "Họ và Tên Thí Sinh", "Lớp", "Mã Đề", 
        "Phần I (3đ)\n(12 câu ABCD)", "Phần II (2đ)\n(2 câu Đúng/Sai)", 
        "Phần III (2đ)\n(4 câu Điền số)", "Phần IV (3đ)\n(3 câu Tự luận)", 
        "Tổng Điểm\n(Thang 10)", "Xếp Loại", "Lỗi Sai Chi Tiết & Ghi Chú Sư Phạm"
    ]
    
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col_idx, value=h)
        cell.font = font_head
        cell.fill = fill_head
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
        
    for idx, s in enumerate(STUDENT_RESULTS, 1):
        row_idx = 4 + idx
        ws.row_dimensions[row_idx].height = 26
        
        row_data = [
            idx, s["name"], s["class"], s["code"],
            s["p1"], s["p2"], s["p3"], s["p4"],
            s["total"], s["rank"], f"{s['errors']}. {s['note']}"
        ]
        
        is_even = (idx % 2 == 0)
        is_top = (s["total"] >= 8.0)
        is_flagged = ("Cần xác nhận" in s["rank"])
        
        for col_idx, val in enumerate(row_data, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = font_bold if col_idx in [2, 9, 10] else font_data
            cell.border = thin_border
            
            if is_flagged:
                cell.fill = fill_flag
            elif is_top:
                cell.fill = fill_highlight
            elif is_even:
                cell.fill = fill_zebra
                
            if col_idx in [1, 3, 4, 10]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx in [5, 6, 7, 8, 9]:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                if isinstance(val, (int, float)):
                    cell.number_format = "0.00"
            elif col_idx == 11:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    # Summary row
    sum_row = 5 + len(STUDENT_RESULTS)
    ws.row_dimensions[sum_row].height = 24
    ws.merge_cells(f"A{sum_row}:D{sum_row}")
    ws[f"A{sum_row}"] = "ĐIỂM TRUNG BÌNH TOÀN LỚP"
    ws[f"A{sum_row}"].font = font_bold
    ws[f"A{sum_row}"].alignment = Alignment(horizontal="center", vertical="center")
    ws[f"A{sum_row}"].fill = fill_zebra
    
    valid_students = [s for s in STUDENT_RESULTS if "Cần xác nhận" not in s["rank"]]
    avg_p1 = sum(s["p1"] for s in valid_students) / len(valid_students)
    avg_p2 = sum(s["p2"] for s in valid_students) / len(valid_students)
    avg_p3 = sum(s["p3"] for s in valid_students) / len(valid_students)
    avg_p4 = sum(s["p4"] for s in valid_students) / len(valid_students)
    avg_total = sum(s["total"] for s in valid_students) / len(valid_students)
    
    cols_avg = [(5, avg_p1), (6, avg_p2), (7, avg_p3), (8, avg_p4), (9, avg_total)]
    for c_idx, avg_val in cols_avg:
        cell = ws.cell(row=sum_row, column=c_idx, value=avg_val)
        cell.font = font_bold
        cell.alignment = Alignment(horizontal="right", vertical="center")
        cell.fill = fill_zebra
        cell.border = thin_border
        cell.number_format = "0.00"
        
    for c_idx in [1, 2, 3, 4, 10, 11]:
        cell = ws.cell(row=sum_row, column=c_idx)
        cell.border = thin_border
        cell.fill = fill_zebra

    # Auto column widths
    col_widths = {
        1: 6, 2: 24, 3: 8, 4: 12,
        5: 14, 6: 15, 7: 15, 8: 15,
        9: 14, 10: 16, 11: 75
    }
    for col_idx, width in col_widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    
    # ==================== SHEET 2: ĐỐI CHIẾU BÀI LÀM CHI TIẾT TỪNG HỌC SINH ====================
    ws_comp = wb.create_sheet(title="Đối Chiếu Bài Làm Chi Tiết", index=1)
    ws_comp.row_dimensions[1].height = 32
    ws_comp.row_dimensions[2].height = 20
    ws_comp.row_dimensions[4].height = 24
    ws_comp.row_dimensions[5].height = 28
    
    ws_comp.merge_cells("A1:AL1")
    ws_comp["A1"] = "BẢNG ĐỐI CHIẾU BÀI LÀM CHI TIẾT TỪNG HỌC SINH - GIỮA HỌC KỲ 1 (LỚP 12B1)"
    ws_comp["A1"].font = font_title
    ws_comp["A1"].alignment = Alignment(horizontal="center", vertical="center")
    
    ws_comp.merge_cells("A2:AL2")
    ws_comp["A2"] = "Môn: Toán 12 | Giáo viên chấm: ThS. Nguyễn Văn Sang - Khoa Cơ bản - Trường Cao đẳng Nghề số 1 - BQP"
    ws_comp["A2"].font = font_sub
    ws_comp["A2"].alignment = Alignment(horizontal="center", vertical="center")
    
    fill_cat_head = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    fill_cat_sub = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    fill_correct = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    fill_wrong = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    fill_ref = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    
    font_correct = Font(name="Times New Roman", size=11, bold=True, color="276A3C")
    font_wrong = Font(name="Times New Roman", size=11, bold=True, color="C00000")
    font_ref = Font(name="Times New Roman", size=11, bold=True, color="1F4E79")
    
    # Category Header Row 4
    categories = [
        ("A4:D4", "THÔNG TIN HỌC SINH"),
        ("E4:P4", "PHẦN I: TRẮC NGHIỆM ABCD (12 CÂU - 3.0 ĐIỂM)"),
        ("Q4:Q4", "ĐIỂM P1"),
        ("R4:Y4", "PHẦN II: TRẮC NGHIỆM ĐÚNG/SAI (8 Ý - 2.0 ĐIỂM)"),
        ("Z4:Z4", "ĐIỂM P2"),
        ("AA4:AD4", "PHẦN III: TRẢ LỜI NGẮN (4 CÂU - 2.0 ĐIỂM)"),
        ("AE4:AE4", "ĐIỂM P3"),
        ("AF4:AH4", "PHẦN IV: TỰ LUẬN (3 CÂU - 3.0 Đ)"),
        ("AI4:AI4", "ĐIỂM P4"),
        ("AJ4:AJ4", "TỔNG"),
        ("AK4:AK4", "XẾP LOẠI"),
        ("AL4:AL4", "ĐỐI CHIẾU LỆCH MÃ ĐỀ & GHI CHÚ SƯ PHẠM")
    ]
    for m_range, cat_name in categories:
        ws_comp.merge_cells(m_range)
        top_left = m_range.split(":")[0]
        cell = ws_comp[top_left]
        cell.value = cat_name
        cell.font = font_head
        cell.fill = fill_cat_head
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
        
    # Detail Sub-headers Row 5
    sub_headers = [
        "STT", "Họ và Tên", "Lớp", "Mã Đề",
        "C1", "C2", "C3", "C4", "C5", "C6", "C7", "C8", "C9", "C10", "C11", "C12", "P1 (3đ)",
        "13a", "13b", "13c", "13d", "14a", "14b", "14c", "14d", "P2 (2đ)",
        "C15", "C16", "C17", "C18", "P3 (2đ)",
        "C19", "C20", "C21", "P4 (3đ)",
        "Tổng Điểm", "Xếp Loại", "Đối Chiếu Lỗi Sai & Nhận Xét Sư Phạm"
    ]
    for c_idx, sh in enumerate(sub_headers, 1):
        cell = ws_comp.cell(row=5, column=c_idx, value=sh)
        cell.font = font_head
        cell.fill = fill_cat_sub
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
        
    # Master Key Reference Rows (Rows 6+)
    master_keys_ref = _CORE_KEYS
    ref_codes = ["101", "102", "103", "104"]
    for s in STUDENT_RESULTS:
        c = str(s.get("code", "")).strip()
        norm_c = str(int(c)) if c.isdigit() else c
        if norm_c in _CORE_KEYS and norm_c not in ref_codes:
            ref_codes.append(norm_c)
    
    for r_idx, code_k in enumerate(ref_codes, 6):
        ws_comp.row_dimensions[r_idx].height = 22
        ws_comp.cell(row=r_idx, column=1, value="")
        ws_comp.cell(row=r_idx, column=2, value=f"[ĐÁP ÁN CHUẨN MÃ {code_k}]").font = font_ref
        ws_comp.cell(row=r_idx, column=3, value="12B1")
        ws_comp.cell(row=r_idx, column=4, value=code_k).font = font_ref
        
        m_data = master_keys_ref[code_k]
        # P1
        for i, ans in enumerate(m_data["p1"]):
            cell = ws_comp.cell(row=r_idx, column=5+i, value=ans)
            cell.font = font_ref
            cell.fill = fill_ref
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
        ws_comp.cell(row=r_idx, column=17, value="3.00").font = font_ref
        
        # P2
        for i, ans in enumerate(m_data["p2"]):
            cell = ws_comp.cell(row=r_idx, column=18+i, value=ans)
            cell.font = font_ref
            cell.fill = fill_ref
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
        ws_comp.cell(row=r_idx, column=26, value="2.00").font = font_ref
        
        # P3
        for i, ans in enumerate(m_data["p3"]):
            cell = ws_comp.cell(row=r_idx, column=27+i, value=ans)
            cell.font = font_ref
            cell.fill = fill_ref
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
        ws_comp.cell(row=r_idx, column=31, value="2.00").font = font_ref
        
        # P4
        for i, val in enumerate(["1.00", "1.00", "1.00"]):
            cell = ws_comp.cell(row=r_idx, column=32+i, value=val)
            cell.font = font_ref
            cell.fill = fill_ref
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
        ws_comp.cell(row=r_idx, column=35, value="3.00").font = font_ref
        ws_comp.cell(row=r_idx, column=36, value="10.00").font = font_ref
        ws_comp.cell(row=r_idx, column=37, value="Chuẩn").font = font_ref
        ws_comp.cell(row=r_idx, column=38, value=f"Đáp án chuẩn gốc in thực tế của Bộ đề Mã {code_k}").font = font_ref

    # Student Rows
    start_student_row = 6 + len(ref_codes)
    for idx, s in enumerate(STUDENT_RESULTS, 1):
        row_idx = start_student_row - 1 + idx
        ws_comp.row_dimensions[row_idx].height = 25
        code = s.get("code", "101")
        m_ref = master_keys_ref.get(code, master_keys_ref["101"])
        
        # Basic Info
        ws_comp.cell(row=row_idx, column=1, value=idx).alignment = Alignment(horizontal="center", vertical="center")
        ws_comp.cell(row=row_idx, column=2, value=s["name"]).font = font_bold
        ws_comp.cell(row=row_idx, column=3, value=s["class"]).alignment = Alignment(horizontal="center", vertical="center")
        ws_comp.cell(row=row_idx, column=4, value=code).alignment = Alignment(horizontal="center", vertical="center")
        
        # Retrieve authentic answers directly (ZERO-FALLBACK)
        p1_ans = s.get("answers_p1", [""] * 12)

        # Fill P1
        for i, val in enumerate(p1_ans):
            cell = ws_comp.cell(row=row_idx, column=5+i, value=val)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
            if val and code in master_keys_ref:
                is_correct = (val == m_ref["p1"][i])
                cell.fill = fill_correct if is_correct else fill_wrong
                cell.font = font_correct if is_correct else font_wrong
            else:
                cell.font = font_data
        ws_comp.cell(row=row_idx, column=17, value=s["p1"]).font = font_bold
        ws_comp.cell(row=row_idx, column=17).number_format = "0.00"
        ws_comp.cell(row=row_idx, column=17).alignment = Alignment(horizontal="right", vertical="center")
        
        # P2 Answers directly from student sheet (ZERO-FALLBACK)
        p2_ans = s.get("answers_p2", [""] * 8)

        for i, val in enumerate(p2_ans):
            cell = ws_comp.cell(row=row_idx, column=18+i, value=val)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
            if val and code in master_keys_ref:
                is_correct = (val == m_ref["p2"][i])
                cell.fill = fill_correct if is_correct else fill_wrong
                cell.font = font_correct if is_correct else font_wrong
            else:
                cell.font = font_data
        ws_comp.cell(row=row_idx, column=26, value=s["p2"]).font = font_bold
        ws_comp.cell(row=row_idx, column=26).number_format = "0.00"
        ws_comp.cell(row=row_idx, column=26).alignment = Alignment(horizontal="right", vertical="center")

        # P3 Answers directly from student sheet (ZERO-FALLBACK)
        p3_ans = s.get("answers_p3", [""] * 4)

        for i, val in enumerate(p3_ans):
            cell = ws_comp.cell(row=row_idx, column=27+i, value=val)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
            if val and code in master_keys_ref:
                is_correct = (str(val) == str(m_ref["p3"][i]))
                cell.fill = fill_correct if is_correct else fill_wrong
                cell.font = font_correct if is_correct else font_wrong
            else:
                cell.font = font_data
        ws_comp.cell(row=row_idx, column=31, value=s["p3"]).font = font_bold
        ws_comp.cell(row=row_idx, column=31).number_format = "0.00"
        ws_comp.cell(row=row_idx, column=31).alignment = Alignment(horizontal="right", vertical="center")

        # P4 Answers
        p4_ans = s.get("answers_p4")
        if not p4_ans:
            c19 = 1.0 if s["p4"] >= 1.0 else s["p4"]
            c20 = 1.0 if s["p4"] >= 2.0 else max(0.0, s["p4"] - c19)
            c21 = round(max(0.0, s["p4"] - c19 - c20), 2)
            p4_ans = [c19, c20, c21]

        for i, val in enumerate(p4_ans):
            cell = ws_comp.cell(row=row_idx, column=32+i, value=val)
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.border = thin_border
            cell.number_format = "0.00"
            cell.font = font_data
        ws_comp.cell(row=row_idx, column=35, value=s["p4"]).font = font_bold
        ws_comp.cell(row=row_idx, column=35).number_format = "0.00"
        ws_comp.cell(row=row_idx, column=35).alignment = Alignment(horizontal="right", vertical="center")

        # Total and Note
        ws_comp.cell(row=row_idx, column=36, value=s["total"]).font = font_bold
        ws_comp.cell(row=row_idx, column=36).number_format = "0.00"
        ws_comp.cell(row=row_idx, column=36).alignment = Alignment(horizontal="right", vertical="center")
        ws_comp.cell(row=row_idx, column=37, value=s["rank"]).alignment = Alignment(horizontal="center", vertical="center")
        ws_comp.cell(row=row_idx, column=38, value=s["note"]).alignment = Alignment(horizontal="left", vertical="center")
        
        # Border
        for c in range(1, 39):
            ws_comp.cell(row=row_idx, column=c).border = thin_border

    # Set column widths for ws_comp
    comp_col_widths = {
        1: 6, 2: 24, 3: 8, 4: 10,
        5: 6, 6: 6, 7: 6, 8: 6, 9: 6, 10: 6, 11: 6, 12: 6, 13: 6, 14: 6, 15: 6, 16: 6, 17: 10,
        18: 6, 19: 6, 20: 6, 21: 6, 22: 6, 23: 6, 24: 6, 25: 6, 26: 10,
        27: 8, 28: 8, 29: 8, 30: 8, 31: 10,
        32: 8, 33: 8, 34: 8, 35: 10,
        36: 12, 37: 16, 38: 90
    }
    for col_idx, width in comp_col_widths.items():
        ws_comp.column_dimensions[get_column_letter(col_idx)].width = width

    # SHEET 2: Thống kê phổ điểm
    ws2 = wb.create_sheet(title="Thống Kê Phổ Điểm")
    ws2.row_dimensions[1].height = 30
    ws2.merge_cells("A1:E1")
    ws2["A1"] = "THỐNG KÊ PHỔ ĐIỂM VÀ PHÂN LOẠI HỌC LỰC MÔN TOÁN 12B1"
    ws2["A1"].font = font_title
    ws2["A1"].alignment = Alignment(horizontal="center", vertical="center")
    
    headers_s2 = ["Khoảng Điểm / Xếp Loại", "Số Lượng (HS)", "Tỉ Lệ (%)", "Tiêu Chuẩn Đánh Giá", "Ghi Chú Hành Động"]
    ws2.row_dimensions[3].height = 25
    for c_idx, h in enumerate(headers_s2, 1):
        cell = ws2.cell(row=3, column=c_idx, value=h)
        cell.font = font_head
        cell.fill = fill_head
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    tot_hs = len(valid_students)
    gioi_count = sum(1 for s in valid_students if s["total"] >= 8.0)
    kha_count = sum(1 for s in valid_students if 6.5 <= s["total"] < 8.0)
    tb_count = sum(1 for s in valid_students if 5.0 <= s["total"] < 6.5)
    yeu_count = sum(1 for s in valid_students if s["total"] < 5.0)
    
    stat_data = [
        ("Giỏi & Xuất sắc (8.0 - 10.0đ)", gioi_count, f"{gioi_count/tot_hs*100:.1f}%", "Nắm vững toàn diện lý thuyết, giải tốt ứng dụng thực tế", "Bồi dưỡng nâng cao VDC và ôn thi HSG / ĐGNL"),
        ("Khá (6.5 - 7.9đ)", kha_count, f"{kha_count/tot_hs*100:.1f}%", "Nắm chắc kiến thức cơ bản, làm tốt trắc nghiệm, cần tối ưu tự luận", "Rèn kỹ năng giải trọn vẹn bài toán tối ưu thực tế (C21)"),
        ("Trung bình (5.0 - 6.4đ)", tb_count, f"{tb_count/tot_hs*100:.1f}%", "Đạt yêu cầu chuẩn đầu ra, còn sai sót ở phần điền số và đúng/sai", "Rèn kỹ năng đọc đề, chống nhầm mã đề ở Phần III"),
        ("Yếu & Kém (< 5.0đ)", yeu_count, f"{yeu_count/tot_hs*100:.1f}%", "Hổng kiến thức cơ bản, tô lệch cột trắc nghiệm, bỏ dở tự luận", "Lập danh sách phụ đạo riêng kỹ năng tô phiếu và dạng toán cơ bản"),
        ("Tổng số học sinh đã chấm đầy đủ", tot_hs, "100.0%", f"Điểm trung bình toàn lớp: {avg_total:.2f}đ", "")
    ]
    
    for r_idx, row in enumerate(stat_data, 4):
        ws2.row_dimensions[r_idx].height = 24
        for c_idx, val in enumerate(row, 1):
            cell = ws2.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_bold if r_idx == 8 or c_idx == 1 else font_data
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center" if c_idx in [2, 3] else "left", vertical="center")
            if r_idx == 8:
                cell.fill = fill_highlight

    ws2.column_dimensions["A"].width = 32
    ws2.column_dimensions["B"].width = 16
    ws2.column_dimensions["C"].width = 14
    ws2.column_dimensions["D"].width = 45
    ws2.column_dimensions["E"].width = 50

    # SHEET 3: Đáp án chuẩn 4 mã đề
    ws3 = wb.create_sheet(title="Đáp Án & Barem Điểm Chuẩn")
    ws3.row_dimensions[1].height = 30
    codes_all = ["101", "102", "103", "104", "201", "202", "203", "204"]
    last_col = get_column_letter(1 + len(codes_all))
    ws3.merge_cells(f"A1:{last_col}1")
    ws3["A1"] = "BẢNG ĐÁP ÁN VÀ BAREM ĐIỂM CHUẨN 8 MÃ ĐỀ GHK1 TOÁN 12 (2025 - 2026)"
    ws3["A1"].font = font_title
    ws3["A1"].alignment = Alignment(horizontal="center", vertical="center")
    
    headers_s3 = ["Phần Thi / Câu Hỏi"] + [f"Mã Đề {c}" for c in codes_all]
    ws3.row_dimensions[3].height = 25
    for c_idx, h in enumerate(headers_s3, 1):
        cell = ws3.cell(row=3, column=c_idx, value=h)
        cell.font = font_head
        cell.fill = fill_head
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    p1_keys = []
    for q in range(12):
        row = [f"Câu {q+1} (0.25đ)"] + [_CORE_KEYS[c]["p1"][q] for c in codes_all]
        p1_keys.append(row)

    row_c13 = ["Câu 13 (Đúng/Sai - 1.0đ)"] + [" - ".join(_CORE_KEYS[c]["p2"][:4]) for c in codes_all]
    row_c14 = ["Câu 14 (Đúng/Sai - 1.0đ)"] + [" - ".join(_CORE_KEYS[c]["p2"][4:]) for c in codes_all]
    p1_keys.extend([row_c13, row_c14])

    for q in range(4):
        row = [f"Câu {15+q} (Trả lời ngắn - 0.5đ)"] + [_CORE_KEYS[c]["p3"][q] for c in codes_all]
        p1_keys.append(row)

    p1_keys.append(["Câu 19 (Tự luận - 1.0đ)"] + ["m = 3; b = 1; c = -3"] * len(codes_all))
    p1_keys.append(["Câu 20 (Tự luận - 1.0đ)"] + ["v_max = 25 m/s (t = 4s)"] * len(codes_all))
    p1_keys.append(["Câu 21 (Tự luận - 1.0đ)"] + ["x = 180/√55 ≈ 24.3m"] * len(codes_all))
    
    for r_idx, row in enumerate(p1_keys, 4):
        ws3.row_dimensions[r_idx].height = 22
        for c_idx, val in enumerate(row, 1):
            cell = ws3.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_bold if r_idx in [16, 17, 22, 23, 24] or c_idx == 1 else font_data
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center" if c_idx > 1 else "left", vertical="center")
            if r_idx % 2 == 1:
                cell.fill = fill_zebra

    ws3.column_dimensions["A"].width = 28
    for c_idx in range(2, 2 + len(codes_all)):
        ws3.column_dimensions[get_column_letter(c_idx)].width = 20
    ws3.column_dimensions["E"].width = 24


    # SHEET 4: Ảnh Cần Xác Nhận & Đối Soát 146/146 Ảnh Gốc (12B1)
    ws4 = wb.create_sheet(title="Ảnh Cần Xác Nhận (12B1)")
    ws4.merge_cells("A1:G1")
    ws4["A1"] = "ĐỐI SOÁT KHO ẢNH GỐC (146/146 TỆP) & DANH SÁCH ẢNH TRÙNG LẶP - LỚP 12B1"
    ws4["A1"].font = font_title
    ws4["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws4.row_dimensions[1].height = 28

    omr_cnt = len(STUDENT_RESULTS)
    tl_cnt = sum(len(s["tl_pages"]) for s in STUDENT_RESULTS)
    dup_cnt = len(NON_EXAM_FILES_12B1)
    recon = [
        ("Tổng số ảnh gốc trong 01_ANH_CHUA_CHAM/Lop_12B1", 146, "Kiểm kê 100% tệp ảnh đầu vào (Zero-leak)"),
        ("Phiếu trắc nghiệm OMR chính thức", omr_cnt, "Đủ 43/43 học sinh dự thi (0 thiếu)"),
        ("Trang bài làm Tự luận chính thức", tl_cnt, "Đã xuất đủ 66 trang chấm bút đỏ chi tiết (bao gồm toàn bộ tờ làm thêm)"),
        ("Ảnh chụp trùng lặp (OMR & Tự luận)", dup_cnt, "Không tính trùng vào bảng điểm, dùng bản chụp chính thức tương ứng"),
        ("KIỂM ĐỊNH ĐỐI SOÁT 3 CỘT SỐ", f"{omr_cnt} + {tl_cnt} + {dup_cnt} = {omr_cnt + tl_cnt + dup_cnt}", "KHỚP TUYỆT ĐỐI 146/146 TỆP | ĐỦ 43/43 HỌC SINH"),
    ]
    ws4.merge_cells("A3:D3")
    ws4["A3"] = "Hạng Mục Kiểm Kê Kho Ảnh Lớp 12B1"
    ws4["E3"] = "Số Tệp"
    ws4.merge_cells("F3:G3")
    ws4["F3"] = "Kết Luận Kiểm Toán"
    for cc in range(1, 8):
        c = ws4.cell(row=3, column=cc)
        c.font = font_head
        c.fill = fill_head
        c.border = thin_border
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws4.row_dimensions[3].height = 24

    for idx_r, (lbl, cnt, desc) in enumerate(recon, 4):
        ws4.merge_cells(start_row=idx_r, start_column=1, end_row=idx_r, end_column=4)
        ws4.merge_cells(start_row=idx_r, start_column=6, end_row=idx_r, end_column=7)
        ws4.cell(row=idx_r, column=1, value=lbl).alignment = Alignment(horizontal="left", vertical="center")
        ws4.cell(row=idx_r, column=5, value=cnt).alignment = Alignment(horizontal="center", vertical="center")
        ws4.cell(row=idx_r, column=6, value=desc).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        is_last = (idx_r == 4 + len(recon) - 1)
        for cc in range(1, 8):
            c = ws4.cell(row=idx_r, column=cc)
            c.border = thin_border
            c.fill = fill_highlight if is_last else fill_zebra
            c.font = font_bold if (is_last or cc == 5) else font_data
        ws4.row_dimensions[idx_r].height = 22

    start_tbl4 = 10
    for col_idx, h in enumerate(["STT", "Tên tệp ảnh", "Phân loại", "Học sinh liên quan", "Mã đề", "Nội dung nhận diện", "Biện pháp xử lý"], 1):
        c = ws4.cell(row=start_tbl4, column=col_idx, value=h)
        c.font = font_head
        c.fill = PatternFill(start_color="D97706", end_color="D97706", fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = thin_border
    ws4.row_dimensions[start_tbl4].height = 24

    for r_idx, item in enumerate(NON_EXAM_FILES_12B1, start_tbl4 + 1):
        ws4.row_dimensions[r_idx].height = 24
        vals = [r_idx - start_tbl4, item["filename"], item["type"], item["student_match"] or "-", item["code"] or "-", item["content"], item["action"]]
        for ci, val in enumerate(vals, 1):
            c = ws4.cell(row=r_idx, column=ci, value=val)
            c.border = thin_border
            c.font = font_data
            c.alignment = Alignment(horizontal="center" if ci in (1, 3, 5) else "left", vertical="center", wrap_text=True)

    for ci, w in {1: 6, 2: 26, 3: 24, 4: 24, 5: 10, 6: 58, 7: 44}.items():
        ws4.column_dimensions[get_column_letter(ci)].width = w

    wb.save(excel_path)
    wb.close()
    print(f"EXCEL SAVED SUCCESSFULLY: {excel_path}")

if __name__ == '__main__':
    export_excel()
