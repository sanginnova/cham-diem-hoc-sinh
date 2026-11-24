#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HỆ THỐNG DOUBLE CHECK TỰ ĐỘNG 2 VÒNG ĐỘC LẬP (ZERO SILENT PASS)
Chủ nhiệm hệ thống: ThS. Nguyễn Văn Sang - Khoa Cơ bản - Trường Cao đẳng Nghề số 1 - BQP

- VÒNG 1 (DATA & INVENTORY GATE):
  * Kiểm định sĩ số thực tế (12B1: 43 HS, 12B6: 34 HS) và không chấp nhận ô None / lỗi kiểu.
  * Nhận diện cột động theo tiêu đề Header.
  * Tái chấm chéo độc lập từng câu (C1..C18, C19..C21) từ sheet Đối Chiếu bằng MASTER_KEYS & P2_SCALE.
  * Với 12B6: Kiểm tra cả file riêng Lớp 12B6 và đối soát 123/123 ảnh gốc (34 OMR + 60 TL + 29 Non-Exam).
- VÒNG 2 (SUB-REGION VISUAL AUDIT GATE):
  * Xác minh đủ 100% số ảnh OMR & Tự luận (12B6: 34 OMR + 60 TL; 12B1: 43 OMR + 53 TL).
  * Trừ nền ảnh gốc (sau khi xoay chuẩn chiều dọc) để đo phân vùng bút chấm độc lập:
    - Chặn vòng tròn bay lệch ra lề hành chính trái (y >= 0.20h, x < 0.24w).
    - Bắt buộc có nét bút chấm ở cả Phần I, Phần II và Phần III (loại trừ Panel lề phải).
    - Bắt buộc có khung chấm chi tiết ở thân trang Tự luận (y > 0.12h, x > 0.40w).
"""

import glob
import os
import re
import sys
import cv2
import numpy as np
import openpyxl

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

BASE_DIR = r"D:\TOAN\05_CONG_CU_VA_ENGINE\Cham_Diem_Hoc_Sinh"
EXCEL_MAIN = os.path.join(BASE_DIR, "03_KET_QUA_EXCEL", "Ket_Qua_Cham_Diem_Giua_Ky_1_Lop12.xlsx")
EXCEL_12B6 = os.path.join(BASE_DIR, "03_KET_QUA_EXCEL", "Ket_Qua_Cham_Diem_Giua_Ky_1_Lop12B6.xlsx")
EXCEL_12B4 = os.path.join(BASE_DIR, "03_KET_QUA_EXCEL", "Ket_Qua_Cham_Diem_Giua_Ky_1_Lop12B4.xlsx")

from common_utils import MASTER_KEYS_12 as MASTER_KEYS
P2_SCALE = {0: 0.0, 1: 0.10, 2: 0.25, 3: 0.50, 4: 1.00}


def _norm_p3(val):
    s = str(val if val is not None else "").strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def _extract_student_ans(cell_val, is_p3=False):
    s = str(cell_val if cell_val is not None else "").strip()
    if "→" in s:
        s = s.split("→", 1)[0].strip()
    if s in ("Trống", "-", "—"):
        return ""
    return _norm_p3(s) if is_p3 else s


def _to_float(val, label):
    if val is None or (isinstance(val, str) and (not val.strip() or val.strip().startswith("="))):
        raise ValueError(f"{label} có giá trị rỗng hoặc công thức chưa cache ({val!r})")
    return float(val)


def _find_col(headers, patterns):
    for pat in patterns:
        for idx, h in enumerate(headers):
            if h and pat.lower() in h.lower():
                return idx
    return None


def _color_pen_mask(img_bgr):
    hsv = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2HSV)
    m_green = cv2.inRange(hsv, np.array([35, 125, 70]), np.array([85, 255, 255]))
    m_amber = cv2.inRange(hsv, np.array([10, 130, 90]), np.array([30, 255, 255]))
    m_red1 = cv2.inRange(hsv, np.array([0, 130, 90]), np.array([10, 255, 255]))
    m_red2 = cv2.inRange(hsv, np.array([165, 130, 90]), np.array([180, 255, 255]))
    return cv2.bitwise_or(cv2.bitwise_or(m_green, m_amber), cv2.bitwise_or(m_red1, m_red2))


def _audit_excel_workbook(wb_path, class_name, expected_count, errors):
    if not os.path.exists(wb_path):
        errors.append(f"Không tìm thấy file Excel: {wb_path}")
        return
    wb = openpyxl.load_workbook(wb_path, data_only=True)
    if class_name == "Lop_12B6":
        score_sheet, detail_sheet = "Bảng Điểm 12B6", "Đối Chiếu 12B6"
    elif class_name == "Lop_12B4":
        score_sheet, detail_sheet = "Bảng Điểm 12B4", "Đối Chiếu 12B4"
    else:
        score_sheet, detail_sheet = "Bảng Điểm Chi Tiết", "Đối Chiếu Bài Làm Chi Tiết"

    if score_sheet not in wb.sheetnames:
        errors.append(f"Thiếu sheet '{score_sheet}' trong {os.path.basename(wb_path)}")
        return
    if detail_sheet not in wb.sheetnames:
        errors.append(f"Thiếu sheet '{detail_sheet}' trong {os.path.basename(wb_path)}")
        return

    ws = wb[score_sheet]
    hdr_row, hdrs = None, []
    for r in range(1, min(10, ws.max_row + 1)):
        vals = [str(ws.cell(r, c).value or "").strip() for c in range(1, ws.max_column + 1)]
        if vals and vals[0] == "STT" and any("Phần I" in v for v in vals):
            hdr_row, hdrs = r, vals
            break
    if hdr_row is None:
        errors.append(f"Không tìm thấy hàng tiêu đề STT trong '{score_sheet}'")
        return

    c_stt = _find_col(hdrs, ["STT"])
    c_name = _find_col(hdrs, ["Họ và tên", "Họ và Tên"])
    c_code = _find_col(hdrs, ["Mã đề", "Mã Đề"])
    c_p1 = _find_col(hdrs, ["Phần I"])
    c_p2 = _find_col(hdrs, ["Phần II"])
    c_p3 = _find_col(hdrs, ["Phần III"])
    c_tn = _find_col(hdrs, ["Tổng TN"])
    c_p4 = _find_col(hdrs, ["Phần IV", "Tự luận"])
    c_tot = _find_col(hdrs, ["Tổng Điểm", "Tổng kết"])

    for req_name, req_col in [("Mã đề", c_code), ("P1", c_p1), ("P2", c_p2), ("P3", c_p3), ("P4", c_p4), ("Tổng", c_tot)]:
        if req_col is None:
            errors.append(f"Thiếu cột '{req_name}' trong '{score_sheet}'")
            return

    score_map = {}
    for r in range(hdr_row + 1, ws.max_row + 1):
        v_stt = ws.cell(r, c_stt + 1).value
        if v_stt is not None and str(v_stt).strip().isdigit():
            stt = int(str(v_stt).strip())
            name = str(ws.cell(r, c_name + 1).value or "").strip()
            code = str(ws.cell(r, c_code + 1).value or "").strip()
            try:
                p1 = _to_float(ws.cell(r, c_p1 + 1).value, f"{score_sheet} STT {stt} P1")
                p2 = _to_float(ws.cell(r, c_p2 + 1).value, f"{score_sheet} STT {stt} P2")
                p3 = _to_float(ws.cell(r, c_p3 + 1).value, f"{score_sheet} STT {stt} P3")
                p4 = _to_float(ws.cell(r, c_p4 + 1).value, f"{score_sheet} STT {stt} P4")
                tot = _to_float(ws.cell(r, c_tot + 1).value, f"{score_sheet} STT {stt} Tổng")
                if c_tn is not None:
                    tn = _to_float(ws.cell(r, c_tn + 1).value, f"{score_sheet} STT {stt} Tổng TN")
                    if abs((p1 + p2 + p3) - tn) > 0.01:
                        errors.append(f"[{os.path.basename(wb_path)}/{score_sheet}] STT {stt} ({name}): P1+P2+P3={p1+p2+p3:.2f} != TổngTN={tn:.2f}")
                if abs((p1 + p2 + p3 + p4) - tot) > 0.01:
                    errors.append(f"[{os.path.basename(wb_path)}/{score_sheet}] STT {stt} ({name}): P1+P2+P3+P4={p1+p2+p3+p4:.2f} != Tổng={tot:.2f}")
                score_map[stt] = (name, code, p1, p2, p3, p4, tot)
            except Exception as e:
                errors.append(f"[{os.path.basename(wb_path)}/{score_sheet}] {e}")

    if len(score_map) != expected_count:
        errors.append(f"[{os.path.basename(wb_path)}/{score_sheet}] Sĩ số không khớp: {len(score_map)}/{expected_count}")

    # Tái chấm chéo từ sheet Đối Chiếu
    ws2 = wb[detail_sheet]
    d_hdr_row, d_hdrs = None, []
    for r in range(1, min(10, ws2.max_row + 1)):
        vals = [str(ws2.cell(r, c).value or "").strip() for c in range(1, ws2.max_column + 1)]
        if vals and vals[0] == "STT" and "C1" in vals and "C19" in vals:
            d_hdr_row, d_hdrs = r, vals
            break
    if d_hdr_row is None:
        errors.append(f"[{os.path.basename(wb_path)}/{detail_sheet}] Thiếu hàng tiêu đề câu hỏi (C1..C21)")
        return

    col_of = {h: i + 1 for i, h in enumerate(d_hdrs) if h}
    detail_count = 0
    for r in range(d_hdr_row + 1, ws2.max_row + 1):
        v_stt = ws2.cell(r, col_of["STT"]).value
        if v_stt is None or not str(v_stt).strip().isdigit():
            continue
        detail_count += 1
        stt = int(str(v_stt).strip())
        code = str(ws2.cell(r, col_of["Mã Đề"]).value or "").strip()
        if code not in MASTER_KEYS:
            errors.append(f"[{detail_sheet}] STT {stt}: Mã đề không hợp lệ ({code!r})")
            continue
        key = MASTER_KEYS[code]

        p1_ans = [_extract_student_ans(ws2.cell(r, col_of[f"C{q}"]).value) for q in range(1, 13)]
        p2_ans = [_extract_student_ans(ws2.cell(r, col_of[f"{qn}{sb}"]).value) for qn in (13, 14) for sb in ("a", "b", "c", "d")]
        p3_ans = [_extract_student_ans(ws2.cell(r, col_of[f"C{q}"]).value, is_p3=True) for q in range(15, 19)]

        rp1 = round(0.25 * sum(1 for i in range(12) if p1_ans[i] == key["p1"][i]), 2)
        c13_ok = sum(1 for i in range(4) if p2_ans[i] == key["p2_13"][i])
        c14_ok = sum(1 for i in range(4) if p2_ans[4 + i] == key["p2_14"][i])
        rp2 = round(P2_SCALE[c13_ok] + P2_SCALE[c14_ok], 2)
        rp3 = round(0.50 * sum(1 for i in range(4) if p3_ans[i] == _norm_p3(key["p3"][i])), 2)

        try:
            c_sp1 = col_of.get("P1 (3đ)") or col_of.get("Điểm P.I")
            c_sp2 = col_of.get("P2 (2đ)") or col_of.get("Điểm P.II")
            c_sp3 = col_of.get("P3 (2đ)") or col_of.get("Điểm P.III")
            c_sp4 = col_of.get("P4 (3đ)") or col_of.get("Điểm P.IV")
            sp1 = _to_float(ws2.cell(r, c_sp1).value, f"{detail_sheet} STT {stt} P1")
            sp2 = _to_float(ws2.cell(r, c_sp2).value, f"{detail_sheet} STT {stt} P2")
            sp3 = _to_float(ws2.cell(r, c_sp3).value, f"{detail_sheet} STT {stt} P3")
            c19 = _to_float(ws2.cell(r, col_of["C19"]).value, f"{detail_sheet} STT {stt} C19")
            c20 = _to_float(ws2.cell(r, col_of["C20"]).value, f"{detail_sheet} STT {stt} C20")
            c21 = _to_float(ws2.cell(r, col_of["C21"]).value, f"{detail_sheet} STT {stt} C21")
            sp4 = _to_float(ws2.cell(r, c_sp4).value, f"{detail_sheet} STT {stt} P4")
            stot = _to_float(ws2.cell(r, col_of["Tổng Điểm"]).value, f"{detail_sheet} STT {stt} Tổng")
        except Exception as e:
            errors.append(f"[{detail_sheet}] {e}")
            continue

        rp4 = round(c19 + c20 + c21, 2)
        rtot = round(rp1 + rp2 + rp3 + rp4, 2)
        if abs(rp1 - sp1) > 0.01 or abs(rp2 - sp2) > 0.01 or abs(rp3 - sp3) > 0.01 or abs(rp4 - sp4) > 0.01 or abs(rtot - stot) > 0.01:
            errors.append(
                f"[{detail_sheet}] STT {stt} Đề {code}: Tái chấm ({rp1}, {rp2}, {rp3}, {rp4} = {rtot}) "
                f"KHÔNG KHỚP sheet ({sp1}, {sp2}, {sp3}, {sp4} = {stot})"
            )
        if stt in score_map:
            _, b_code, bp1, bp2, bp3, bp4, btot = score_map[stt]
            if b_code != code or abs(bp1 - rp1) > 0.01 or abs(bp2 - rp2) > 0.01 or abs(bp3 - rp3) > 0.01 or abs(bp4 - rp4) > 0.01 or abs(btot - rtot) > 0.01:
                errors.append(
                    f"[{os.path.basename(wb_path)}] STT {stt}: Bảng Điểm ({b_code}, {bp1},{bp2},{bp3},{bp4}={btot}) "
                    f"lệch Đối Chiếu ({code}, {rp1},{rp2},{rp3},{rp4}={rtot})"
                )

    if detail_count != expected_count:
        errors.append(f"[{os.path.basename(wb_path)}/{detail_sheet}] Số HS không khớp: {detail_count}/{expected_count}")


def run_double_check(class_name="Lop_12B6", expected_count=34):
    is_12b6 = (class_name == "Lop_12B6")
    print("=" * 72)
    print(f"BẮT ĐẦU DOUBLE CHECK 2 VÒNG ĐỘC LẬP CHO {class_name.upper()} (Kỳ vọng: {expected_count} HS)")
    print("=" * 72)

    dir_chua_cham = os.path.join(BASE_DIR, "01_ANH_CHUA_CHAM", class_name)
    dir_da_cham = os.path.join(BASE_DIR, "02_ANH_DA_CHAM", class_name)
    if BASE_DIR not in sys.path:
        sys.path.append(BASE_DIR)

    errors_round_1 = []
    errors_round_2 = []

    # =========================================================================
    # VÒNG 1: KIỂM TOÁN EXCEL, ĐÁP ÁN & TRIPLE RECONCILIATION
    # =========================================================================
    print("\n--- [VÒNG 1/2] KIỂM TOÁN EXCEL, TÁI CHẤM ĐÁP ÁN & KHO ẢNH GỐC ---")
    _audit_excel_workbook(EXCEL_MAIN, class_name=class_name, expected_count=expected_count, errors=errors_round_1)
    if class_name == "Lop_12B6":
        _audit_excel_workbook(EXCEL_12B6, class_name="Lop_12B6", expected_count=expected_count, errors=errors_round_1)
        import data_12b6
        _students = data_12b6.STUDENTS_12B6
        _non_exam = data_12b6.NON_EXAM_FILES
        _expected_raw = 123
    elif class_name == "Lop_12B4":
        _audit_excel_workbook(EXCEL_12B4, class_name="Lop_12B4", expected_count=expected_count, errors=errors_round_1)
        import data_12b4
        _students = data_12b4.STUDENTS_12B4
        _non_exam = data_12b4.NON_EXAM_FILES
        _expected_raw = 105
    else:
        import data_12b1
        _students = data_12b1.STUDENTS_12B1
        _non_exam = data_12b1.NON_EXAM_FILES_12B1
        _expected_raw = 146

    if os.path.isdir(dir_chua_cham):
        raw_files = set(os.listdir(dir_chua_cham))
        omr_set = {s["file_omr"] for s in _students}
        tl_set = {pg["file"] for s in _students for pg in s["tl_pages"]}
        non_set = {item["filename"] for item in _non_exam}
        missing = raw_files - (omr_set | tl_set | non_set)
        extra = (omr_set | tl_set | non_set) - raw_files
        if missing or extra or len(raw_files) != _expected_raw:
            errors_round_1.append(f"Đối soát kho {class_name} thất bại: raw={len(raw_files)}, missing={missing}, extra={extra}")
        else:
            print(f"-> Đối soát kho gốc {class_name}: {len(raw_files)}/{_expected_raw} tệp ({len(omr_set)} OMR + {len(tl_set)} TL + {len(non_set)} Non-Exam) KHỚP 100%.")
    else:
        errors_round_1.append(f"Thiếu thư mục ảnh gốc: {dir_chua_cham}")

    if not errors_round_1:
        print("-> VÒNG 1: [PASS 100%] Dữ liệu Excel, tái chấm chéo đáp án & đối soát kho ảnh hợp lệ!")
    else:
        print(f"-> VÒNG 1: [FAIL] Phát hiện {len(errors_round_1)} lỗi:")
        for e in errors_round_1:
            print(f"   ❌ {e}")

    # =========================================================================
    # VÒNG 2: KIỂM TOÁN THỊ GIÁC PHÂN VÙNG ĐỘC LẬP (SUB-REGION PIXEL AUDIT)
    # =========================================================================
    print("\n--- [VÒNG 2/2] KIỂM TOÁN THỊ GIÁC PHÂN VÙNG 100% ẢNH ĐÃ CHẤM ---")
    if not os.path.isdir(dir_da_cham):
        errors_round_2.append(f"Không tồn tại thư mục ảnh đã chấm: {dir_da_cham}")
    else:
        all_out = sorted(glob.glob(os.path.join(dir_da_cham, "*_DaCham.jpg")))
        omr_files = [f for f in all_out if "TuLuan" not in os.path.basename(f) and "_TL" not in os.path.basename(f)]
        tl_files = [f for f in all_out if "TuLuan" in os.path.basename(f) or "_TL" in os.path.basename(f)]

        expected_tl = 60 if class_name == "Lop_12B6" else (67 if class_name == "Lop_12B4" else 66)
        print(f"-> Ảnh OMR đã xuất bản: {len(omr_files)}/{expected_count} | Ảnh Tự luận đã xuất bản: {len(tl_files)}/{expected_tl}")
        if len(omr_files) != expected_count:
            errors_round_2.append(f"Số lượng ảnh OMR không đúng: {len(omr_files)}/{expected_count}")
        if len(tl_files) != expected_tl:
            errors_round_2.append(f"Số lượng ảnh Tự luận không đúng: {len(tl_files)}/{expected_tl}")

        orig_omr_map = {s["stt"]: s["file_omr"] for s in _students}

        for img_p in omr_files:
            fname = os.path.basename(img_p)
            img = cv2.imread(img_p)
            if img is None:
                errors_round_2.append(f"Không đọc được ảnh OMR: {fname}")
                continue
            h, w = img.shape[:2]
            mask_d = _color_pen_mask(img)

            # Tìm STT từ tên file ("12B6_01_..." hoặc "01_Dao_...")
            m_stt = re.match(r"^(?:12B\d+_)?(\d+)_", fname)
            if not m_stt:
                errors_round_2.append(f"Tên file OMR sai quy ước không trích xuất được STT: {fname}")
                continue
            stt = int(m_stt.group(1))
            orig_fn = orig_omr_map.get(stt)
            if not orig_fn or not os.path.exists(os.path.join(dir_chua_cham, orig_fn)):
                errors_round_2.append(f"Không tìm thấy ảnh gốc OMR đối chiếu cho {fname} (STT {stt})")
                continue

            img_o = cv2.imread(os.path.join(dir_chua_cham, orig_fn))
            if img_o is None:
                errors_round_2.append(f"Không đọc được ảnh gốc OMR: {orig_fn}")
                continue
            if img_o.shape[1] > img_o.shape[0]:
                # Match orientation against annotated img (CW or CCW)
                cand_ccw = cv2.rotate(img_o, cv2.ROTATE_90_COUNTERCLOCKWISE)
                cand_cw = cv2.rotate(img_o, cv2.ROTATE_90_CLOCKWISE)
                if cand_ccw.shape[:2] != (h, w):
                    cand_ccw = cv2.resize(cand_ccw, (w, h))
                    cand_cw = cv2.resize(cand_cw, (w, h))
                d_ccw = np.mean(cv2.absdiff(img, cand_ccw))
                d_cw = np.mean(cv2.absdiff(img, cand_cw))
                img_o = cand_cw if d_cw < d_ccw else cand_ccw
            elif img_o.shape[:2] != (h, w):
                img_o = cv2.resize(img_o, (w, h))

            changed = (cv2.cvtColor(cv2.absdiff(img, img_o), cv2.COLOR_BGR2GRAY) > 25).astype(np.uint8) * 255
            diff_mask = cv2.bitwise_and(mask_d, changed)

            # 4 phân vùng kiểm toán độc lập (loại trừ banner y < 0.15h & Right Panel x > 0.77w)
            left_px = cv2.countNonZero(diff_mask[int(0.20 * h):, :int(0.24 * w)])
            p1_px = cv2.countNonZero(diff_mask[int(0.32 * h):int(0.58 * h), int(0.41 * w):int(0.77 * w)])
            p2_px = cv2.countNonZero(diff_mask[int(0.51 * h):int(0.70 * h), int(0.26 * w):int(0.50 * w)])
            p3_px = cv2.countNonZero(diff_mask[int(0.65 * h):int(0.95 * h), int(0.26 * w):int(0.76 * w)])

            if left_px >= 1200:
                errors_round_2.append(f"LỖI BAY LỀ TRÁI {fname}: {left_px} px màu lọt vào lề hành chính (Số phách/Điểm số)")
            if p1_px < 800:
                errors_round_2.append(f"THIẾU CHẤM PHẦN I {fname}: chỉ có {p1_px} px màu (< 800)")
            if p2_px < 350:
                errors_round_2.append(f"THIẾU CHẤM PHẦN II {fname}: chỉ có {p2_px} px màu (< 350)")
            if p3_px < 400:
                errors_round_2.append(f"THIẾU CHẤM PHẦN III {fname}: chỉ có {p3_px} px màu (< 400)")

        for tl_p in tl_files:
            fname = os.path.basename(tl_p)
            img = cv2.imread(tl_p)
            if img is None:
                errors_round_2.append(f"Không đọc được ảnh Tự luận: {fname}")
                continue
            h, w = img.shape[:2]
            body_px = cv2.countNonZero(_color_pen_mask(img)[int(0.12 * h):, int(0.40 * w):])
            if body_px < 1200:
                errors_round_2.append(f"THIẾU KHUNG CHẤM TỰ LUẬN {fname}: chỉ có {body_px} px màu ở thân trang (< 1200)")

    if not errors_round_2:
        print("-> VÒNG 2: [PASS 100%] Tất cả ảnh OMR & Tự luận vượt qua kiểm toán phân vùng độc lập!")
    else:
        print(f"-> VÒNG 2: [FAIL] Phát hiện {len(errors_round_2)} lỗi thị giác:")
        for e in errors_round_2:
            print(f"   ❌ {e}")

    print("\n" + "=" * 72)
    passed = (len(errors_round_1) == 0 and len(errors_round_2) == 0)
    if passed:
        print(f"KẾT LUẬN DOUBLE CHECK ({class_name}): PASS 100% — ĐỦ ĐIỀU KIỆN BÀN GIAO CHO THẦY SANG!")
    else:
        print(f"KẾT LUẬN DOUBLE CHECK ({class_name}): FAIL — CHẶN BÀN GIAO!")
    print("=" * 72)
    return passed


if __name__ == "__main__":
    cls = sys.argv[1] if len(sys.argv) > 1 else "Lop_12B6"
    cnt = int(sys.argv[2]) if len(sys.argv) > 2 else (43 if cls == "Lop_12B1" else (36 if cls == "Lop_12B4" else 34))
    ok = run_double_check(cls, cnt)
    sys.exit(0 if ok else 1)
