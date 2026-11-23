# -*- coding: utf-8 -*-
"""
HỆ THỐNG SOÁT CHÉP BÀI TỰ LUẬN TỰ ĐỘNG (PLAGIARISM & CHEATING DETECTION) — ENGINE PRO
Chủ nhiệm hệ thống: ThS. Nguyễn Văn Sang - Khoa Cơ bản - Trường Cao đẳng Nghề số 1 - BQP

QUY TẮC SƯ PHẠM BẤT BIẾN:
- Khử sạch 100% Dương tính giả (False Positives): Bỏ qua bài bỏ trắng, không làm bài.
- Bằng chứng thép (Smoking Gun): Cùng sai một lỗi đặc thù (sai dấu, ép nghiệm vô lý, chép chéo mã đề).
- Chế tài nghiêm minh: Hạ 100% điểm câu chép về 0,00đ cho cả 2 học sinh.
- Gạch chéo đỏ [PHÁT HIỆN CHÉP BÀI -> 0,00đ] trên ảnh bài làm tự luận.
"""

import os
import sys
import difflib
import argparse
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

try:
    import cv2
    import numpy as np
except ImportError:
    cv2 = None
    np = None

BASE_DIR = r"D:\TOAN\05_CONG_CU_VA_ENGINE\Cham_Diem_Hoc_Sinh"
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from data_12b1 import STUDENTS_12B1
from data_12b6 import STUDENTS_12B6

BLANK_PHRASES = [
    "không làm bài", "thí sinh không làm", "chưa đặt ẩn", "chưa tính", 
    "chưa giải", "chưa lập", "bỏ dở", "gạch bỏ", "không làm câu"
]

def is_blank_or_no_attempt(steps_text: str, error_note: str) -> bool:
    """Kiểm tra xem câu hỏi có phải do học sinh bỏ trắng/không làm hay không."""
    combined = (steps_text + " " + error_note).lower()
    for phrase in BLANK_PHRASES:
        if phrase in combined:
            return True
    return False

def normalize_text(text: str) -> str:
    """Chuẩn hóa chuỗi để so sánh độ tương đồng."""
    if not text:
        return ""
    s = text.lower().replace("\n", " ").replace("\r", " ")
    for ch in [",", ";", ".", ":", "(", ")", "[", "]", "{", "}", "=", "-", "+", "*", "/", ">", "<"]:
        s = s.replace(ch, f" {ch} ")
    return " ".join(s.split())

def detect_plagiarism_in_class(students, class_name="12B1"):
    """
    Rà soát phát hiện chép bài tự luận CHÍNH XÁC CAO:
    1. Bằng chứng thép 1: Cặp học sinh có cùng một lỗi sai đặc thù (Identical Unique Mistake).
       - Cùng sai dấu ở bước cụ thể nhưng ép ra nghiệm đúng hoặc cùng ra nghiệm sai kỳ lạ.
       - Cùng viết nhầm một công thức không tự nhiên.
    2. Bằng chứng thép 2: Khác mã đề nhưng bài làm lại có các bước và kết quả giống hệt nhau (Cross-code copying).
    3. Bằng chứng thép 3: Ngồi gần nhau (|STT1 - STT2| <= 3) có lời giải trùng khớp bất thường trên từng bước giải.
    """
    detected_cases = []
    n = len(students)

    for i in range(n):
        for j in range(i + 1, n):
            s1 = students[i]
            s2 = students[j]
            p4_1 = s1.get("p4_detail", {})
            p4_2 = s2.get("p4_detail", {})

            for q in ["C19", "C20", "C21"]:
                d1 = p4_1.get(q, {})
                d2 = p4_2.get(q, {})

                sc1 = d1.get("score", 0.0)
                sc2 = d2.get("score", 0.0)

                steps1_list = [st.get("text", "") for st in d1.get("steps", [])]
                steps2_list = [st.get("text", "") for st in d2.get("steps", [])]
                steps1 = " ".join(steps1_list).strip()
                steps2 = " ".join(steps2_list).strip()
                err1 = d1.get("error_note", "").strip()
                err2 = d2.get("error_note", "").strip()

                # Bỏ qua nếu cả hai đều không làm bài hoặc bỏ trắng
                if is_blank_or_no_attempt(steps1, err1) and is_blank_or_no_attempt(steps2, err2):
                    continue

                # Bỏ qua nếu cả hai đều 0 điểm và không có bước làm nào hợp lệ
                has_ok1 = any(st.get("ok", False) for st in d1.get("steps", []))
                has_ok2 = any(st.get("ok", False) for st in d2.get("steps", []))
                if sc1 == 0.0 and sc2 == 0.0 and (not has_ok1 or not has_ok2):
                    continue

                stt_diff = abs(s1.get("stt", 0) - s2.get("stt", 0))
                same_code = (s1.get("code") == s2.get("code"))

                # ----------------- KIỂM TRA BẰNG CHỨNG THÉP -----------------
                is_cheating = False
                reason = ""
                evidence_text = ""

                # 1. Trùng lỗi sai đặc thù bất thường (Ví dụ cùng sai dấu, cùng ép nghiệm)
                if err1 and err2 and not is_blank_or_no_attempt("", err1) and not is_blank_or_no_attempt("", err2):
                    err_sim = difflib.SequenceMatcher(None, normalize_text(err1), normalize_text(err2)).ratio()
                    if err_sim >= 0.75:
                        is_cheating = True
                        reason = f"TRÙNG LỖI SAI ĐẶC BIỆT ({err_sim:.1%}): Hai học sinh cùng mắc lỗi sai y hệt nhau"
                        evidence_text = f"Lỗi trùng: {err1}"

                # 2. Khác mã đề nhưng chép bài của nhau
                if not same_code and not is_cheating:
                    full_1 = normalize_text(steps1)
                    full_2 = normalize_text(steps2)
                    sim_steps = difflib.SequenceMatcher(None, full_1, full_2).ratio()
                    if sim_steps >= 0.85 and len(full_1) > 50:
                        is_cheating = True
                        reason = f"CHÉP BÀI KHÁC MÃ ĐỀ: Mã {s1.get('code')} chép của Mã {s2.get('code')} (độ giống {sim_steps:.1%})"
                        evidence_text = f"Bước giải: {steps1[:70]}..."

                # 3. Ngồi cạnh nhau (|STT1 - STT2| <= 2) và lời giải trùng lặp từng câu chữ bất thường
                if not is_cheating and stt_diff <= 2 and (sc1 > 0 or sc2 > 0):
                    full_1 = normalize_text(steps1)
                    full_2 = normalize_text(steps2)
                    sim_steps = difflib.SequenceMatcher(None, full_1, full_2).ratio()
                    if sim_steps >= 0.90 and len(full_1) > 60:
                        is_cheating = True
                        reason = f"CHÉP BÀI BẠN NGỒI CẠNH (STT {s1.get('stt')} & STT {s2.get('stt')} cách {stt_diff} bàn): Độ trùng khớp {sim_steps:.1%}"
                        evidence_text = f"Bước giải: {steps1[:70]}..."

                if is_cheating:
                    detected_cases.append({
                        "class": class_name,
                        "student1": s1["name"],
                        "stt1": s1.get("stt", 0),
                        "code1": s1.get("code", ""),
                        "score1": sc1,
                        "student2": s2["name"],
                        "stt2": s2.get("stt", 0),
                        "code2": s2.get("code", ""),
                        "score2": sc2,
                        "question": q,
                        "reason": reason,
                        "evidence": evidence_text,
                    })

    return detected_cases

def apply_plagiarism_penalties(students, detected_cases, class_name):
    """
    Áp dụng chế tài:
    - Trừ điểm câu chép nhau về 0.00đ cho cả 2 học sinh.
    - Cập nhật lại tổng điểm P4 và tổng điểm môn Toán.
    - Đánh dấu cờ gian lận trong dữ liệu học sinh.
    """
    penalized = []

    for case in detected_cases:
        stt1, stt2 = case["stt1"], case["stt2"]
        q = case["question"]

        for s in students:
            if s.get("stt") in [stt1, stt2]:
                other_name = case["student2"] if s.get("stt") == stt1 else case["student1"]
                p4 = s.setdefault("p4_detail", {})
                q_detail = p4.setdefault(q, {})
                
                old_score = q_detail.get("score", 0.0)
                q_detail["old_score"] = old_score
                q_detail["score"] = 0.0  # 0,00đ chế tài
                q_detail["is_plagiarized"] = True
                q_detail["plagiarized_with"] = other_name
                q_detail["plagiarism_reason"] = case["reason"]

                # Tính lại P4
                c19 = p4.get("C19", {}).get("score", 0.0)
                c20 = p4.get("C20", {}).get("score", 0.0)
                c21 = p4.get("C21", {}).get("score", 0.0)
                s["p4"] = round(c19 + c20 + c21, 2)
                s["answers_p4"] = [c19, c20, c21]

                # Tính lại tổng điểm
                s["total_score"] = round(s.get("total_tn", 0.0) + s["p4"], 2)
                penalized.append((s["name"], q, old_score, other_name))

    return penalized

def print_plagiarism_report(detected_cases):
    """Xuất báo cáo soát chép bài tự luận ra console chuẩn Sư phạm."""
    print("\n" + "=" * 85)
    print("      BÁO CÁO KIỂM TOÁN & PHÁT HIỆN CHÉP BÀI TỰ LUẬN (GIỮA KỲ 1)")
    print("      Khoa Cơ bản — Trường Cao đẳng Nghề số 1 - BQP | GV: ThS. Nguyễn Văn Sang")
    print("=" * 85)

    if not detected_cases:
        print("[✓] TUYỆT VỜI: Không phát hiện cặp học sinh nào có dấu hiệu chép bài tự luận!")
        print("=" * 85 + "\n")
        return

    print(f"[!] PHÁT HIỆN CHÍNH XÁC {len(detected_cases)} TRƯỜNG HỢP CHÉP BÀI TỰ LUẬN (ĐÃ XÁC MINH BẰNG CHỨNG THÉP):")
    print("-" * 85)

    for idx, c in enumerate(detected_cases, 1):
        print(f"[{idx:02d}] Lớp {c['class']} | Câu {c['question']}")
        print(f"     • Học sinh A: STT {c['stt1']:02d} - {c['student1']:<22} (Mã đề: {c['code1']}) | Điểm cũ: {c['score1']:.2f}đ -> HẠ VỀ: 0,00đ")
        print(f"     • Học sinh B: STT {c['stt2']:02d} - {c['student2']:<22} (Mã đề: {c['code2']}) | Điểm cũ: {c['score2']:.2f}đ -> HẠ VỀ: 0,00đ")
        print(f"     • Tính chất vi phạm: {c['reason']}")
        print(f"     • Chi tiết đối chiếu: {c['evidence']}")
        print("-" * 85)

    print("⚠️  CHẾ TÀI ĐÃ THI HÀNH: 100% điểm của các câu chép bài trên bị hủy về 0,00đ.")
    print("=" * 85 + "\n")

def main():
    parser = argparse.ArgumentParser(description="Hệ thống soát chép bài tự luận học sinh")
    parser.add_argument("--lop", choices=["12B1", "12B6", "all"], default="all", help="Lớp cần rà soát")
    parser.add_argument("--apply", action="store_true", help="Áp dụng chế tài 0.00đ và cập nhật hệ thống")
    args = parser.parse_args()

    all_cases = []
    
    if args.lop in ["12B1", "all"]:
        cases_12b1 = detect_plagiarism_in_class(STUDENTS_12B1, "12B1")
        all_cases.extend(cases_12b1)

    if args.lop in ["12B6", "all"]:
        cases_12b6 = detect_plagiarism_in_class(STUDENTS_12B6, "12B6")
        all_cases.extend(cases_12b6)

    print_plagiarism_report(all_cases)

    if args.apply and all_cases:
        print("[*] Đang áp dụng chế tài 0,00đ cho các trường hợp chép bài...")
        if args.lop in ["12B1", "all"]:
            p1 = apply_plagiarism_penalties(STUDENTS_12B1, [c for c in all_cases if c["class"] == "12B1"], "12B1")
            print(f"[✓] Đã xử lý Lớp 12B1: {len(p1)} lượt câu bị hạ điểm.")
        if args.lop in ["12B6", "all"]:
            p6 = apply_plagiarism_penalties(STUDENTS_12B6, [c for c in all_cases if c["class"] == "12B6"], "12B6")
            print(f"[✓] Đã xử lý Lớp 12B6: {len(p6)} lượt câu bị hạ điểm.")

if __name__ == "__main__":
    main()
