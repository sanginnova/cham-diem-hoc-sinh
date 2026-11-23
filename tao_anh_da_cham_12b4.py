# -*- coding: utf-8 -*-
"""
HỆ THỐNG XUẤT BẢN BỘ ẢNH BÀI LÀM ĐÃ CHẤM ĐỎ TRỰC QUAN - LỚP 12B4
Giáo viên: ThS. Nguyễn Văn Sang - Khoa Cơ bản - Trường Cao đẳng Nghề số 1 - BQP
Tái sử dụng engine Bilinear 4 góc, Bảng Đối Soát Sư Phạm lề phải, và Tự luận phân bước chống chồng lấn
từ tao_anh_da_cham_12b6.py; tự động nhận diện chiều xoay CW/CCW cho ảnh chụp ngang.
"""

import glob
import os
import sys
import tempfile
import cv2
import data_12b4
import tao_anh_da_cham_12b6 as engine

# Ensure engine uses data_12b4's MASTER_KEYS (same 101-104)
engine.MASTER_KEYS = data_12b4.MASTER_KEYS


import numpy as np


def ensure_upright_tl_file(in_path, tmp_dir):
    img = engine.imread_unicode(in_path)
    if img is None:
        return in_path
    h, w = img.shape[:2]
    if h >= w:
        return in_path
    # Check red ink (#) in left 25% vs right 25%
    hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
    m_red = cv2.bitwise_or(
        cv2.inRange(hsv, np.array([0, 90, 70]), np.array([12, 255, 255])),
        cv2.inRange(hsv, np.array([165, 90, 70]), np.array([180, 255, 255])),
    )
    red_left = cv2.countNonZero(m_red[:, :int(0.25 * w)])
    red_right = cv2.countNonZero(m_red[:, int(0.75 * w):])
    if red_left > red_right and red_left > 80:
        cand = cv2.rotate(img, cv2.ROTATE_90_CLOCKWISE)
        tmp_path = os.path.join(tmp_dir, "tl_" + os.path.basename(in_path))
        engine.imwrite_unicode(tmp_path, cand)
        return tmp_path
    if red_right > red_left and red_right > 80:
        return in_path  # default CCW in engine
    # Fallback: check mean brightness of left 10% vs right 10% (top margin is brighter/cleaner)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    if np.mean(gray[:, :int(0.10 * w)]) > np.mean(gray[:, int(0.90 * w):]):
        cand = cv2.rotate(img, cv2.ROTATE_90_CLOCKWISE)
        tmp_path = os.path.join(tmp_dir, "tl_" + os.path.basename(in_path))
        engine.imwrite_unicode(tmp_path, cand)
        return tmp_path
    return in_path


def ensure_upright_omr_file(in_path, tmp_dir):
    img = engine.imread_unicode(in_path)
    if img is None:
        return in_path
    h, w = img.shape[:2]
    if h >= w:
        return in_path
    # Test both 90 CCW (default in engine) and 90 CW
    for rot_code in (cv2.ROTATE_90_COUNTERCLOCKWISE, cv2.ROTATE_90_CLOCKWISE):
        cand = cv2.rotate(img, rot_code)
        grid = engine.extract_6_grid_rows(cand)
        if grid is not None:
            t1, b1, t2, b2, t3, b3 = grid
            ch, cw = cand.shape[:2]
            if t1[0][0] >= 0.25 * cw and 0.32 * ch <= t1[0][1] <= 0.48 * ch and b3[0][1] >= 0.83 * ch:
                if rot_code == cv2.ROTATE_90_COUNTERCLOCKWISE:
                    return in_path  # engine's default handles CCW directly
                tmp_path = os.path.join(tmp_dir, os.path.basename(in_path))
                engine.imwrite_unicode(tmp_path, cand)
                return tmp_path
    return in_path


def run_batch():
    base_in = r"D:\TOAN\05_CONG_CU_VA_ENGINE\Cham_Diem_Hoc_Sinh\01_ANH_CHUA_CHAM\Lop_12B4"
    base_out = r"D:\TOAN\05_CONG_CU_VA_ENGINE\Cham_Diem_Hoc_Sinh\02_ANH_DA_CHAM\Lop_12B4"
    os.makedirs(base_out, exist_ok=True)

    for f in glob.glob(os.path.join(base_out, "*.*")):
        os.remove(f)

    students = data_12b4.STUDENTS_12B4
    print(f"Bắt đầu xuất bản bộ ảnh chấm bút đỏ Lớp 12B4 ({len(students)} học sinh dự thi)...")

    omr_ok = 0
    tl_ok = 0
    with tempfile.TemporaryDirectory() as tmp_dir:
        for s in students:
            stt = s["stt"]
            name_clean = engine.sanitize_filename(s["name"])
            code = s["code"]

            raw_omr = os.path.join(base_in, s["file_omr"])
            in_omr = ensure_upright_omr_file(raw_omr, tmp_dir)
            out_omr = os.path.join(base_out, f"12B4_{stt:02d}_{name_clean}_{code}_DaCham.jpg")
            assert engine.annotate_omr_sheet(in_omr, s, out_omr), f"OMR failed for 12B4 STT {stt}"
            omr_ok += 1

            tl_pages = s["tl_pages"]
            total_tl_pages = len(tl_pages)
            present_qs = {q for pg in tl_pages for q in pg.get("questions", [])}
            missing_qs = [q for q in ("C19", "C20", "C21") if q not in present_qs]

            for p_idx, pg in enumerate(tl_pages, start=1):
                raw_tl = os.path.join(base_in, pg["file"])
                in_tl = ensure_upright_tl_file(raw_tl, tmp_dir)
                if total_tl_pages == 1:
                    out_tl = os.path.join(base_out, f"12B4_{stt:02d}_{name_clean}_{code}_TuLuan_DaCham.jpg")
                else:
                    out_tl = os.path.join(base_out, f"12B4_{stt:02d}_{name_clean}_{code}_TuLuan_Trang{p_idx}_DaCham.jpg")
                m_final = missing_qs if (p_idx == total_tl_pages) else ()
                assert engine.annotate_essay_sheet_pedagogical(
                    in_tl, s, out_tl, page_info=pg, page_idx=p_idx, total_pages=total_tl_pages, missing_on_final=m_final
                ), f"TL failed for 12B4 STT {stt} page {p_idx}"
                tl_ok += 1

    print(f"HOÀN THÀNH: Đã xuất bản {omr_ok} ảnh OMR và {tl_ok} ảnh Tự luận -> {base_out}")

    print("\n>>> TỰ ĐỘNG CHẠY HỆ THỐNG DOUBLE CHECK 2 VÒNG ĐỘC LẬP...")
    from double_check_cham_diem import run_double_check
    if not run_double_check("Lop_12B4", expected_count=36):
        sys.exit(1)


if __name__ == "__main__":
    run_batch()
