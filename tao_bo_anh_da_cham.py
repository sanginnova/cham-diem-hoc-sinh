# -*- coding: utf-8 -*-
"""
HỆ THỐNG XUẤT BẢN BỘ ẢNH BÀI LÀM ĐÃ CHẤM ĐỎ TRỰC QUAN - LỚP 12B1
Giáo viên: ThS. Nguyễn Văn Sang - Khoa Cơ bản - Trường Cao đẳng Nghề số 1 - BQP
Tái sử dụng 100% engine Bilinear 4 góc, Bảng Đối Soát Sư Phạm lề phải, và Tự luận phân bước chống chồng lấn
từ tao_anh_da_cham_12b6.py trên tập dữ liệu kiểm toán 146/146 ảnh của data_12b1.py.
"""

import glob
import os
import sys
import data_12b1
from tao_anh_da_cham_12b6 import (
    annotate_omr_sheet,
    annotate_essay_sheet_pedagogical,
    sanitize_filename,
)


def run_batch():
    base_in = r"D:\TOAN\05_CONG_CU_VA_ENGINE\Cham_Diem_Hoc_Sinh\01_ANH_CHUA_CHAM\Lop_12B1"
    base_out = r"D:\TOAN\05_CONG_CU_VA_ENGINE\Cham_Diem_Hoc_Sinh\02_ANH_DA_CHAM\Lop_12B1"
    os.makedirs(base_out, exist_ok=True)

    for f in glob.glob(os.path.join(base_out, "*.*")):
        os.remove(f)

    students = data_12b1.STUDENTS_12B1
    print(f"Bắt đầu xuất bản bộ ảnh chấm bút đỏ Lớp 12B1 ({len(students)} học sinh)...")

    omr_ok = 0
    tl_ok = 0
    for s in students:
        stt = s["stt"]
        name_clean = sanitize_filename(s["name"])
        code = s["code"]

        in_omr = os.path.join(base_in, s["file_omr"])
        out_omr = os.path.join(base_out, f"{stt:02d}_{name_clean}_{code}_TN_DaCham.jpg")
        assert annotate_omr_sheet(in_omr, s, out_omr), f"OMR failed for 12B1 STT {stt}"
        omr_ok += 1

        tl_pages = s["tl_pages"]
        total_tl_pages = len(tl_pages)
        present_qs = {q for pg in tl_pages for q in pg.get("questions", [])}
        missing_qs = [q for q in ("C19", "C20", "C21") if q not in present_qs]

        for p_idx, pg in enumerate(tl_pages, start=1):
            in_tl = os.path.join(base_in, pg["file"])
            suffix = f"_TL_Trang{p_idx}" if total_tl_pages > 1 else "_TL"
            out_tl = os.path.join(base_out, f"{stt:02d}_{name_clean}_{code}{suffix}_DaCham.jpg")
            m_final = missing_qs if (p_idx == total_tl_pages) else ()
            assert annotate_essay_sheet_pedagogical(
                in_tl, s, out_tl, page_info=pg, page_idx=p_idx, total_pages=total_tl_pages, missing_on_final=m_final
            ), f"TL failed for 12B1 STT {stt} page {p_idx}"
            tl_ok += 1

    print(f"HOÀN THÀNH: Đã xuất bản {omr_ok} ảnh OMR và {tl_ok} ảnh Tự luận -> {base_out}")

    print("\n>>> TỰ ĐỘNG CHẠY HỆ THỐNG DOUBLE CHECK 2 VÒNG ĐỘC LẬP...")
    from double_check_cham_diem import run_double_check
    if not run_double_check("Lop_12B1", expected_count=43):
        sys.exit(1)


if __name__ == "__main__":
    run_batch()
