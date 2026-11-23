# -*- coding: utf-8 -*-
"""
HỆ THỐNG TỰ ĐỘNG DỌN DẸP & BẢO TOÀN WORKSPACE CHẤM ĐIỂM
Giáo viên: ThS. Nguyễn Văn Sang - Khoa Cơ bản - Trường Cao đẳng Nghề số 1 - BQP

Chức năng:
1. Quét và loại bỏ 100% file rác trung gian (ảnh debug, crop, cache, test script tạm).
2. Bảo toàn tuyệt đối 3 thư mục hạt nhân:
   - 01_ANH_CHUA_CHAM/
   - 02_ANH_DA_CHAM/
   - 03_KET_QUA_EXCEL/
3. Báo cáo dung lượng giải phóng và trạng thái thư mục chuẩn hóa.
"""

import os
import glob
import shutil

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Các mẫu file rác cần quét và xóa sạch
GARBAGE_PATTERNS = [
    # Ảnh tạm, ảnh crop, ảnh canvas trung gian
    "canvas_*.jpg", "canvas_*.png",
    "crop*.jpg", "crop*.png", "crop_*.jpg", "crop_*.png",
    "exact_*.jpg", "exact_*.png",
    "pdf_*.png", "pdf_*.jpg",
    "phone_*.jpg", "phone_*.png",
    "resized_*.jpg", "resized_*.png",
    "row_*.jpg", "row_*.png",
    "sample_*.jpg", "sample_*.png",
    "debug_*.jpg", "debug_*.png",
    "test_*.jpg", "test_*.png",

    # File script test tạm, scratch scripts
    "check_*.py",
    "scan_*.py",
    "find_*.py",
    "extract_*.py",
    "analyze_*.py",
    "test_thu_fix.py",
    "map_all.py",
    "ocr_de.py",

    # File tạm Office và hệ điều hành
    "~$*.xlsx", "~$*.docx", "*.tmp", "*.temp",
]

def clean_garbage():
    print("=" * 70)
    print("   HỆ THỐNG TỰ ĐỘNG DỌN DẸP RÁC WORKSPACE CHẤM ĐIỂM")
    print("=" * 70)
    
    deleted_files = 0
    freed_bytes = 0

    # 1. Quét và xóa file rác theo pattern ở thư mục gốc
    for pattern in GARBAGE_PATTERNS:
        matches = glob.glob(os.path.join(BASE_DIR, pattern))
        for fp in matches:
            try:
                sz = os.path.getsize(fp)
                os.remove(fp)
                deleted_files += 1
                freed_bytes += sz
                print(f"  [Xóa file rác] {os.path.basename(fp)} ({sz:,} bytes)")
            except Exception as e:
                print(f"  [Lỗi khi xóa] {os.path.basename(fp)}: {e}")

    # 2. Quét dọn __pycache__
    for root, dirs, files in os.walk(BASE_DIR):
        if "__pycache__" in dirs:
            pycache_dir = os.path.join(root, "__pycache__")
            try:
                shutil.rmtree(pycache_dir, ignore_errors=True)
                print(f"  [Xóa cache] {os.path.relpath(pycache_dir, BASE_DIR)}")
            except Exception as e:
                pass

    # 2b. Quét dọn file khóa mồ côi trong 03_KET_QUA_EXCEL và thư mục outputs cũ
    excel_dir = os.path.join(BASE_DIR, "03_KET_QUA_EXCEL")
    for lock_fp in glob.glob(os.path.join(excel_dir, "~$*.xlsx")):
        try:
            os.remove(lock_fp)
            deleted_files += 1
            print(f"  [Xóa lock mồ côi] {os.path.basename(lock_fp)}")
        except OSError:
            pass

    outputs_dir = os.path.join(BASE_DIR, "outputs")
    if os.path.isdir(outputs_dir):
        for old_sub in glob.glob(os.path.join(outputs_dir, "cham_*")):
            try:
                sub_sz = sum(os.path.getsize(os.path.join(r, f)) for r, _, fs in os.walk(old_sub) for f in fs)
                shutil.rmtree(old_sub)
                freed_bytes += sub_sz
                print(f"  [Xóa outputs cũ] {os.path.relpath(old_sub, BASE_DIR)}")
            except OSError:
                pass

    # 3. Quét dọn file debug tạm ở thư mục cha d:\TOAN nếu có
    parent_dir = os.path.dirname(BASE_DIR)
    for fn in ["debug_crop_made.jpg", "debug_crop_p1.jpg", "crop_all_p1.png"]:
        p_fp = os.path.join(parent_dir, fn)
        if os.path.exists(p_fp):
            try:
                sz = os.path.getsize(p_fp)
                os.remove(p_fp)
                deleted_files += 1
                freed_bytes += sz
                print(f"  [Xóa rác parent] {fn}")
            except Exception:
                pass

    print("-" * 70)
    print(f"--> ĐÃ DỌN SẠCH: {deleted_files} tệp rác | Giải phóng: {freed_bytes / (1024*1024):.2f} MB")
    print("-" * 70)

    # 4. Kiểm kê trạng thái 3 thư mục chuẩn hóa
    c1 = os.path.join(BASE_DIR, "01_ANH_CHUA_CHAM")
    c2 = os.path.join(BASE_DIR, "02_ANH_DA_CHAM")
    c3 = os.path.join(BASE_DIR, "03_KET_QUA_EXCEL")

    def count_files(d):
        if not os.path.exists(d): return 0
        total = 0
        for r, _, fs in os.walk(d):
            total += len(fs)
        return total

    print("KIỂM KÊ CẤU TRÚC KHOA HỌC HIỆN TẠI:")
    print(f"  1. [01_ANH_CHUA_CHAM]  : {count_files(c1)} tệp ảnh gốc (Đã phân lớp)")
    if os.path.exists(os.path.join(c1, "Lop_12B1")):
        print(f"     + Lop_12B1        : {len(os.listdir(os.path.join(c1, 'Lop_12B1')))} tệp")
    if os.path.exists(os.path.join(c1, "Lop_12B6")):
        print(f"     + Lop_12B6        : {len(os.listdir(os.path.join(c1, 'Lop_12B6')))} tệp")
        
    print(f"  2. [02_ANH_DA_CHAM]    : {count_files(c2)} tệp ảnh chấm bút đỏ")
    if os.path.exists(os.path.join(c2, "Lop_12B1")):
        print(f"     + Lop_12B1        : {len(os.listdir(os.path.join(c2, 'Lop_12B1')))} tệp")
    if os.path.exists(os.path.join(c2, "Lop_12B6")):
        print(f"     + Lop_12B6        : {len(os.listdir(os.path.join(c2, 'Lop_12B6')))} tệp")

    print(f"  3. [03_KET_QUA_EXCEL]  : {count_files(c3)} tệp Excel kết quả")
    if os.path.exists(c3):
        for f in os.listdir(c3):
            print(f"     + {f}")
    print("=" * 70)

if __name__ == "__main__":
    clean_garbage()
