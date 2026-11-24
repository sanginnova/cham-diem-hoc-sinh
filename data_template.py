# -*- coding: utf-8 -*-
"""
FILE MẪU CẤU HÌNH ĐÁP ÁN & DỮ LIỆU HỌC SINH (DÙNG CHO MỌI MÔN HỌC - GDPT 2018)
Hướng dẫn:
1. Sao chép file này thành data_<ten_lop>.py (ví dụ: data_12a1.py).
2. Điền bộ đáp án chuẩn của môn học / kỳ thi vào MASTER_KEYS (hỗ trợ nhiều mã đề).
3. Nhập kết quả nhận diện từ 01_ANH_CHUA_CHAM/<Lop>/ vào STUDENTS.
"""

# Bảng đáp án chuẩn theo từng mã đề (ví dụ 2 mã đề 101, 102 - có thể thêm 103, 104, 201...)
MASTER_KEYS = {
    "101": {
        "p1": ["A", "B", "C", "D", "A", "B", "C", "D", "A", "B", "C", "D"],  # 12 câu ABCD (0.25đ/câu)
        "p2_13": ["Đ", "Đ", "S", "S"],  # Câu 13 Đúng/Sai (4 ý)
        "p2_14": ["S", "Đ", "Đ", "S"],  # Câu 14 Đúng/Sai (4 ý)
        "p3": ["2", "1", "7", "18"],    # 4 câu Trả lời ngắn (0.5đ/câu)
    },
    "102": {
        "p1": ["D", "C", "B", "A", "D", "C", "B", "A", "D", "C", "B", "A"],
        "p2_13": ["S", "Đ", "Đ", "S"],
        "p2_14": ["Đ", "Đ", "S", "S"],
        "p3": ["7", "18", "1", "2"],
    },
}

# Thang điểm Phần II chuẩn Bộ GD&ĐT (0 ý = 0đ, 1 ý = 0.1đ, 2 ý = 0.25đ, 3 ý = 0.5đ, 4 ý = 1.0đ)
P2_SCALE = {0: 0.0, 1: 0.10, 2: 0.25, 3: 0.50, 4: 1.00}

# Danh sách bài làm học sinh mẫu
STUDENTS = [
    {
        "name": "Nguyễn Văn An",
        "class": "12A1",
        "code": "101",
        "omr_file": "sample_omr.jpg",
        "tl_files": ["sample_tl_1.jpg"],
        "answers_p1": ["A", "B", "C", "D", "A", "B", "C", "D", "A", "B", "C", "D"],
        "answers_p2": ["Đ", "Đ", "S", "S", "S", "Đ", "Đ", "S"],
        "answers_p3": ["2", "1", "7", "18"],
        "answers_p4": [1.0, 0.75, 0.5],  # Điểm tự luận Câu 19, 20, 21
    },
]
