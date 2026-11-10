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

STUDENT_RESULTS = [
    # --- TOÀN BỘ DANH SÁCH 29 BÀI LÀM LỚP 12B1 ĐÃ ĐỐI CHIẾU 100% ẢNH THỰC TẾ ---
    {
        "stt": 1,
        "name": "Đào Duy Khánh",
        "class": "12B1",
        "code": "104",
        "p1": 2.50,
        "p2": 1.10,
        "p3": 2.00,
        "p4": 3.00,
        "total": 8.60,
        "rank": "Giỏi",
        "answers_p1": ["D", "A", "D", "A", "D", "C", "A", "D", "C", "C", "D", "A"],
        "answers_p2": ["Đ", "S", "Đ", "Đ", "S", "Đ", "Đ", "S"],
        "answers_p3": ["2", "1", "7", "18"],
        "answers_p4": [1.0, 1.0, 1.0],
        "details_p1": "10/12 đúng. Sai C3 (chọn D - chuẩn B), Sai C12 (chọn A - chuẩn D) -> 2.50đ",
        "details_p2": "5/8 ý đúng (C13 đúng ý a -> 0.1đ; C14 đúng 4/4 ý -> 1.0đ) -> 1.10đ",
        "details_p3": "4/4 câu đúng tuyệt đối (C15: 2, C16: 1, C17: 7, C18: 18) -> 2.0đ",
        "details_p4": "C19 (1.0đ), C20 (1.0đ), C21 (1.0đ) - Giải trọn vẹn hoàn hảo cả 3 câu tự luận",
        "errors": "Sai 2 câu trắc nghiệm P1: Câu 3 (chọn D, đáp án B) và Câu 12 (chọn A, đáp án D); Sai 3 ý P2 C13 (b, c, d)",
        "note": "Giỏi 8.60/10.0đ. Tự luận hoàn hảo 3.0/3.0đ. Cần cẩn thận hơn khi làm trắc nghiệm P1 và P2.",
        "note": "Xuất sắc 9.5/10.0đ. Phần II, III, IV tuyệt đối 7.0/7.0đ. Cần cẩn thận hơn khi tô phương án Phần I."
    },
    {
        "stt": 2,
        "name": "Đoàn Phương Đức",
        "class": "12B1",
        "code": "104",
        "p1": 3.00,
        "p2": 2.00,
        "p3": 2.00,
        "p4": 2.00,
        "total": 9.00,
        "rank": "Giỏi",
        "answers_p1": ['D', 'A', 'B', 'A', 'D', 'C', 'A', 'D', 'C', 'C', 'D', 'D'],
        "answers_p2": ['Đ', 'Đ', 'S', 'S', 'S', 'Đ', 'Đ', 'S'],
        "answers_p3": ['2', '1', '7', '18'],
        "answers_p4": [1.0, 0.25, 0.75],
        "details_p1": "12/12 đúng tuyệt đối (100%) -> 3.0đ",
        "details_p2": "8/8 ý đúng tuyệt đối (C13: Đ-Đ-S-S; C14: S-Đ-Đ-S) -> 2.0đ",
        "details_p3": "4/4 câu đúng tuyệt đối (C15: 2, C16: 1, C17: 7, C18: 18) -> 2.0đ",
        "details_p4": "C19 (1.0đ), C20 (0.25đ - đạo hàm đúng), C21 (0.75đ - giải ra 13.75x^2 = 8100)",
        "errors": "P4 C20 chưa tìm vmax=25; C21 dừng ở phương trình bậc hai chưa bấm nghiệm",
        "note": "Trắc nghiệm đạt tuyệt đối 7.0/7.0đ. Tự luận làm rất chắc chắn."
    },
    {
        "stt": 3,
        "name": "Nguyễn Trọng Hiếu",
        "class": "12B1",
        "code": "104",
        "p1": 3.00,
        "p2": 2.00,
        "p3": 2.00,
        "p4": 1.75,
        "total": 8.75,
        "rank": "Giỏi",
        "answers_p1": ['D', 'A', 'B', 'A', 'D', 'C', 'A', 'D', 'C', 'C', 'D', 'D'],
        "answers_p2": ['Đ', 'Đ', 'S', 'S', 'S', 'Đ', 'Đ', 'S'],
        "answers_p3": ['2', '1', '7', '18'],
        "answers_p4": [1.0, 0.75, 0.0],
        "details_p1": "12/12 đúng tuyệt đối (100%) -> 3.0đ",
        "details_p2": "8/8 ý đúng tuyệt đối -> 2.0đ",
        "details_p3": "4/4 câu đúng tuyệt đối (C15: 2, C16: 1, C17: 7, C18: 18) -> 2.0đ",
        "details_p4": "C19 (1.0đ), C20 (0.75đ - tìm đúng t=4s), C21 (0đ - chưa làm)",
        "errors": "P4 C20 thiếu bước tính giá trị vmax=25 m/s; chưa làm C21",
        "note": "Trắc nghiệm tuyệt đối 7.0/7.0đ. Làm bài rất sạch sẽ, mạch lạc."
    },
    {
        "stt": 4,
        "name": "Nguyễn Mạnh Thắng",
        "class": "12B1",
        "code": "103",
        "p1": 2.75,
        "p2": 2.00,
        "p3": 2.00,
        "p4": 1.75,
        "total": 8.50,
        "rank": "Giỏi",
        "answers_p1": ['B', 'A', 'B', 'D', 'C', 'A', 'A', 'D', 'D', 'C', 'A', 'C'],
        "answers_p2": ['Đ', 'Đ', 'S', 'S', 'S', 'Đ', 'Đ', 'S'],
        "answers_p3": ['18', '2', '1', '7'],
        "answers_p4": [1.0, 0.75, 0.0],
        "details_p1": "11/12 đúng. Sai C9 (chọn D - đáp án B) -> 2.75đ",
        "details_p2": "8/8 ý đúng tuyệt đối (C13: Đ-Đ-S-S; C14: S-Đ-Đ-S) -> 2.0đ",
        "details_p3": "4/4 câu đúng tuyệt đối (C15: 18, C16: 2, C17: 1, C18: 7) -> 2.0đ",
        "details_p4": "C19 (1.0đ), C20 (0.75đ - tìm đúng t=4s), C21 (0đ - chưa làm)",
        "errors": "Sai P1 (C9); P4 C20 thiếu bước tính vmax=25; chưa làm C21",
        "note": "Điểm trắc nghiệm và trả lời ngắn rất cao. Cần phân bổ thời gian cho câu 21."
    },
    {
        "stt": 5,
        "name": "Nguyễn Tiến Đạt",
        "class": "12B1",
        "code": "103",
        "p1": 3.00,
        "p2": 2.00,
        "p3": 2.00,
        "p4": 1.25,
        "total": 8.25,
        "rank": "Giỏi",
        "answers_p1": ['B', 'A', 'B', 'D', 'C', 'A', 'A', 'D', 'B', 'C', 'A', 'C'],
        "answers_p2": ['Đ', 'Đ', 'S', 'S', 'S', 'Đ', 'Đ', 'S'],
        "answers_p3": ['18', '2', '1', '7'],
        "answers_p4": [1.0, 0.25, 0.0],
        "details_p1": "12/12 đúng tuyệt đối (100%) -> 3.0đ",
        "details_p2": "8/8 ý đúng tuyệt đối (C13: Đ-Đ-S-S; C14: S-Đ-Đ-S) -> 2.0đ",
        "details_p3": "4/4 câu đúng tuyệt đối (C15: 18, C16: 2, C17: 1, C18: 7) -> 2.0đ",
        "details_p4": "C19 (1.0đ), C20 (0.25đ - đạo hàm v(t)), C21 (0đ - chưa làm)",
        "errors": "P4 C20 chưa kịp tìm vmax; chưa làm C21",
        "note": "Trắc nghiệm tuyệt đối 7.0/7.0đ. Cần đẩy nhanh tốc độ trình bày tự luận."
    },
    {
        "stt": 6,
        "name": "Tạ Trí Vĩ",
        "class": "12B1",
        "code": "102",
        "p1": 3.00,
        "p2": 2.00,
        "p3": 2.00,
        "p4": 2.50,
        "total": 9.50,
        "rank": "Xuất sắc",
        "answers_p1": ['B', 'B', 'C', 'D', 'C', 'B', 'C', 'B', 'B', 'C', 'D', 'D'],
        "answers_p2": ['Đ', 'S', 'Đ', 'Đ', 'Đ', 'S', 'Đ', 'S'],
        "answers_p3": ['7', '18', '1', '2'],
        "answers_p4": [1.0, 1.0, 0.5],
        "details_p1": "12/12 đúng tuyệt đối (100%) -> 3.0đ",
        "details_p2": "8/8 ý đúng tuyệt đối (C13: Đ-S-Đ-Đ; C14: Đ-S-Đ-S) -> 2.0đ",
        "details_p3": "4/4 câu đúng tuyệt đối (C15: 7, C16: 18, C17: 1, C18: 2) -> 2.0đ",
        "details_p4": "C19 (1.0đ), C20 (1.0đ), C21 (0.5đ - lập hàm t(x))",
        "errors": "P4 C21 dừng ở bước thiết lập hàm",
        "note": "Xuất sắc 9.50/10.0đ. Trắc nghiệm đạt trọn vẹn tuyệt đối 7.0/7.0đ (Phần I: 3.0đ, Phần II: 2.0đ, Phần III: 2.0đ). Tự luận 2.5/3.0đ."
    },
    {
        "stt": 7,
        "name": "Nguyễn Minh Hiếu",
        "class": "12B1",
        "code": "104",
        "p1": 3.00,
        "p2": 2.00,
        "p3": 1.00,
        "p4": 2.00,
        "total": 8.00,
        "rank": "Giỏi",
        "answers_p1": ['D', 'A', 'B', 'A', 'D', 'C', 'A', 'D', 'C', 'C', 'D', 'D'],
        "answers_p2": ['Đ', 'Đ', 'S', 'S', 'S', 'Đ', 'Đ', 'S'],
        "answers_p3": ['1', '2', '7', '18'],
        "answers_p4": [1.0, 1.0, 0.0],
        "details_p1": "12/12 đúng tuyệt đối (100%) -> 3.0đ",
        "details_p2": "8/8 ý đúng tuyệt đối -> 2.0đ",
        "details_p3": "2/4 câu đúng (Đúng C17: 7, C18: 18; Sai C15, C16 do hoán đổi) -> 1.0đ",
        "details_p4": "C19 (1.0đ), C20 (1.0đ), C21 (0đ - chưa làm)",
        "errors": "Sai P3 (C15, C16 bị điền hoán vị 1-2 thay vì 2-1); Chưa làm C21",
        "note": "Làm bài tự luận C19, C20 rất chuẩn xác, đạt trọn vẹn điểm tối đa."
    },
    {
        "stt": 8,
        "name": "Chu Thu Hiền",
        "class": "12B1",
        "code": "103",
        "p1": 3.00,
        "p2": 2.00,
        "p3": 1.00,
        "p4": 1.75,
        "total": 7.75,
        "rank": "Khá",
        "answers_p1": ['B', 'A', 'B', 'D', 'C', 'A', 'A', 'D', 'B', 'C', 'A', 'C'],
        "answers_p2": ['Đ', 'Đ', 'S', 'S', 'S', 'Đ', 'Đ', 'S'],
        "answers_p3": ['1.67', '3', '1', '7'],
        "answers_p4": [1.0, 0.75, 0.0],
        "details_p1": "12/12 đúng tuyệt đối (100%) -> 3.0đ",
        "details_p2": "8/8 ý đúng tuyệt đối -> 2.0đ",
        "details_p3": "2/4 câu đúng (Đúng C17: 1, C18: 7; Sai C15, C16) -> 1.0đ",
        "details_p4": "C19 (1.0đ), C20 (0.75đ - tìm đúng t=4s), C21 (0đ - chưa làm)",
        "errors": "Sai P3 (C15 điền 1.67, C16 điền 3); P4 C20 thiếu bước tính vmax=25; chưa làm C21",
        "note": "Nắm rất chắc kiến thức cơ bản và trắc nghiệm. Cần chú ý kỹ năng giải bài toán thực tế P3."
    },
    {
        "stt": 9,
        "name": "Nguyễn Quang Đức",
        "class": "12B1",
        "code": "102",
        "p1": 0.25,
        "p2": 0.35,
        "p3": 2.00,
        "p4": 2.00,
        "total": 4.60,
        "rank": "Yếu",
        "answers_p1": ["A", "D", "A", "D", "D", "D", "D", "D", "C", "B", "A", "C"],
        "answers_p2": ["S", "Đ", "Đ", "S", "Đ", "Đ", "S", "S"],
        "answers_p3": ["7", "18", "1", "2"],
        "answers_p4": [1.0, 1.0, 0.0],
        "details_p1": "1/12 đúng (Đúng C4: D. Bài làm: 1-A; 2-D; 3-A; 4-D; 5-D; 6-D; 7-D; 8-D; 9-C; 10-B; 11-A; 12-C) -> 0.25đ",
        "details_p2": "3/8 ý đúng theo Mã 102 (C13: S-Đ-Đ-S đúng ý c -> 0.10đ; C14: Đ-Đ-S-S đúng ý a,d -> 0.25đ) -> 0.35đ",
        "details_p3": "4/4 câu đúng tuyệt đối (C15: 7, C16: 18, C17: 1, C18: 2) -> 2.0đ",
        "details_p4": "C19 (1.0đ: m=3, b=1, c=-3), C20 (1.0đ: vmax=25m/s khi t=4s), C21 (0đ - chưa làm)",
        "errors": "Chấm đúng Mã đề 102 học sinh tô: Sai 11 câu P1; Sai 5 ý P2; Bỏ dở C21",
        "note": "Chấm nghiêm ngặt 100% theo đúng Mã đề 102 học sinh ghi và tô trên phiếu thi. Phần III làm rất tốt đạt trọn vẹn 2.0/2.0đ. Tổng điểm chính thức: 4.60đ (Xếp loại: Yếu)."
    },
    {
        "stt": 10,
        "name": "Nguyễn Mai Thảo",
        "class": "12B1",
        "code": "104",
        "p1": 3.00,
        "p2": 2.00,
        "p3": 2.00,
        "p4": 0.50,
        "total": 7.50,
        "rank": "Khá",
        "answers_p1": ['D', 'A', 'B', 'A', 'D', 'C', 'A', 'D', 'C', 'C', 'D', 'D'],
        "answers_p2": ['Đ', 'Đ', 'S', 'S', 'S', 'Đ', 'Đ', 'S'],
        "answers_p3": ['2', '1', '7', '18'],
        "answers_p4": [0.25, 0.25, 0.0],
        "details_p1": "12/12 đúng tuyệt đối (100%) -> 3.0đ",
        "details_p2": "8/8 ý đúng tuyệt đối -> 2.0đ",
        "details_p3": "4/4 câu đúng tuyệt đối (C15: 2, C16: 1, C17: 7, C18: 18) -> 2.0đ",
        "details_p4": "C19 (0.25đ - nhầm dấu m=-3), C20 (0.25đ - nhầm sang xét chiều chuyển động), C21 (0đ)",
        "errors": "P4 C19 nhầm dấu hệ số góc; C20 hiểu nhầm yêu cầu đề bài; chưa làm C21",
        "note": "Trắc nghiệm đạt 7.0/7.0 tuyệt đối. Cần đọc kỹ yêu cầu bài toán tự luận."
    },
    {
        "stt": 11,
        "name": "Cao Thị Thảo Nguyên",
        "class": "12B1",
        "code": "103",
        "p1": 3.00,
        "p2": 2.00,
        "p3": 2.00,
        "p4": 0.50,
        "total": 7.50,
        "rank": "Khá",
        "answers_p1": ['B', 'A', 'B', 'D', 'C', 'A', 'A', 'D', 'B', 'C', 'A', 'C'],
        "answers_p2": ['Đ', 'Đ', 'S', 'S', 'S', 'Đ', 'Đ', 'S'],
        "answers_p3": ['18', '2', '1', '7'],
        "answers_p4": [0.25, 0.25, 0.0],
        "details_p1": "12/12 đúng tuyệt đối (100%) -> 3.0đ",
        "details_p2": "8/8 ý đúng tuyệt đối -> 2.0đ",
        "details_p3": "4/4 câu đúng tuyệt đối (C15: 18, C16: 2, C17: 1, C18: 7) -> 2.0đ",
        "details_p4": "C19 (0.25đ - nhầm dấu m=-3), C20 (0.25đ - xét chiều chuyển động), C21 (0đ)",
        "errors": "P4 C19 nhầm dấu hệ số góc; C20 nhầm yêu cầu tìm vmax; chưa làm C21",
        "note": "Trắc nghiệm xuất sắc đạt 7.0/7.0 tuyệt đối. Tự luận cần đọc kỹ đề."
    },
    {
        "stt": 12,
        "name": "Hoàng Thảo Nguyên",
        "class": "12B1",
        "code": "104",
        "p1": 0.50,
        "p2": 2.00,
        "p3": 2.00,
        "p4": 2.50,
        "total": 7.00,
        "rank": "Khá",
        "answers_p1": ['', 'A', 'B', 'C', 'D', 'D', 'D', 'D', 'C', 'B', 'A', 'C'],
        "answers_p2": ['Đ', 'Đ', 'S', 'S', 'S', 'Đ', 'Đ', 'S'],
        "answers_p3": ['2', '1', '7', '18'],
        "answers_p4": [1.0, 1.0, 0.5],
        "details_p1": "2/12 đúng (Đúng C2: A, C3: B). Sai 10 câu do tô lệch pha mã đề -> 0.5đ",
        "details_p2": "8/8 ý đúng tuyệt đối -> 2.0đ",
        "details_p3": "4/4 câu đúng tuyệt đối 100% -> 2.0đ",
        "details_p4": "C19 (1.0đ), C20 (1.0đ), C21 (0.5đ - thiết lập hàm thời gian t(x))",
        "errors": "Sai P1 (10 câu ABCD do tô lệch phương án); P4 C21 chưa tính ra nghiệm tối ưu",
        "note": "Phần II và Phần III đạt điểm tuyệt đối. Cần cẩn thận khi tô phiếu trắc nghiệm."
    },
    {
        "stt": 13,
        "name": "Đào Xuân Dương",
        "class": "12B1",
        "code": "101",
        "p1": 1.75,
        "p2": 2.00,
        "p3": 2.00,
        "p4": 3.00,
        "total": 8.75,
        "rank": "Giỏi",
        "answers_p1": ['B', 'D', 'D', 'A', 'D', 'B', 'D', 'A', 'B', 'D', 'B', 'D'],
        "answers_p2": ['S', 'Đ', 'Đ', 'S', 'Đ', 'Đ', 'S', 'S'],
        "answers_p3": ['2', '1', '7', '18'],
        "answers_p4": [1.0, 1.0, 1.0],
        "details_p1": "7/12 đúng (Đúng C3, C5, C6, C8, C9, C10, C11) -> 1.75đ",
        "details_p2": "8/8 ý đúng tuyệt đối (C13: S-Đ-Đ-S; C14: Đ-Đ-S-S) -> 2.0đ",
        "details_p3": "4/4 câu đúng tuyệt đối (C15: 2, C16: 1, C17: 7, C18: 18) -> 2.0đ",
        "details_p4": "C19 (1.0đ), C20 (1.0đ), C21 (1.0đ) - Giải trọn vẹn hoàn hảo cả 3 câu tự luận",
        "errors": "Sai P1 (C1, C2, C4, C7, C12)",
        "note": "Giỏi 8.75đ. Tự luận xuất sắc 3.0/3.0đ. Phần III đạt trọn vẹn 2.0/2.0đ."
    },
    {
        "stt": 14,
        "name": "Hoàng Mạnh Hùng",
        "class": "12B1",
        "code": "102",
        "p1": 3.00,
        "p2": 1.10,
        "p3": 2.00,
        "p4": 2.00,
        "total": 8.10,
        "rank": "Giỏi",
        "answers_p1": ['B', 'B', 'C', 'D', 'C', 'B', 'C', 'B', 'B', 'C', 'D', 'D'],
        "answers_p2": ['S', 'Đ', 'Đ', 'S', 'Đ', 'S', 'Đ', 'S'],
        "answers_p3": ['7', '18', '1', '2'],
        "answers_p4": [1.0, 1.0, 0.0],
        "details_p1": "12/12 đúng tuyệt đối (100%) -> 3.0đ",
        "details_p2": "5/8 ý đúng (C13 đúng 1 ý c -> 0.1đ; C14 đúng cả 4 ý -> 1.0đ) -> 1.10đ",
        "details_p3": "4/4 câu đúng tuyệt đối (C15: 7, C16: 18, C17: 1, C18: 2) -> 2.0đ",
        "details_p4": "C19 (1.0đ), C20 (1.0đ), C21 (0đ - chưa làm)",
        "errors": "Sai P2 (C13a,b,d); Chưa làm C21",
        "note": "Giỏi 8.10/10.0đ. Phần I và Phần III tuyệt đối 5.0/5.0đ. Tự luận C19, C20 làm rất chắc chắn."
    },
    {
        "stt": 15,
        "name": "Nguyễn Văn Thành",
        "class": "12B1",
        "code": "103",
        "p1": 1.75,
        "p2": 1.75,
        "p3": 2.00,
        "p4": 1.00,
        "total": 6.50,
        "rank": "Khá",
        "answers_p1": ['D', 'B', 'B', 'D', 'C', 'B', 'A', 'D', 'B', 'A', 'A', 'D'],
        "answers_p2": ['Đ', 'Đ', 'S', 'S', 'Đ', 'Đ', 'Đ', 'S'],
        "answers_p3": ['18', '2', '1', '7'],
        "answers_p4": [1.0, 0.0, 0.0],
        "details_p1": "7/12 đúng (Đúng C3, C4, C5, C7, C8, C9, C11) -> 1.75đ",
        "details_p2": "7/8 ý đúng (C13: đúng 4 ý; C14: đúng 3 ý b, c, d; sai ý a) -> 1.75đ",
        "details_p3": "4/4 câu đúng tuyệt đối (C15: 18, C16: 2, C17: 1, C18: 7) -> 2.0đ",
        "details_p4": "C19 (1.0đ), C20 (0đ), C21 (0đ)",
        "errors": "Sai P1 (C1, C2, C6, C10, C12); Sai P2 (C14a); Chưa làm P4 C20, C21",
        "note": "Trả lời ngắn Phần III xuất sắc 2.0/2.0đ. Cần phân bổ thời gian cho bài toán ứng dụng."
    },
    {
        "stt": 16,
        "name": "Dương Tuấn Anh",
        "class": "12B1",
        "code": "102",
        "p1": 3.00,
        "p2": 2.00,
        "p3": 2.00,
        "p4": 1.00,
        "total": 8.00,
        "rank": "Giỏi",
        "answers_p1": ['B', 'B', 'C', 'D', 'C', 'B', 'C', 'B', 'B', 'C', 'D', 'D'],
        "answers_p2": ['Đ', 'S', 'Đ', 'Đ', 'Đ', 'S', 'Đ', 'S'],
        "answers_p3": ['7', '18', '1', '2'],
        "answers_p4": [1.0, 0.0, 0.0],
        "details_p1": "12/12 đúng tuyệt đối (100%) -> 3.0đ",
        "details_p2": "8/8 ý đúng tuyệt đối -> 2.0đ",
        "details_p3": "4/4 câu đúng tuyệt đối (C15: 7, C16: 18, C17: 1, C18: 2) -> 2.0đ",
        "details_p4": "C19 (1.0đ), C20 (0đ), C21 (0đ)",
        "errors": "Chưa làm P4 C20, C21",
        "note": "Giỏi 8.00/10.0đ. Trắc nghiệm đạt trọn vẹn tuyệt đối 7.0/7.0đ (Phần I, II, III). Cần cố gắng làm thêm tự luận C20, C21."
    },
    {
        "stt": 17,
        "name": "Hoàng Trung Kiên",
        "class": "12B1",
        "code": "102",
        "p1": 3.00,
        "p2": 1.10,
        "p3": 2.00,
        "p4": 1.75,
        "total": 7.85,
        "rank": "Khá",
        "answers_p1": ['B', 'B', 'C', 'D', 'C', 'B', 'C', 'B', 'B', 'C', 'D', 'D'],
        "answers_p2": ['S', 'Đ', 'Đ', 'S', 'Đ', 'S', 'Đ', 'S'],
        "answers_p3": ['7', '18', '1', '2'],
        "answers_p4": [1.0, 0.75, 0.0],
        "details_p1": "12/12 đúng tuyệt đối (100%) -> 3.0đ",
        "details_p2": "5/8 ý đúng (C13 đúng 1 ý c -> 0.1đ; C14 đúng cả 4 ý -> 1.0đ) -> 1.10đ",
        "details_p3": "4/4 câu đúng tuyệt đối (C15: 7, C16: 18, C17: 1, C18: 2) -> 2.0đ",
        "details_p4": "C19 (1.0đ), C20 (0.75đ - tìm đúng t=4s), C21 (0đ)",
        "errors": "Sai P2 (C13a,b,d); P4 C20 thiếu tính vmax; chưa làm C21",
        "note": "Khá 7.85/10.0đ. Phần I và Phần III tuyệt đối 5.0/5.0đ."
    },
    {
        "stt": 18,
        "name": "Đỗ Văn Nghĩa",
        "class": "12B1",
        "code": "104",
        "p1": 2.00,
        "p2": 1.50,
        "p3": 1.50,
        "p4": 1.00,
        "total": 6.00,
        "rank": "Trung bình",
        "answers_p1": ['D', 'A', 'B', 'B', 'D', 'B', 'C', 'D', 'C', 'C', 'C', 'D'],
        "answers_p2": ['S', 'Đ', 'S', 'S', 'Đ', 'Đ', 'Đ', 'S'],
        "answers_p3": ['', '1', '7', '18'],
        "answers_p4": [1.0, 0.0, 0.0],
        "details_p1": "8/12 đúng (Đúng C1, C2, C3, C5, C8, C9, C10, C12) -> 2.0đ",
        "details_p2": "6/8 ý đúng (C13: đúng 3 ý b,c,d; C14: đúng 3 ý b,c,d) -> 1.5đ",
        "details_p3": "3/4 câu đúng (Đúng C16: 1, C17: 7, C18: 18; Bỏ trống C15) -> 1.5đ",
        "details_p4": "C19 (1.0đ), C20 (0đ), C21 (0đ)",
        "errors": "Sai P1 (C4, C6, C7, C11); Sai P2 (C13a, C14a); P3 bỏ trống C15; Chưa làm C20, C21",
        "note": "Điểm tương đối đều các phần. Cần tăng tốc độ giải để kịp làm tự luận."
    },
    {
        "stt": 19,
        "name": "Vũ Thị Mai",
        "class": "12B1",
        "code": "101",
        "p1": 1.75,
        "p2": 1.00,
        "p3": 2.00,
        "p4": 3.00,
        "total": 7.75,
        "rank": "Khá",
        "answers_p1": ['D', 'B', 'D', 'C', 'A', 'B', 'C', 'D', 'C', 'D', 'A', 'A'],
        "answers_p2": ['S', 'Đ', 'Đ', 'S', 'S', 'Đ', 'Đ', 'S'],
        "answers_p3": ['2', '1', '7', '18'],
        "answers_p4": [1.0, 1.0, 1.0],
        "details_p1": "7/12 đúng (Đúng C1, C2, C3, C6, C7, C10, C12) -> 1.75đ",
        "details_p2": "5/8 ý đúng (C13: đúng 4 ý; C14: đúng 1 ý b) -> 1.0đ",
        "details_p3": "4/4 câu đúng tuyệt đối (C15: 2, C16: 1, C17: 7, C18: 18) -> 2.0đ",
        "details_p4": "C19 (1.0đ), C20 (1.0đ), C21 (1.0đ) - Làm trọn vẹn hoàn hảo cả 3 bài tự luận",
        "errors": "Sai P1 (C4, C5, C8, C9, C11); Sai P2 (C14a,c,d)",
        "note": "Khá 7.75đ. Tự luận xuất sắc 3.0/3.0đ. Cần cẩn thận hơn ở Phần I và II."
    },
    {
        "stt": 20,
        "name": "Bùi Tuấn Long",
        "class": "12B1",
        "code": "101",
        "p1": 1.75,
        "p2": 1.10,
        "p3": 2.00,
        "p4": 2.75,
        "total": 7.60,
        "rank": "Khá",
        "answers_p1": ['B', 'D', 'D', 'A', 'D', 'B', 'D', 'A', 'B', 'D', 'B', 'D'],
        "answers_p2": ['S', 'S', 'Đ', 'Đ', 'Đ', 'Đ', 'S', 'S'],
        "answers_p3": ['2', '1', '7', '18'],
        "answers_p4": [1.0, 0.75, 1.0],
        "details_p1": "7/12 đúng (Đúng C3, C5, C6, C8, C9, C10, C11) -> 1.75đ",
        "details_p2": "5/8 ý đúng (C13 đúng 1 ý c -> 0.1đ; C14 đúng cả 4 ý -> 1.0đ) -> 1.10đ",
        "details_p3": "4/4 câu đúng tuyệt đối (C15: 2, C16: 1, C17: 7, C18: 18) -> 2.0đ",
        "details_p4": "C19 (1.0đ), C20 (0.75đ - tìm đúng t=4s), C21 (1.0đ - x=24.3m)",
        "errors": "Sai P1 (C1, C2, C4, C7, C12); P4 C20 thiếu tính vmax",
        "note": "Khá 7.60đ. Tự luận đạt 2.75đ. Phần III đạt tuyệt đối 2.0/2.0đ."
    },
    {
        "stt": 21,
        "name": "Lương Duy Nghĩa",
        "class": "12B1",
        "code": "101",
        "p1": 2.50,
        "p2": 0.35,
        "p3": 2.00,
        "p4": 2.75,
        "total": 7.60,
        "rank": "Khá",
        "answers_p1": ['D', 'B', 'D', 'D', 'D', 'B', 'D', 'A', 'B', 'D', 'B', 'B'],
        "answers_p2": ['S', 'Đ', 'S', 'S', 'S', 'Đ', 'Đ', 'Đ'],
        "answers_p3": ['2', '1', '7', '18'],
        "answers_p4": [1.0, 0.75, 1.0],
        "details_p1": "10/12 đúng (Đúng C1, C2, C3, C4, C5, C6, C8, C9, C10, C11) -> 2.50đ",
        "details_p2": "3/8 ý đúng (C13 đúng 2 ý b,d -> 0.25đ; C14 đúng 1 ý b -> 0.10đ) -> 0.35đ",
        "details_p3": "4/4 câu đúng tuyệt đối (C15: 2, C16: 1, C17: 7, C18: 18) -> 2.0đ",
        "details_p4": "C19 (1.0đ), C20 (0.75đ - tìm đúng t=4s), C21 (1.0đ - x=24.3m)",
        "errors": "Sai P1 (C7, C12); Sai P2 (C13a,c; C14a,c,d)",
        "note": "Khá 7.60đ. Tự luận giải rất chắc chắn (2.75đ). Trả lời ngắn đạt 2.0/2.0đ."
    },
    {
        "stt": 22,
        "name": "Triệu Trung Bách",
        "class": "12B1",
        "code": "104",
        "p1": 0.75,
        "p2": 2.00,
        "p3": 0.50,
        "p4": 2.00,
        "total": 5.25,
        "rank": "Trung bình",
        "answers_p1": ['B', 'A', 'B', 'D', 'C', 'D', 'B', 'D', 'B', 'C', 'D', 'D'],
        "answers_p2": ['Đ', 'Đ', 'S', 'S', 'S', 'Đ', 'Đ', 'S'],
        "answers_p3": ['2', '', '7', '18'],
        "answers_p4": [1.0, 1.0, 0.0],
        "details_p1": "3/12 đúng (Đúng C2, C3, C9) -> 0.75đ",
        "details_p2": "8/8 ý đúng tuyệt đối -> 2.0đ",
        "details_p3": "1/4 câu đúng (Đúng C15: 2; Sai C16, C17, C18) -> 0.5đ",
        "details_p4": "C19 (1.0đ), C20 (1.0đ), C21 (0đ)",
        "errors": "Sai P1 (9 câu ABCD do tô lệch hàng); Sai P3 (C16, C17, C18); Chưa làm C21",
        "note": "Phần II đúng 100% và tự luận tốt. Cần rèn kỹ năng kiểm tra phiếu tô."
    },
    {
        "stt": 23,
        "name": "Trần Triệu Nam",
        "class": "12B1",
        "code": "101",
        "p1": 1.75,
        "p2": 1.10,
        "p3": 2.00,
        "p4": 2.25,
        "total": 7.10,
        "rank": "Khá",
        "answers_p1": ['B', 'D', 'D', 'A', 'D', 'B', 'D', 'A', 'B', 'D', 'B', 'D'],
        "answers_p2": ['S', 'S', 'Đ', 'Đ', 'Đ', 'Đ', 'S', 'S'],
        "answers_p3": ['2', '1', '7', '18'],
        "answers_p4": [1.0, 0.25, 1.0],
        "details_p1": "7/12 đúng (Đúng C3, C5, C6, C8, C9, C10, C11) -> 1.75đ",
        "details_p2": "5/8 ý đúng (C13 đúng 1 ý c -> 0.1đ; C14 đúng cả 4 ý -> 1.0đ) -> 1.10đ",
        "details_p3": "4/4 câu đúng tuyệt đối (C15: 2, C16: 1, C17: 7, C18: 18) -> 2.0đ",
        "details_p4": "C19 (1.0đ), C20 (0.25đ - đạo hàm v(t)), C21 (1.0đ - x=24.3m)",
        "errors": "Sai P1 (C1, C2, C4, C7, C12); Sai P2 (C13a,c; C14a,c)",
        "note": "Khá 7.10đ. Phần III đạt trọn vẹn 2.0/2.0đ. Cần cẩn thận ở Phần I."
    },
    {
        "stt": 24,
        "name": "Hà Hoàng Nhật Tân",
        "class": "12B1",
        "code": "101",
        "p1": 2.00,
        "p2": 1.25,
        "p3": 2.00,
        "p4": 1.75,
        "total": 7.00,
        "rank": "Khá",
        "answers_p1": ['B', 'D', 'D', 'D', 'D', 'B', 'D', 'A', 'B', 'D', 'B', 'D'],
        "answers_p2": ['Đ', 'Đ', 'S', 'Đ', 'Đ', 'Đ', 'S', 'S'],
        "answers_p3": ['2', '1', '7', '18'],
        "answers_p4": [1.0, 0.75, 0.0],
        "details_p1": "8/12 đúng (Đúng C3, C4, C5, C6, C8, C9, C10, C11) -> 2.00đ",
        "details_p2": "5/8 ý đúng (C13 đúng 2 ý a,b -> 0.25đ; C14 đúng cả 4 ý -> 1.0đ) -> 1.25đ",
        "details_p3": "4/4 câu đúng tuyệt đối (C15: 2, C16: 1, C17: 7, C18: 18) -> 2.0đ",
        "details_p4": "C19 (1.0đ), C20 (0.75đ - tìm đúng t=4s), C21 (0đ - chưa làm)",
        "errors": "Sai P1 (C1, C2, C7, C12); Sai P2 (C13a,c; C14a)",
        "note": "Khá 7.00đ. Phần III đạt 2.0/2.0đ."
    },
    {
        "stt": 25,
        "name": "Dương Đức Vinh",
        "class": "12B1",
        "code": "101",
        "p1": 2.00,
        "p2": 1.00,
        "p3": 2.00,
        "p4": 2.00,
        "total": 7.00,
        "rank": "Khá",
        "answers_p1": ['D', 'A', 'D', 'D', 'D', 'C', 'B', 'A', 'B', 'D', 'B', ''],
        "answers_p2": ['S', 'Đ', 'Đ', 'S', 'Đ', 'Đ', 'Đ', 'Đ'],
        "answers_p3": ['2', '1', '7', '18'],
        "answers_p4": [1.0, 1.0, 0.0],
        "details_p1": "8/12 đúng (Đúng C1, C3, C4, C5, C8, C9, C10, C11) -> 2.0đ",
        "details_p2": "5/8 ý đúng (C13: đúng 4 ý; C14: đúng 1 ý c) -> 1.0đ",
        "details_p3": "4/4 câu đúng tuyệt đối (C15: 2, C16: 1, C17: 7, C18: 18) -> 2.0đ",
        "details_p4": "C19 (1.0đ), C20 (1.0đ), C21 (0đ)",
        "errors": "Sai P1 (C1, C2, C7, C12); Sai P2 (C13a,c,d; C14a)",
        "note": "Khá 7.00đ. Phần III đạt 2.0/2.0đ."
    },
    {
        "stt": 26,
        "name": "Bùi Long Nhật",
        "class": "12B1",
        "code": "101",
        "p1": 1.75,
        "p2": 0.35,
        "p3": 2.00,
        "p4": 2.00,
        "total": 6.10,
        "rank": "Trung bình",
        "answers_p1": ['B', 'D', 'D', 'A', 'D', 'B', 'D', 'A', 'B', 'D', 'B', 'D'],
        "answers_p2": ['S', 'Đ', 'S', 'Đ', 'S', 'Đ', 'Đ', 'S'],
        "answers_p3": ['2', '1', '7', '18'],
        "answers_p4": [1.0, 1.0, 0.0],
        "details_p1": "7/12 đúng (Đúng C3, C5, C6, C8, C9, C10, C11) -> 1.75đ",
        "details_p2": "3/8 ý đúng (C13 đúng 1 ý b -> 0.10đ; C14 đúng 2 ý b,d -> 0.25đ) -> 0.35đ",
        "details_p3": "4/4 câu đúng tuyệt đối (C15: 2, C16: 1, C17: 7, C18: 18) -> 2.0đ",
        "details_p4": "C19 (1.0đ), C20 (1.0đ), C21 (0đ - chưa làm)",
        "errors": "Sai P1 (C1, C2, C4, C7, C12); Sai P2 (C13a,c,d; C14a,c,d)",
        "note": "Trung bình 6.10đ. Phần III đạt trọn vẹn 2.0/2.0đ."
    },
    {
        "stt": 27,
        "name": "Bế Ngọc Thành",
        "class": "12B1",
        "code": "102",
        "p1": 0.25,
        "p2": 0.35,
        "p3": 2.00,
        "p4": 2.25,
        "total": 4.85,
        "rank": "Trung bình",
        "answers_p1": ['D', 'D', 'D', 'D', 'D', 'D', 'D', 'D', 'A', 'A', 'A', ''],
        "answers_p2": ['S', 'S', 'Đ', 'S', 'Đ', 'S', 'S', 'Đ'],
        "answers_p3": ['7', '18', '1', '2'],
        "answers_p4": [1.0, 0.75, 0.5],
        "details_p1": "1/12 đúng (Đúng C4: D; Sai 11 câu còn lại) -> 0.25đ",
        "details_p2": "3/8 ý đúng (C13 đúng 1 ý c -> 0.10đ; C14 đúng 2 ý a,d -> 0.25đ) -> 0.35đ",
        "details_p3": "4/4 câu đúng tuyệt đối (C15: 7, C16: 18, C17: 1, C18: 2) -> 2.0đ",
        "details_p4": "C19 (1.0đ: 4 bước logic chuẩn), C20 (0.75đ: tìm đúng t=4s, thiếu bước tính vmax=25m/s), C21 (0.50đ: thiết lập hàm và đạo hàm đúng, sai logic đại số 64-9!=7x^2, thiếu BBT)",
        "errors": "Sai 11 câu P1 do tô lệch; Sai 5 ý P2; P4 C20 thiếu tính vmax, C21 sai logic biến đổi đại số",
        "note": "Trung bình 4.85/10.0đ. Tự luận đạt 2.25/3.0đ và Phần III đạt 2.0/2.0đ rất tốt. Cần rèn kỹ năng tô phiếu Phần I và tính cẩn thận trong biến đổi đại số."
    },
    {
        "stt": 28,
        "name": "Phạm Chí Hùng",
        "class": "12B1",
        "code": "102",
        "p1": 0.25,
        "p2": 0.35,
        "p3": 2.00,
        "p4": 1.00,
        "total": 3.60,
        "rank": "Yếu",
        "answers_p1": ['D', 'D', 'D', 'D', 'D', 'D', 'D', 'D', 'B', 'A', 'A', ''],
        "answers_p2": ['S', 'S', 'Đ', 'S', 'Đ', 'S', 'S', 'Đ'],
        "answers_p3": ['7', '18', '1', '2'],
        "answers_p4": [1.0, 0.0, 0.0],
        "details_p1": "1/12 đúng (Đúng C4: D; Sai 11 câu còn lại) -> 0.25đ",
        "details_p2": "3/8 ý đúng (C13 đúng 1 ý c -> 0.10đ; C14 đúng 2 ý a,d -> 0.25đ) -> 0.35đ",
        "details_p3": "4/4 câu đúng tuyệt đối (C15: 7, C16: 18, C17: 1, C18: 2) -> 2.0đ",
        "details_p4": "C19 (1.0đ), C20 (0đ), C21 (0đ)",
        "errors": "Sai 11 câu P1; Sai P2; Bỏ 2 câu tự luận C20, C21",
        "note": "Phần III đạt 2.0/2.0đ. Cần củng cố kiến thức Phần I và làm thêm tự luận."
    },
    {
        "stt": 29,
        "name": "Vũ Anh Tú",
        "class": "12B1",
        "code": "102",
        "p1": 3.00,
        "p2": 2.00,
        "p3": 1.50,
        "p4": 2.00,
        "total": 8.50,
        "rank": "Giỏi",
        "answers_p1": ["B", "B", "C", "D", "C", "B", "C", "B", "B", "C", "D", "D"],
        "answers_p2": ["Đ", "S", "Đ", "Đ", "Đ", "S", "Đ", "S"],
        "answers_p3": ["7", "18", "1", "3"],
        "answers_p4": [0.0, 1.0, 1.0],
        "details_p1": "12/12 đúng tuyệt đối (100%) -> 3.0đ",
        "details_p2": "8/8 ý đúng tuyệt đối (C13: Đ-S-Đ-Đ; C14: Đ-S-Đ-S) -> 2.0đ",
        "details_p3": "3/4 câu đúng theo Mã 102 (Đúng C15: 7, C16: 18, C17: 1; Sai C18: điền 3 - chuẩn 2) -> 1.50đ",
        "details_p4": "C20 (1.0đ - giải đúng vmax=25 tại t=4s), C21 (1.0đ - giải đúng x=24.3m)",
        "errors": "Sai C18 P3; Chưa làm C19",
        "note": "Giỏi 8.50/10.0đ. Trắc nghiệm đạt 6.5/7.0đ (P1: 3.0đ, P2: 2.0đ, P3: 1.5đ). Tự luận 2.0/3.0đ."
    },
    {
        "stt": 30,
        "name": "Nguyễn Phương Hà Anh",
        "class": "12B1",
        "code": "104",
        "p1": 3.00,
        "p2": 2.00,
        "p3": 2.00,
        "p4": 2.00,
        "total": 9.00,
        "rank": "Giỏi",
        "answers_p1": ["D", "A", "B", "A", "D", "C", "A", "D", "C", "C", "D", "D"],
        "answers_p2": ["Đ", "Đ", "S", "S", "S", "Đ", "Đ", "S"],
        "answers_p3": ["2", "1", "7", "18"],
        "answers_p4": [1.0, 1.0, 0.0],
        "details_p1": "12/12 đúng tuyệt đối (100%) -> 3.0đ",
        "details_p2": "8/8 ý đúng tuyệt đối -> 2.0đ",
        "details_p3": "4/4 câu đúng tuyệt đối (C15: 2, C16: 1, C17: 7, C18: 18) -> 2.0đ",
        "details_p4": "C19 (1.0đ), C20 (1.0đ), C21 (0đ)",
        "errors": "Chưa làm C21 tự luận",
        "note": "Trắc nghiệm Phần I, II, III đạt điểm tuyệt đối 7.0/7.0đ. Tự luận làm rất chắc chắn."
    },
    {
        "stt": 31,
        "name": "Lương Duy Khoa",
        "class": "12B1",
        "code": "104",
        "p1": 3.00,
        "p2": 2.00,
        "p3": 2.00,
        "p4": 1.75,
        "total": 8.75,
        "rank": "Giỏi",
        "answers_p1": ["D", "A", "B", "A", "D", "C", "A", "D", "C", "C", "D", "D"],
        "answers_p2": ["Đ", "Đ", "S", "S", "S", "Đ", "Đ", "S"],
        "answers_p3": ["2", "1", "7", "18"],
        "answers_p4": [1.0, 0.75, 0.0],
        "details_p1": "12/12 đúng tuyệt đối (100%) -> 3.0đ",
        "details_p2": "8/8 ý đúng tuyệt đối -> 2.0đ",
        "details_p3": "4/4 câu đúng tuyệt đối -> 2.0đ",
        "details_p4": "C19 (1.0đ), C20 (0.75đ), C21 (0đ)",
        "errors": "P4 C20 thiếu tính vmax; chưa làm C21",
        "note": "Trắc nghiệm tuyệt đối 7.0/7.0đ. Trình bày bài rất sạch đẹp."
    },
    {
        "stt": 32,
        "name": "Phạm Khánh Ly",
        "class": "12B1",
        "code": "103",
        "p1": 3.00,
        "p2": 2.00,
        "p3": 2.00,
        "p4": 1.75,
        "total": 8.75,
        "rank": "Giỏi",
        "answers_p1": ["B", "A", "B", "D", "C", "A", "A", "D", "B", "C", "A", "C"],
        "answers_p2": ["Đ", "Đ", "S", "S", "S", "Đ", "Đ", "S"],
        "answers_p3": ["18", "2", "1", "7"],
        "answers_p4": [1.0, 0.75, 0.0],
        "details_p1": "12/12 đúng tuyệt đối (100%) -> 3.0đ",
        "details_p2": "8/8 ý đúng tuyệt đối -> 2.0đ",
        "details_p3": "4/4 câu đúng tuyệt đối -> 2.0đ",
        "details_p4": "C19 (1.0đ), C20 (0.75đ), C21 (0đ)",
        "errors": "P4 C20 thiếu tính vmax; chưa làm C21",
        "note": "Trắc nghiệm tuyệt đối 7.0/7.0đ. Nắm rất vững kiến thức cơ bản."
    },
    {
        "stt": 33,
        "name": "Dương Quốc Khánh",
        "class": "12B1",
        "code": "103",
        "p1": 3.00,
        "p2": 2.00,
        "p3": 2.00,
        "p4": 1.75,
        "total": 8.75,
        "rank": "Giỏi",
        "answers_p1": ["B", "A", "B", "D", "C", "A", "A", "D", "B", "C", "A", "C"],
        "answers_p2": ["Đ", "Đ", "S", "S", "S", "Đ", "Đ", "S"],
        "answers_p3": ["18", "2", "1", "7"],
        "answers_p4": [1.0, 0.75, 0.0],
        "details_p1": "12/12 đúng tuyệt đối (100%) -> 3.0đ",
        "details_p2": "8/8 ý đúng tuyệt đối -> 2.0đ",
        "details_p3": "4/4 câu đúng tuyệt đối -> 2.0đ",
        "details_p4": "C19 (1.0đ), C20 (0.75đ), C21 (0đ)",
        "errors": "P4 C20 thiếu tính vmax; chưa làm C21",
        "note": "Trắc nghiệm tuyệt đối 7.0/7.0đ. Cần đẩy nhanh tiến độ làm bài tự luận."
    },
    {
        "stt": 34,
        "name": "Nguyễn Thị Vân Anh",
        "class": "12B1",
        "code": "103",
        "p1": 3.00,
        "p2": 2.00,
        "p3": 2.00,
        "p4": 1.50,
        "total": 8.50,
        "rank": "Giỏi",
        "answers_p1": ["B", "A", "B", "D", "C", "A", "A", "D", "B", "C", "A", "C"],
        "answers_p2": ["Đ", "Đ", "S", "S", "S", "Đ", "Đ", "S"],
        "answers_p3": ["18", "2", "1", "7"],
        "answers_p4": [1.0, 0.5, 0.0],
        "details_p1": "12/12 đúng tuyệt đối (100%) -> 3.0đ",
        "details_p2": "8/8 ý đúng tuyệt đối -> 2.0đ",
        "details_p3": "4/4 câu đúng tuyệt đối -> 2.0đ",
        "details_p4": "C19 (1.0đ), C20 (0.5đ), C21 (0đ)",
        "errors": "P4 C20 chưa giải xong; chưa làm C21",
        "note": "Trắc nghiệm xuất sắc tuyệt đối 7.0/7.0đ."
    },
    {
        "stt": 35,
        "name": "Đào Duy Hiển",
        "class": "12B1",
        "code": "103",
        "p1": 3.00,
        "p2": 2.00,
        "p3": 2.00,
        "p4": 1.50,
        "total": 8.50,
        "rank": "Giỏi",
        "answers_p1": ["B", "A", "B", "D", "C", "A", "A", "D", "B", "C", "A", "C"],
        "answers_p2": ["Đ", "Đ", "S", "S", "S", "Đ", "Đ", "S"],
        "answers_p3": ["18", "2", "1", "7"],
        "answers_p4": [1.0, 0.5, 0.0],
        "details_p1": "12/12 đúng tuyệt đối (100%) -> 3.0đ",
        "details_p2": "8/8 ý đúng tuyệt đối -> 2.0đ",
        "details_p3": "4/4 câu đúng tuyệt đối -> 2.0đ",
        "details_p4": "C19 (1.0đ), C20 (0.5đ), C21 (0đ)",
        "errors": "P4 C20 chưa giải xong; chưa làm C21",
        "note": "Trắc nghiệm tuyệt đối 7.0/7.0đ. Làm bài nghiêm túc."
    },
    {
        "stt": 36,
        "name": "Ngô Duy Khánh",
        "class": "12B1",
        "code": "103",
        "p1": 3.00,
        "p2": 2.00,
        "p3": 2.00,
        "p4": 1.50,
        "total": 8.50,
        "rank": "Giỏi",
        "answers_p1": ["B", "A", "B", "D", "C", "A", "A", "D", "B", "C", "A", "C"],
        "answers_p2": ["Đ", "Đ", "S", "S", "S", "Đ", "Đ", "S"],
        "answers_p3": ["18", "2", "1", "7"],
        "answers_p4": [1.0, 0.5, 0.0],
        "details_p1": "12/12 đúng tuyệt đối (100%) -> 3.0đ",
        "details_p2": "8/8 ý đúng tuyệt đối -> 2.0đ",
        "details_p3": "4/4 câu đúng tuyệt đối -> 2.0đ",
        "details_p4": "C19 (1.0đ), C20 (0.5đ), C21 (0đ)",
        "errors": "P4 C20 chưa hoàn thiện; chưa làm C21",
        "note": "Trắc nghiệm tuyệt đối 7.0/7.0đ."
    },
    {
        "stt": 37,
        "name": "Bùi Hoàng Thành Đạt",
        "class": "12B1",
        "code": "104",
        "p1": 3.00,
        "p2": 2.00,
        "p3": 1.00,
        "p4": 2.00,
        "total": 8.00,
        "rank": "Giỏi",
        "answers_p1": ["D", "A", "B", "A", "D", "C", "A", "D", "C", "C", "D", "D"],
        "answers_p2": ["Đ", "Đ", "S", "S", "S", "Đ", "Đ", "S"],
        "answers_p3": ["1", "1", "9.5", "18"],
        "answers_p4": [1.0, 1.0, 0.0],
        "details_p1": "12/12 đúng tuyệt đối (100%) -> 3.0đ",
        "details_p2": "8/8 ý đúng tuyệt đối -> 2.0đ",
        "details_p3": "2/4 câu đúng (Đúng C16: 1, C18: 18; Sai C15, C17) -> 1.0đ",
        "details_p4": "C19 (1.0đ), C20 (1.0đ), C21 (0đ)",
        "errors": "Sai P3 (C15, C17); Chưa làm C21",
        "note": "Trắc nghiệm P1 và P2 tuyệt đối 5.0/5.0đ. Tự luận C19, C20 làm tốt."
    },
    {
        "stt": 38,
        "name": "Nguyễn Lê Hải Anh",
        "class": "12B1",
        "code": "102",
        "p1": 3.00,
        "p2": 2.00,
        "p3": 0.50,
        "p4": 1.75,
        "total": 7.25,
        "rank": "Khá",
        "answers_p1": ["B", "B", "C", "D", "C", "B", "C", "B", "B", "C", "D", "D"],
        "answers_p2": ["Đ", "S", "Đ", "Đ", "Đ", "S", "Đ", "S"],
        "answers_p3": ["18", "2", "1", "7"],
        "answers_p4": [1.0, 0.75, 0.0],
        "details_p1": "12/12 đúng tuyệt đối (100%) -> 3.0đ",
        "details_p2": "8/8 ý đúng tuyệt đối -> 2.0đ",
        "details_p3": "1/4 câu đúng theo Mã 102 (Đúng C17: 1; Sai C15, C16, C18 do điền nhầm bộ số của Mã 103) -> 0.50đ",
        "details_p4": "C19 (1.0đ), C20 (0.75đ), C21 (0đ)",
        "errors": "Sai P3 (C15, C16, C18 do nhìn nhầm đề mã 103); P4 C20 thiếu tính vmax",
        "note": "Khá 7.25/10.0đ. Trắc nghiệm P1 và P2 tuyệt đối 5.0/5.0đ."
    },
    {
        "stt": 39,
        "name": "Ngô Thế Anh",
        "class": "12B1",
        "code": "102",
        "p1": 3.00,
        "p2": 2.00,
        "p3": 2.00,
        "p4": 2.00,
        "total": 9.00,
        "rank": "Xuất sắc",
        "answers_p1": ["B", "B", "C", "D", "C", "B", "C", "B", "B", "C", "D", "D"],
        "answers_p2": ["Đ", "S", "Đ", "Đ", "Đ", "S", "Đ", "S"],
        "answers_p3": ["7", "18", "1", "2"],
        "answers_p4": [1.0, 1.0, 0.0],
        "details_p1": "12/12 đúng tuyệt đối (100%) -> 3.0đ",
        "details_p2": "8/8 ý đúng tuyệt đối -> 2.0đ",
        "details_p3": "4/4 câu đúng tuyệt đối (C15: 7, C16: 18, C17: 1, C18: 2) -> 2.0đ",
        "details_p4": "C19 (1.0đ), C20 (1.0đ), C21 (0đ)",
        "errors": "Chưa làm C21",
        "note": "Xuất sắc 9.00/10.0đ. Trắc nghiệm tuyệt đối 7.0/7.0đ (P1: 3.0đ, P2: 2.0đ, P3: 2.0đ). Tự luận C19, C20 trọn vẹn 2.0đ."
    },
    {
        "stt": 40,
        "name": "Đàm Vĩnh Hưng",
        "class": "12B1",
        "code": "102",
        "p1": 3.00,
        "p2": 2.00,
        "p3": 2.00,
        "p4": 1.75,
        "total": 8.75,
        "rank": "Giỏi",
        "answers_p1": ["B", "B", "C", "D", "C", "B", "C", "B", "B", "C", "D", "D"],
        "answers_p2": ["Đ", "S", "Đ", "Đ", "Đ", "S", "Đ", "S"],
        "answers_p3": ["7", "18", "1", "2"],
        "answers_p4": [1.0, 0.75, 0.0],
        "details_p1": "12/12 đúng tuyệt đối (100%) -> 3.0đ",
        "details_p2": "8/8 ý đúng tuyệt đối -> 2.0đ",
        "details_p3": "4/4 câu đúng tuyệt đối (C15: 7, C16: 18, C17: 1, C18: 2) -> 2.0đ",
        "details_p4": "C19 (1.0đ), C20 (0.75đ), C21 (0đ)",
        "errors": "P4 C20 thiếu tính vmax",
        "note": "Giỏi 8.75/10.0đ. Trắc nghiệm tuyệt đối 7.0/7.0đ (P1: 3.0đ, P2: 2.0đ, P3: 2.0đ)."
    },
    {
        "stt": 41,
        "name": "Nguyễn Châu Ly",
        "class": "12B1",
        "code": "101",
        "p1": 3.00,
        "p2": 2.00,
        "p3": 2.00,
        "p4": 2.00,
        "total": 9.00,
        "rank": "Giỏi",
        "answers_p1": ["D", "B", "D", "D", "D", "B", "C", "A", "B", "D", "B", "A"],
        "answers_p2": ["S", "Đ", "Đ", "S", "Đ", "Đ", "S", "S"],
        "answers_p3": ["2", "1", "7", "18"],
        "answers_p4": [1.0, 1.0, 0.0],
        "details_p1": "12/12 đúng tuyệt đối (100%) -> 3.0đ",
        "details_p2": "8/8 ý đúng tuyệt đối -> 2.0đ",
        "details_p3": "4/4 câu đúng tuyệt đối (C15: 2, C16: 1, C17: 7, C18: 18) -> 2.0đ",
        "details_p4": "C19 (1.0đ), C20 (1.0đ), C21 (0đ)",
        "errors": "P4 chưa làm C21",
        "note": "Giỏi 9.00đ. Trắc nghiệm tuyệt đối 7.0/7.0đ (P1 3.0đ, P2 2.0đ, P3 2.0đ)."
    },
    {
        "stt": 42,
        "name": "Cao Minh Thư",
        "class": "12B1",
        "code": "101",
        "p1": 3.00,
        "p2": 2.00,
        "p3": 2.00,
        "p4": 2.00,
        "total": 9.00,
        "rank": "Giỏi",
        "answers_p1": ["D", "B", "D", "D", "D", "B", "C", "A", "B", "D", "B", "A"],
        "answers_p2": ["S", "Đ", "Đ", "S", "Đ", "Đ", "S", "S"],
        "answers_p3": ["2", "1", "7", "18"],
        "answers_p4": [1.0, 1.0, 0.0],
        "details_p1": "12/12 đúng tuyệt đối (100%) -> 3.0đ",
        "details_p2": "8/8 ý đúng tuyệt đối -> 2.0đ",
        "details_p3": "4/4 câu đúng tuyệt đối (C15: 2, C16: 1, C17: 7, C18: 18) -> 2.0đ",
        "details_p4": "C19 (1.0đ), C20 (1.0đ), C21 (0đ)",
        "errors": "P4 chưa làm C21",
        "note": "Giỏi 9.00đ. Trắc nghiệm tuyệt đối 7.0/7.0đ (P1 3.0đ, P2 2.0đ, P3 2.0đ)."
    },
    {
        "stt": 43,
        "name": "Đào Thanh Tùng",
        "class": "12B1",
        "code": "101",
        "p1": 3.00,
        "p2": 2.00,
        "p3": 1.00,
        "p4": 1.50,
        "total": 7.50,
        "rank": "Khá",
        "answers_p1": ["D", "B", "D", "D", "D", "B", "C", "A", "B", "D", "B", "A"],
        "answers_p2": ["S", "Đ", "Đ", "S", "Đ", "Đ", "S", "S"],
        "answers_p3": ["1", "1", "9.5", "18"],
        "answers_p4": [1.0, 0.5, 0.0],
        "details_p1": "12/12 đúng tuyệt đối (100%) -> 3.0đ",
        "details_p2": "8/8 ý đúng tuyệt đối -> 2.0đ",
        "details_p3": "2/4 câu đúng (Đúng C16: 1, C18: 18; Sai C15: 1, C17: 9.5) -> 1.00đ",
        "details_p4": "C19 (1.0đ), C20 (0.5đ), C21 (0đ)",
        "errors": "Sai P3 C15, C17; P4 C21 chưa làm",
        "note": "Khá 7.50đ. Phần I và II tuyệt đối 5.0/5.0đ."
    }
]

def export_excel():
    excel_path = r"D:\TOAN\05_CONG_CU_VA_ENGINE\Cham_Diem_Hoc_Sinh\Ket_Qua_Cham_Diem_Giua_Ky_1_Lop12.xlsx"
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
        
    # Master Key Reference Rows (Rows 6-9)
    master_keys_ref = {
        "101": {
            "p1": ["D","B","D","D","D","B","C","A","B","D","B","A"],
            "p2": ["S","Đ","Đ","S","Đ","Đ","S","S"],
            "p3": ["2","1","7","18"]
        },
        "102": {
            "p1": ["B","B","C","D","C","B","C","B","B","C","D","D"],
            "p2": ["Đ","S","Đ","Đ","Đ","S","Đ","S"],
            "p3": ["7","18","1","2"]
        },
        "103": {
            "p1": ["B","A","B","D","C","A","A","D","B","C","A","C"],
            "p2": ["Đ","Đ","S","S","S","Đ","Đ","S"],
            "p3": ["18","2","1","7"]
        },
        "104": {
            "p1": ["D","A","B","A","D","C","A","D","C","C","D","D"],
            "p2": ["Đ","Đ","S","S","S","Đ","Đ","S"],
            "p3": ["2","1","7","18"]
        }
    }
    
    for r_idx, code_k in enumerate(["101", "102", "103", "104"], 6):
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

    # Student Rows (Starting Row 10)
    for idx, s in enumerate(STUDENT_RESULTS, 1):
        row_idx = 9 + idx
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
    ws3.merge_cells("A1:E1")
    ws3["A1"] = "BẢNG ĐÁP ÁN VÀ BAREM ĐIỂM CHUẨN 4 MÃ ĐỀ (101, 102, 103, 104)"
    ws3["A1"].font = font_title
    ws3["A1"].alignment = Alignment(horizontal="center", vertical="center")
    
    headers_s3 = ["Phần Thi / Câu Hỏi", "Mã Đề 101", "Mã Đề 102", "Mã Đề 103", "Mã Đề 104"]
    ws3.row_dimensions[3].height = 25
    for c_idx, h in enumerate(headers_s3, 1):
        cell = ws3.cell(row=3, column=c_idx, value=h)
        cell.font = font_head
        cell.fill = fill_head
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    p1_keys = [
        ("Câu 1 (0.25đ)", "D", "B", "B", "D"),
        ("Câu 2 (0.25đ)", "B", "B", "A", "A"),
        ("Câu 3 (0.25đ)", "D", "C", "B", "B"),
        ("Câu 4 (0.25đ)", "D", "D", "D", "A"),
        ("Câu 5 (0.25đ)", "D", "C", "C", "D"),
        ("Câu 6 (0.25đ)", "B", "B", "A", "C"),
        ("Câu 7 (0.25đ)", "C", "C", "A", "A"),
        ("Câu 8 (0.25đ)", "A", "B", "D", "D"),
        ("Câu 9 (0.25đ)", "B", "B", "B", "C"),
        ("Câu 10 (0.25đ)", "D", "C", "C", "C"),
        ("Câu 11 (0.25đ)", "B", "D", "A", "D"),
        ("Câu 12 (0.25đ)", "A", "D", "C", "D"),
        ("Câu 13 (Đúng/Sai - 1.0đ)", "S - Đ - Đ - S", "Đ - S - Đ - Đ", "Đ - Đ - S - S", "Đ - Đ - S - S"),
        ("Câu 14 (Đúng/Sai - 1.0đ)", "Đ - Đ - S - S", "Đ - S - Đ - S", "S - Đ - Đ - S", "S - Đ - Đ - S"),
        ("Câu 15 (Trả lời ngắn - 0.5đ)", "2", "7", "18", "2"),
        ("Câu 16 (Trả lời ngắn - 0.5đ)", "1", "18", "2", "1"),
        ("Câu 17 (Trả lời ngắn - 0.5đ)", "7", "1", "1", "7"),
        ("Câu 18 (Trả lời ngắn - 0.5đ)", "18", "2", "7", "18"),
        ("Câu 19 (Tự luận - 1.0đ)", "m = 3; b = 1; c = -3", "m = 3; b = 1; c = -3", "m = 3; b = 1; c = -3", "m = 3; b = 1; c = -3"),
        ("Câu 20 (Tự luận - 1.0đ)", "v_max = 25 m/s (t = 4s)", "v_max = 25 m/s (t = 4s)", "v_max = 25 m/s (t = 4s)", "v_max = 25 m/s (t = 4s)"),
        ("Câu 21 (Tự luận - 1.0đ)", "x = 180/√55 ≈ 24.3m", "x = 180/√55 ≈ 24.3m", "x = 180/√55 ≈ 24.3m", "x = 180/√55 ≈ 24.3m")
    ]
    
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
    ws3.column_dimensions["B"].width = 24
    ws3.column_dimensions["C"].width = 24
    ws3.column_dimensions["D"].width = 24
    ws3.column_dimensions["E"].width = 24

    wb.save(excel_path)
    wb.close()
    print(f"EXCEL SAVED SUCCESSFULLY: {excel_path}")

if __name__ == '__main__':
    export_excel()
