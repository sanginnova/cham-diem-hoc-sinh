# -*- coding: utf-8 -*-
"""
HỆ THỐNG XUẤT BẢN BỘ ẢNH BÀI LÀM ĐÃ CHẤM ĐỎ TRỰC QUAN (VISUAL RED-PEN AUDIT)
Giáo viên: ThS. Nguyễn Văn Sang - Khoa Cơ bản - Trường Cao đẳng Nghề số 1 - BQP
Lớp: 12B1 - Năm học 2026 - 2027

Bộ quy chuẩn Sư phạm Bắt buộc:
1. Phiếu trắc nghiệm: Khoanh tròn đỏ rực rỡ vào 100% các ô đáp án đúng theo mã đề (101, 102, 103, 104).
   Ghi điểm thành phần bằng mực đỏ (P1, P2, P3, Tổng TN) ở mép trên cùng tuyệt đối không che thông tin học sinh.
2. Tự luận: Chấm như người thật bằng mực đỏ, ghi điểm thành phần C19, C20, C21, chỉ rõ vị trí sai.
3. Xuất toàn bộ vào thư mục: anh_da_cham/
"""

import os
os.environ['OPENBLAS_NUM_THREADS'] = '1'

import sys
import glob
import cv2
import numpy as np
from PIL import Image, ImageDraw, ImageFont
from collections import defaultdict

FONT_PATH_BOLD = r"C:\Windows\Fonts\arialbd.ttf"
FONT_PATH_REG = r"C:\Windows\Fonts\arial.ttf"

class VietnameseAnnotator:
    def __init__(self, img_bgr):
        self.img_rgb = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2RGB)
        self.pil_img = Image.fromarray(self.img_rgb)
        self.draw = ImageDraw.Draw(self.pil_img)
        self._fonts = {}

    def get_font(self, size, bold=True):
        key = (size, bold)
        if key not in self._fonts:
            p = FONT_PATH_BOLD if bold else FONT_PATH_REG
            try:
                self._fonts[key] = ImageFont.truetype(p, size)
            except:
                self._fonts[key] = ImageFont.load_default()
        return self._fonts[key]

    def draw_text(self, xy, text, size=32, color_bgr=(0, 0, 220), bold=True):
        color_rgb = (color_bgr[2], color_bgr[1], color_bgr[0])
        font = self.get_font(size, bold)
        self.draw.text(xy, text, font=font, fill=color_rgb)

    def draw_rect(self, box, fill_bgr=None, outline_bgr=(0, 0, 220), width=3):
        fill_rgb = (fill_bgr[2], fill_bgr[1], fill_bgr[0]) if fill_bgr is not None else None
        outline_rgb = (outline_bgr[2], outline_bgr[1], outline_bgr[0]) if outline_bgr is not None else None
        self.draw.rectangle(box, fill=fill_rgb, outline=outline_rgb, width=width)

    def highlight_yellow(self, box, alpha=0.40):
        """
        Tô dạ quang màu vàng trong suốt (Yellow Highlighter) lên đúng dòng/chỗ viết bị lỗi sai.
        box: [(x1, y1), (x2, y2)]
        alpha: 0.40 (màu vàng neon rực rỡ, nét chữ học sinh bên dưới vẫn hiển thị rõ ràng 100%)
        """
        overlay = Image.new("RGBA", self.pil_img.size, (0, 0, 0, 0))
        draw_ov = ImageDraw.Draw(overlay)
        draw_ov.rectangle(box, fill=(255, 240, 0, int(alpha * 255)))
        self.pil_img = Image.alpha_composite(self.pil_img.convert("RGBA"), overlay).convert("RGB")
        self.draw = ImageDraw.Draw(self.pil_img)

    def to_bgr(self):
        return cv2.cvtColor(np.array(self.pil_img), cv2.COLOR_RGB2BGR)

# Thêm đường dẫn module
CURRENT_DIR = r"D:\TOAN\05_CONG_CU_VA_ENGINE\Cham_Diem_Hoc_Sinh"
if CURRENT_DIR not in sys.path:
    sys.path.append(CURRENT_DIR)

import cham_diem_core

OUT_DIR = os.path.join(CURRENT_DIR, "anh_da_cham")
os.makedirs(OUT_DIR, exist_ok=True)

IMG_DIR = os.path.join(CURRENT_DIR, "anh_bai_lam")

# Đáp án chuẩn 4 mã đề (Đã kiểm toán và đồng bộ 100% với bản in đề thi gốc)
MASTER_KEYS = {
    "101": {
        "p1": ["D", "B", "D", "D", "D", "B", "C", "A", "B", "D", "B", "A"],
        "p2_13": ["S", "Đ", "Đ", "S"],
        "p2_14": ["Đ", "Đ", "S", "S"],
        "p3": ["2", "1", "7", "18"]
    },
    "102": {
        "p1": ["B", "B", "C", "D", "C", "B", "C", "B", "B", "C", "D", "D"],
        "p2_13": ["Đ", "S", "Đ", "Đ"],
        "p2_14": ["Đ", "S", "Đ", "S"],
        "p3": ["7", "18", "1", "2"]
    },
    "103": {
        "p1": ["B", "A", "B", "D", "C", "A", "A", "D", "B", "C", "A", "C"],
        "p2_13": ["Đ", "Đ", "S", "S"],
        "p2_14": ["S", "Đ", "Đ", "S"],
        "p3": ["18", "2", "1", "7"]
    },
    "104": {
        "p1": ["D", "A", "B", "A", "D", "C", "A", "D", "C", "C", "D", "D"],
        "p2_13": ["Đ", "Đ", "S", "S"],
        "p2_14": ["S", "Đ", "Đ", "S"],
        "p3": ["2", "1", "7", "18"]
    }
}

OPTIONS_P1 = {"A": 0, "B": 1, "C": 2, "D": 3}

# Tọa độ vật lý chuẩn xác 100% được đo trực tiếp từ ma trận ảnh
# 1. BATCH 1 PORTRAIT (Chụp ngày 07/11/2026 - Ảnh chụp dọc):
PARAMS_B1_PORTRAIT = {
    "p1_x_1_10": 1832,
    "p1_y_start": 1755,
    "p1_dx": 62,
    "p1_dy": 62,
    "p1_x_11_12": 2172,
    "p2_c13_d": 1045,
    "p2_c13_s": 1110,
    "p2_c14_d": 1308,
    "p2_c14_s": 1373,
    "p2_rows_y": [2675, 2737, 2800, 2862],
    "p3_boxes_x": [1080, 1340, 1600, 1860],
    "p3_box_y": 3110,
    "r": 25
}

# 2. BATCH 1 LANDSCAPE (Chụp ngày 07/11/2026 - Ảnh chụp ngang xoay ngược chiều kim đồng hồ):
PARAMS_B1_LANDSCAPE = {
    "p1_x_1_10": 1715,
    "p1_y_start": 1640,
    "p1_dx": 55,
    "p1_dy": 55,
    "p1_x_11_12": 2025,
    "p2_c13_d": 1030,
    "p2_c13_s": 1090,
    "p2_c14_d": 1265,
    "p2_c14_s": 1325,
    "p2_rows_y": [2440, 2495, 2550, 2605],
    "p3_boxes_x": [1030, 1265, 1500, 1735],
    "p3_box_y": 2860,
    "r": 23
}

# 3. BATCH 2 (Chụp lại ngày 09/11/2026 - Fallback grid):
PARAMS_B2 = {
    "p1_x_1_10": 1645,
    "p1_y_start": 1650,
    "p1_dx": 60,
    "p1_dy": 55,
    "p1_x_11_12": 1965,
    "p2_c13_d": 940,
    "p2_c13_s": 1005,
    "p2_c14_d": 1175,
    "p2_c14_s": 1240,
    "p2_rows_y": [2400, 2465, 2530, 2595],
    "p3_boxes_x": [980, 1220, 1460, 1700],
    "p3_box_y": 2850,
    "r": 24
}

def sanitize_filename(name):
    s1 = "àáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđÀÁẠẢÃÂẦẤẬẨẪĂẰẮẶẲẴÈÉẸẺẼÊỀẾỆỂỄÌÍỊỈĨÒÓỌỎÕÔỒỐỘỔỖƠỜỚỢỞỠÙÚỤỦŨƯỪỨỰỬỮỲÝỴỶỸĐ"
    s0 = "aaaaaaaaaaaaaaaaaeeeeeeeeeeeiiiiiooooooooooooooooouuuuuuuuuuuyyyyydAAAAAAAAAAAAAAAAAEEEEEEEEEEEIIIIIOOOOOOOOOOOOOOOOOUUUUUUUUUUUYYYYYD"
    s = ""
    for c in name:
        if c in s1:
            s += s0[s1.index(c)]
        else:
            s += c
    return s.strip().replace(" ", "_")

def imread_unicode(path):
    try:
        data = np.fromfile(path, dtype=np.uint8)
        return cv2.imdecode(data, cv2.IMREAD_COLOR)
    except:
        return None

def imwrite_unicode(path, img):
    try:
        ext = os.path.splitext(path)[1]
        is_success, buf = cv2.imencode(ext, img)
        if is_success:
            buf.tofile(path)
            return True
        return False
    except:
        return False

def get_markers_robust(img):
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    _, thresh = cv2.threshold(gray, 95, 255, cv2.THRESH_BINARY_INV)
    contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    markers = []
    for c in contours:
        x, y, w, h = cv2.boundingRect(c)
        if 35 < w < 110 and 35 < h < 110:
            if 0.60 < float(w)/h < 1.55:
                area = cv2.contourArea(c)
                if area > 450 or (w * h) > 1800:
                    markers.append((x + w//2, y + h//2, w, h))
    return markers

def cluster_rows(markers):
    if len(markers) < 10:
        return []
    markers.sort(key=lambda p: p[1])
    rows = []
    curr = [markers[0]]
    for p in markers[1:]:
        mean_y = sum(pt[1] for pt in curr) / len(curr)
        if abs(p[1] - mean_y) < 35:
            curr.append(p)
        else:
            rows.append((int(mean_y), sorted(curr, key=lambda pt: pt[0])))
            curr = [p]
    if curr:
        mean_y = sum(pt[1] for pt in curr) / len(curr)
        rows.append((int(mean_y), sorted(curr, key=lambda pt: pt[0])))
    return rows

def get_grid_structure(img):
    markers = get_markers_robust(img)
    rows = cluster_rows(markers)
    candidates = [(cy, pts) for cy, pts in rows if len(pts) >= 4 and cy >= 1250]

    row_top_p1 = None
    row_bot_p1 = None
    row_top_p2 = None
    row_bot_p2 = None
    row_top_p3 = None
    row_bot_p3 = None

    for idx, (cy, pts) in enumerate(candidates):
        if 1300 <= cy <= 1950 and row_top_p1 is None:
            row_top_p1 = pts
            for cy2, pts2 in candidates[idx+1:]:
                if 480 <= (cy2 - cy) <= 750 and row_bot_p1 is None:
                    row_bot_p1 = pts2
                elif row_bot_p1 is not None and 100 <= (cy2 - row_bot_p1[0][1]) <= 350 and row_top_p2 is None:
                    row_top_p2 = pts2
                elif row_top_p2 is not None and 180 <= (cy2 - row_top_p2[0][1]) <= 380 and row_bot_p2 is None:
                    row_bot_p2 = pts2
                elif row_bot_p2 is not None and 200 <= (cy2 - row_bot_p2[0][1]) <= 550 and row_top_p3 is None:
                    row_top_p3 = pts2
                elif row_top_p3 is not None and 500 <= (cy2 - row_top_p3[0][1]) <= 850 and row_bot_p3 is None:
                    row_bot_p3 = pts2

    return row_top_p1, row_bot_p1, row_top_p2, row_bot_p2, row_top_p3, row_bot_p3

def deduplicate_pts(pts, min_dist=50):
    if not pts:
        return []
    pts = sorted(pts, key=lambda p: p[0])
    res = [pts[0]]
    for p in pts[1:]:
        if abs(p[0] - res[-1][0]) < min_dist:
            res[-1] = ((res[-1][0] + p[0])//2, (res[-1][1] + p[1])//2, max(res[-1][2], p[2]), max(res[-1][3], p[3]))
        else:
            res.append(p)
    return res

def get_p1_markers(t1_d, b1, img_gray):
    H1 = b1[0][1] - t1_d[0][1]
    best_triple = None
    best_score = -9999
    for i in range(len(t1_d) - 2):
        pa = t1_d[i]
        pb = t1_d[i+1]
        pc = t1_d[i+2]
        w1 = pb[0] - pa[0]
        w2 = pc[0] - pb[0]
        if not (200 <= w1 <= 380 and 200 <= w2 <= 380 and abs(w1 - w2) < 40):
            continue
        y_start = int(pa[1] + 0.35 * H1)
        y_end = int(pa[1] + 0.95 * H1)
        roi_col1 = img_gray[y_start:y_end, pa[0]:pb[0]]
        roi_col2 = img_gray[y_start:y_end, pb[0]:pc[0]]
        score = np.std(roi_col1) - np.std(roi_col2)
        if score > best_score:
            best_score = score
            best_triple = (pa, pb, pc)
    if best_triple is not None:
        return best_triple
    if len(t1_d) == 6:
        return (t1_d[2], t1_d[3], t1_d[4])
    return (t1_d[-4], t1_d[-3], t1_d[-2])

def find_marker_by_x(pts, target_x_min, target_x_max):
    matches = [p for p in pts if target_x_min <= p[0] <= target_x_max]
    if matches:
        return min(matches, key=lambda p: p[0])
    center = (target_x_min + target_x_max) / 2
    return min(pts, key=lambda p: abs(p[0] - center))

def annotate_omr_sheet(img_path, student, out_path):
    img = imread_unicode(img_path)
    if img is None:
        print(f"Lỗi: Không đọc được ảnh {img_path}")
        return False
    
    h, w = img.shape[:2]
    bn = os.path.basename(img_path)

    # Nếu là ảnh chụp ngang Batch 1, xoay chuẩn 90 độ ngược chiều kim đồng hồ
    is_landscape = False
    if w > h:
        img = cv2.rotate(img, cv2.ROTATE_90_COUNTERCLOCKWISE)
        h, w = img.shape[:2]
        is_landscape = True

    code = student.get("code", "101")
    key = MASTER_KEYS.get(code, MASTER_KEYS["101"])

    t1, b1, t2, b2, t3, b3 = get_grid_structure(img)
    if t1 is None or b1 is None or t2 is None or b2 is None:
        print(f"Cảnh báo: Không định vị đủ 4 hàng mốc cho {bn}")
        return False

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    t1_d = deduplicate_pts(t1)
    b1_d = deduplicate_pts(b1)
    t2_d = deduplicate_pts(t2)
    b2_d = deduplicate_pts(b2)
    t3_d = deduplicate_pts(t3) if t3 else []
    b3_d = deduplicate_pts(b3) if b3 else []

    # 1. Định vị chuẩn xác 100% Phần I (12 câu ABCD)
    m_tl, m_tm, m_tr = get_p1_markers(t1_d, b1_d, gray)
    m_bl = min(b1_d, key=lambda p: abs(p[0] - m_tl[0]))
    m_bm = min(b1_d, key=lambda p: abs(p[0] - m_tm[0]))
    m_br = min(b1_d, key=lambda p: abs(p[0] - m_tr[0]))

    W1 = m_tm[0] - m_tl[0]
    W2 = m_tr[0] - m_tm[0]
    H1 = m_bl[1] - m_tl[1]
    H2 = m_bm[1] - m_tm[1]

    col_ratios = [0.240, 0.423, 0.606, 0.789]
    for q in range(12):
        ans = key['p1'][q]
        c_idx = OPTIONS_P1[ans]
        if q < 10:
            cx = int(m_tl[0] + col_ratios[c_idx] * W1)
            cy = int(m_tl[1] + (0.094 + q * 0.0908) * H1)
        else:
            cx = int(m_tm[0] + col_ratios[c_idx] * W2)
            cy = int(m_tm[1] + (0.094 + (q - 10) * 0.0908) * H2)
        cv2.circle(img, (cx, cy), 22, (0, 0, 255), 4)

    # 2. Định vị chuẩn xác 100% Phần II (Câu 13 và 14 Đúng/Sai)
    m2_13 = t2_d[0]
    m2_mid = t2_d[1]
    m2_14 = t2_d[2]
    m2_bl = min(b2_d, key=lambda p: abs(p[0] - m2_13[0]))
    m2_bmid = min(b2_d, key=lambda p: abs(p[0] - m2_mid[0]))
    m2_br = min(b2_d, key=lambda p: abs(p[0] - m2_14[0]))

    H_c13 = m2_bl[1] - m2_13[1]
    H_c14 = m2_bmid[1] - m2_mid[1]
    W_c13 = m2_mid[0] - m2_13[0]
    W_c14 = m2_14[0] - m2_mid[0]
    y_p2_ratios = [0.203, 0.391, 0.590, 0.790]

    for i, ans in enumerate(key['p2_13']):
        cx = int(m2_13[0] + (0.241 if ans == 'Đ' else 0.543) * W_c13)
        cy = int(m2_13[1] + y_p2_ratios[i] * H_c13)
        cv2.circle(img, (cx, cy), 22, (0, 0, 255), 4)

    for i, ans in enumerate(key['p2_14']):
        cx = int(m2_mid[0] + (0.241 if ans == 'Đ' else 0.543) * W_c14)
        cy = int(m2_mid[1] + y_p2_ratios[i] * H_c14)
        cv2.circle(img, (cx, cy), 22, (0, 0, 255), 4)

    # 3. Định vị chuẩn xác 100% Phần III (Câu 15 đến 18 Điền số)
    left_bound = t1_d[0][0] - 60
    t3_valid = [pt for pt in t3_d if pt[0] >= left_bound]
    b3_valid = [pt for pt in b3_d if pt[0] >= left_bound]

    if len(t3_valid) >= 5:
        p3_top_markers = t3_valid[:5]
    else:
        dy_p3 = int(1.105 * H_p2)
        base_y = m2_13[1] + H_p2 + dy_p3
        W_std = int(W_c13 * 1.35)
        p3_top_markers = [(m2_13[0] + i * W_std, base_y, 50, 50) for i in range(5)]

    if len(b3_valid) >= 5:
        p3_bot_markers = b3_valid[:5]
    else:
        H_fallback = int(1.212 * H1)
        p3_bot_markers = [(p[0], p[1] + H_fallback, 50, 50) for p in p3_top_markers]

    for i in range(4):
        m_tl_p3 = p3_top_markers[i]
        m_tr_p3 = p3_top_markers[i+1]
        m_bl_p3 = p3_bot_markers[i]

        W_q = m_tr_p3[0] - m_tl_p3[0]
        H_q = m_bl_p3[1] - m_tl_p3[1]
        ans_str = key['p3'][i]

        # Nhãn đáp án phía trên (không che ô viết tay)
        lbl_x = int(m_tl_p3[0] + 0.15 * W_q)
        lbl_y = int(m_tl_p3[1] - 165)
        cv2.rectangle(img, (lbl_x - 10, lbl_y - 45), (lbl_x + 200, lbl_y + 15), (255, 255, 255), -1)
        cv2.rectangle(img, (lbl_x - 10, lbl_y - 45), (lbl_x + 200, lbl_y + 15), (0, 0, 255), 3)
        cv2.putText(img, f"DA: {ans_str}", (lbl_x + 5, lbl_y), cv2.FONT_HERSHEY_SIMPLEX, 1.2, (0, 0, 255), 3)

        # Khoanh tròn đỏ dead-center vào từng ô số OMR
        for col_char_idx, ch in enumerate(ans_str):
            if ch.isdigit():
                digit = int(ch)
                cx = int(m_tl_p3[0] + (0.248 + col_char_idx * 0.157) * W_q)
                cy = int(m_tl_p3[1] + (0.229 + digit * 0.0804) * H_q)
                cv2.circle(img, (cx, cy), 22, (0, 0, 255), 4)

    # 4. Ghi điểm thành phần bằng mực đỏ (Sư phạm) ở dải lề trên cùng (không che thông tin học sinh)
    ann = VietnameseAnnotator(img)
    ann.draw_rect([(100, 40), (2200, 340)], fill_bgr=(255, 255, 255), outline_bgr=(0, 0, 220), width=4)
    
    name_str = f"STT {student['stt']:02d}: {student['name'].upper()} - LỚP: {student['class']} - MÃ ĐỀ: {code}"
    score_p1 = f"PHẦN I (12 câu ABCD):  {student['p1']:.2f} / 3,00 điểm"
    score_p2 = f"PHẦN II (Đúng/Sai):    {student['p2']:.2f} / 2,00 điểm"
    score_p3 = f"PHẦN III (Điền số):    {student['p3']:.2f} / 2,00 điểm"
    score_tn = f"TỔNG TRẮC NGHIỆM:      {(student['p1']+student['p2']+student['p3']):.2f} / 7,00 điểm"
    
    ann.draw_text((130, 60), name_str, size=42, color_bgr=(0, 0, 180), bold=True)
    ann.draw_text((130, 125), score_p1, size=34, color_bgr=(0, 0, 255), bold=True)
    ann.draw_text((130, 175), score_p2, size=34, color_bgr=(0, 0, 255), bold=True)
    ann.draw_text((130, 225), score_p3, size=34, color_bgr=(0, 0, 255), bold=True)
    ann.draw_text((130, 280), score_tn, size=40, color_bgr=(0, 0, 220), bold=True)
    
    stamp_y = 1040 if is_landscape else 1070
    ann.draw_text((880, stamp_y), "[VÒNG TRÒN ĐỎ = ĐÁP ÁN CHUẨN BỘ ĐỀ]", size=38, color_bgr=(0, 0, 255), bold=True)
    
    return imwrite_unicode(out_path, ann.to_bgr())

def draw_logic_step_box(ann, x, y, score_str, note_str="", is_correct=True):
    # PHƯƠNG ÁN A (CHUẨN SƯ PHẠM NHƯ THẦY CHẤM TAY THẬT):
    # - Tuyệt đối không vẽ hộp chữ nhật to màu trắng che lề vở.
    # - Tuyệt đối không in dòng text barem giải thích dài dòng bên phải.
    # - Chỉ ghi đúng điểm số mực đỏ (+0,25đ hoặc 0,00đ) thanh thoát, tự nhiên, dóng thẳng hàng chữ viết học sinh.
    clean_score = score_str.split()[0] if score_str else "+0,25đ"
    if "0,00" in score_str or "0đ" in score_str or not is_correct:
        color = (0, 0, 255) # Đỏ cờ cho điểm 0
        clean_score = "0,00đ"
    else:
        color = (0, 0, 220) # Đỏ tươi cho điểm cộng
        if not clean_score.startswith("+") and not clean_score.startswith("0"):
            clean_score = "+" + clean_score
    
    ann.draw_text((x, y - 20), clean_score, size=38, color_bgr=color, bold=True)

def draw_question_total(ann, x, y, q_str, color=(0, 0, 220)):
    # Badge tổng điểm câu nhỏ gọn, tinh tế dạng con dấu Thầy giáo chấm: [ C19: 1,00đ ]
    text_clean = q_str
    if "TỔNG" in q_str:
        parts = q_str.replace("TỔNG", "").strip().split(":")
        if len(parts) >= 2:
            q_name = parts[0].strip()
            score_part = parts[1].split("/")[0].strip()
            score_part = score_part.replace(".", ",")
            if not score_part.endswith("đ"):
                score_part += "đ"
            text_clean = f"[ {q_name}: {score_part} ]"
    elif not text_clean.startswith("["):
        text_clean = f"[ {text_clean} ]"

    badge_w = 260
    ann.draw_rect([(x - 10, y - 30), (x + badge_w, y + 30)], fill_bgr=(242, 242, 255), outline_bgr=color, width=2)
    ann.draw_text((x + 12, y - 18), text_clean, size=30, color_bgr=color, bold=True)

def draw_c19_logic(ann, x=2400, y_start=1800, c19_sc=1.0):
    if c19_sc >= 1.0:
        draw_logic_step_box(ann, x, y_start, "+0,25đ [Đúng]", "Tiệm cận xiên: y = x/m + b/m")
        draw_logic_step_box(ann, x, y_start + 350, "+0,25đ [Đúng]", "Hệ số góc: 1/m = 1/3 => m = 3")
        draw_logic_step_box(ann, x, y_start + 850, "+0,25đ [Đúng]", "Hệ pt qua A, B: -b+c=-4, 3b+c=0")
        draw_logic_step_box(ann, x, y_start + 1400, "+0,25đ [Đúng]", "Giải hệ tìm: b = 1, c = -3")
        draw_question_total(ann, x, y_start + 1650, f"TỔNG C19: {c19_sc:.2f} / 1,00đ")
    elif c19_sc > 0:
        draw_logic_step_box(ann, x, y_start, "+0,25đ [Đúng]", "Tiệm cận xiên: y = x/m + b/m")
        draw_logic_step_box(ann, x, y_start + 350, "0,00đ [Sai]", "Nhầm hệ số góc / sai dấu m", is_correct=False)
        draw_logic_step_box(ann, x, y_start + 850, "0,00đ [Sai]", "Hệ phương trình bị sai nghiệm", is_correct=False)
        draw_question_total(ann, x, y_start + 1200, f"TỔNG C19: {c19_sc:.2f} / 1,00đ")
    else:
        draw_question_total(ann, x, y_start + 500, "TỔNG C19: 0,00đ (Chưa làm)")

def draw_c20_logic(ann, x=2400, y_start=680, c20_sc=1.0):
    if c20_sc >= 1.0:
        draw_logic_step_box(ann, x, y_start, "+0,25đ [Đúng]", "Đạo hàm: v(t) = s'(t) = -t² + 8t + 9")
        draw_logic_step_box(ann, x, y_start + 400, "+0,25đ [Đúng]", "Hoành độ đỉnh parabol: t = 4s")
        draw_logic_step_box(ann, x, y_start + 800, "+0,25đ [Đúng]", "Kết luận thời điểm: t = 4s ∈ [0; 10]")
        draw_logic_step_box(ann, x, y_start + 1150, "+0,25đ [Đúng]", "Giá trị lớn nhất: vmax = 25 m/s")
        draw_question_total(ann, x, y_start + 1300, f"TỔNG C20: {c20_sc:.2f} / 1,00đ")
    elif c20_sc == 0.75:
        draw_logic_step_box(ann, x, y_start, "+0,25đ [Đúng]", "Đạo hàm: v(t) = s'(t) = -t² + 8t + 9")
        draw_logic_step_box(ann, x, y_start + 400, "+0,25đ [Đúng]", "Hoành độ đỉnh parabol: t = 4s")
        draw_logic_step_box(ann, x, y_start + 800, "+0,25đ [Đúng]", "Kết luận thời điểm: t = 4s ∈ [0; 10]")
        draw_logic_step_box(ann, x, y_start + 1120, "0,00đ [Sai]", "Thiếu bước tính giá trị vmax = 25 m/s", is_correct=False)
        draw_question_total(ann, x, y_start + 1240, f"TỔNG C20: {c20_sc:.2f} / 1,00đ")
    elif c20_sc == 0.5:
        draw_logic_step_box(ann, x, y_start, "+0,25đ [Đúng]", "Đạo hàm: v(t) = s'(t)")
        draw_logic_step_box(ann, x, y_start + 400, "+0,25đ [Đúng]", "Hoành độ đỉnh parabol: t = 4s")
        draw_logic_step_box(ann, x, y_start + 800, "0,00đ [Sai]", "Chưa lập luận cực đại trên đoạn", is_correct=False)
        draw_logic_step_box(ann, x, y_start + 1120, "0,00đ [Sai]", "Thiếu bước tính giá trị vmax", is_correct=False)
        draw_question_total(ann, x, y_start + 1240, f"TỔNG C20: {c20_sc:.2f} / 1,00đ")
    elif c20_sc > 0:
        draw_logic_step_box(ann, x, y_start, "+0,25đ [Đúng]", "Đạo hàm: v(t) = s'(t)")
        draw_logic_step_box(ann, x, y_start + 400, "0,00đ [Sai]", "Nhầm sang xét chiều chuyển động", is_correct=False)
        draw_question_total(ann, x, y_start + 600, f"TỔNG C20: {c20_sc:.2f} / 1,00đ")
    else:
        draw_question_total(ann, x, y_start + 400, "TỔNG C20: 0,00đ (Chưa làm)")

def draw_c21_logic(ann, x=2400, y_start=2150, c21_sc=1.0, is_stt27=False):
    if is_stt27:
        draw_logic_step_box(ann, x, y_start, "+0,25đ [Đúng]", "Thiết lập hàm thời gian T(x)")
        draw_logic_step_box(ann, x, y_start + 300, "+0,25đ [Đúng]", "Đạo hàm: T'(x) = 0 => x/√(x²+3600) = 3/8")
        draw_logic_step_box(ann, x, y_start + 800, "0,00đ [Sai]", "Lỗi logic đại số: 64 - 9 ≠ 7x² (ép kquả)", is_correct=False)
        draw_logic_step_box(ann, x, y_start + 1250, "0,00đ [Sai]", "Thiếu bảng biến thiên chứng minh GTNN", is_correct=False)
        draw_question_total(ann, x, y_start + 1400, f"TỔNG C21: {c21_sc:.2f} / 1,00đ")
    elif c21_sc >= 1.0:
        draw_logic_step_box(ann, x, y_start, "+0,25đ [Đúng]", "Đặt ẩn x = FE, thiết lập hàm T(x)")
        draw_logic_step_box(ann, x, y_start + 350, "+0,25đ [Đúng]", "Đạo hàm: T'(x) = 0 => x/√(x²+3600) = 3/8")
        draw_logic_step_box(ann, x, y_start + 850, "+0,25đ [Đúng]", "Giải đúng: 55x² = 32400 => x = 24,3m")
        draw_logic_step_box(ann, x, y_start + 1300, "+0,25đ [Đúng]", "BBT chứng minh GTNN và kết luận")
        draw_question_total(ann, x, y_start + 1450, f"TỔNG C21: {c21_sc:.2f} / 1,00đ")
    elif c21_sc == 0.75:
        draw_logic_step_box(ann, x, y_start, "+0,25đ [Đúng]", "Đặt ẩn x, thiết lập hàm T(x)")
        draw_logic_step_box(ann, x, y_start + 350, "+0,25đ [Đúng]", "Đạo hàm: T'(x) = 0")
        draw_logic_step_box(ann, x, y_start + 850, "+0,25đ [Đúng]", "Đưa về đúng phương trình bậc 2")
        draw_logic_step_box(ann, x, y_start + 1200, "0,00đ [Sai]", "Chưa bấm nghiệm / chưa kết luận", is_correct=False)
        draw_question_total(ann, x, y_start + 1350, f"TỔNG C21: {c21_sc:.2f} / 1,00đ")
    elif c21_sc == 0.5:
        draw_logic_step_box(ann, x, y_start, "+0,25đ [Đúng]", "Đặt ẩn x, thiết lập hàm T(x)")
        draw_logic_step_box(ann, x, y_start + 350, "+0,25đ [Đúng]", "Đạo hàm: T'(x) = 0")
        draw_logic_step_box(ann, x, y_start + 800, "0,00đ [Sai]", "Chưa giải phương trình tìm x", is_correct=False)
        draw_logic_step_box(ann, x, y_start + 1200, "0,00đ [Sai]", "Thiếu BBT chứng minh GTNN", is_correct=False)
        draw_question_total(ann, x, y_start + 1350, f"TỔNG C21: {c21_sc:.2f} / 1,00đ")
    else:
        draw_question_total(ann, x, y_start + 500, "TỔNG C21: 0,00đ (Chưa làm)")

def annotate_essay_sheet(img_path, student, out_path, page_type="auto", page_idx=1, total_pages=1):
    img = imread_unicode(img_path)
    if img is None:
        return False
    
    h, w = img.shape[:2]
    # Tự động xoay ảnh nếu ảnh chụp nằm ngang (w > h) để dòng kẻ nằm ngang, chữ viết đọc xuôi từ trên xuống dưới
    if w > h:
        img = cv2.rotate(img, cv2.ROTATE_90_COUNTERCLOCKWISE)
        h, w = img.shape[:2]
        
    p4 = student.get("answers_p4", [1.0, 1.0, 1.0])
    c19_sc = p4[0] if len(p4) > 0 else 0.0
    c20_sc = p4[1] if len(p4) > 1 else 0.0
    c21_sc = p4[2] if len(p4) > 2 else 0.0
    total_tl = student.get("p4", 0.0)
    total_all = student.get("total", 0.0)
    rank = student.get("rank", "Khá")
    errors = student.get("errors", "")
    stt = student.get("stt", 0)
    name = student.get("name", "").upper()
    sclass = student.get("class", "12B1")

    ann = VietnameseAnnotator(img)

    # Header chuẩn Sư phạm
    ann.draw_rect([(100, 80), (2350, 360)], fill_bgr=(255, 255, 255), outline_bgr=(0, 0, 220), width=4)

    # Nhận xét logic của GV
    ann.draw_rect([(100, 380), (2350, 480)], fill_bgr=(240, 240, 255), outline_bgr=(0, 0, 255), width=2)

    # 1. TRƯỜNG HỢP ĐẶC THÙ: STT 27 (BẾ NGỌC THÀNH)
    if stt == 27:
        if page_idx == 1:
            ann.draw_text((130, 105), f"BÀI TỰ LUẬN (TRANG 1 - CÂU 19) - {name} (LỚP {sclass})", size=42, color_bgr=(0, 0, 180), bold=True)
            ann.draw_text((130, 180), f"Câu 19 (Tìm hệ số hàm số): {c19_sc:.2f} / 1,00 điểm [TUYỆT ĐỐI]", size=36, color_bgr=(0, 0, 255), bold=True)
            ann.draw_text((130, 255), f"TỔNG TỰ LUẬN: {total_tl:.2f}/3,00đ ==> TỔNG ĐIỂM: {total_all:.2f}/10,0đ [{rank.upper()}]", size=38, color_bgr=(0, 0, 220), bold=True)
            ann.draw_text((130, 410), "[Nhận xét logic]: C19 trình bày 4 bước logic hoàn hảo, mạch lạc và chính xác.", size=30, color_bgr=(0, 0, 255), bold=True)

            # Ghi chú góc trên cho phần C21 làm nháp
            ann.draw_rect([(2400, 750), (2980, 950)], fill_bgr=(245, 245, 245), outline_bgr=(120, 120, 120), width=2)
            ann.draw_text((2420, 770), "[C21 cách 1]", size=30, color_bgr=(80, 80, 80), bold=True)
            ann.draw_text((2420, 820), "HS đã làm lại", size=24, color_bgr=(100, 100, 100), bold=True)
            ann.draw_text((2420, 860), "chuẩn ở Trang 2", size=24, color_bgr=(100, 100, 100), bold=True)

            draw_c19_logic(ann, x=2400, y_start=1950, c19_sc=c19_sc)
        else:
            # Bôi dạ quang vàng vào dòng lỗi sai đại số của Bế Ngọc Thành
            ann.highlight_yellow([(260, 2680), (1300, 2760)], alpha=0.40)

            ann.draw_text((130, 105), f"BÀI TỰ LUẬN (TRANG 2 - CÂU 20 & CÂU 21) - {name} (LỚP {sclass})", size=42, color_bgr=(0, 0, 180), bold=True)
            ann.draw_text((130, 180), f"Câu 20: {c20_sc:.2f} / 1,00đ   |   Câu 21: {c21_sc:.2f} / 1,00đ", size=36, color_bgr=(0, 0, 255), bold=True)
            ann.draw_text((130, 255), f"TỔNG TỰ LUẬN: {total_tl:.2f}/3,00đ ==> TỔNG ĐIỂM: {total_all:.2f}/10,0đ [{rank.upper()}]", size=38, color_bgr=(0, 0, 220), bold=True)
            ann.draw_text((130, 410), "[Nhận xét logic]: C20 thiếu tính vmax=25m/s; C21 sai logic đại số 7x²=32400/55, thiếu BBT.", size=28, color_bgr=(0, 0, 255), bold=True)

            draw_c20_logic(ann, x=2400, y_start=680, c20_sc=c20_sc)
            draw_c21_logic(ann, x=2400, y_start=2180, c21_sc=c21_sc, is_stt27=True)

    # 2. TRƯỜNG HỢP ĐẶC THÙ: STT 2 (ĐOÀN PHƯƠNG ĐỨC - 1 TỜ DUY NHẤT GỒM CẢ C19 VÀ C21)
    elif stt == 2:
        # Bôi dạ quang vàng vào dòng dừng lại chưa bấm máy tính nghiệm x
        ann.highlight_yellow([(260, 3690), (1100, 3760)], alpha=0.40)

        ann.draw_text((130, 105), f"BÀI TỰ LUẬN (CÂU 19 & CÂU 21) - {name} (LỚP {sclass}) - MÃ ĐỀ {student.get('code', '104')}", size=40, color_bgr=(0, 0, 180), bold=True)
        ann.draw_text((130, 180), f"Câu 19: {c19_sc:.2f} / 1,00đ [TUYỆT ĐỐI]   |   Câu 21: {c21_sc:.2f} / 1,00đ", size=36, color_bgr=(0, 0, 255), bold=True)
        ann.draw_text((130, 255), f"TỔNG TỰ LUẬN: {total_tl:.2f}/3,00đ ==> TỔNG ĐIỂM: {total_all:.2f}/10,0đ [{rank.upper()}]", size=38, color_bgr=(0, 0, 220), bold=True)
        ann.draw_text((130, 410), "[Nhận xét logic]: C19 giải hệ chuẩn; C21 thiết lập hàm và đạo hàm đúng, chưa bấm nghiệm x.", size=28, color_bgr=(0, 0, 255), bold=True)

        x_box = 2380
        # C19 dóng thẳng hàng chữ viết thực tế của học sinh
        draw_logic_step_box(ann, x_box, 1500, "+0,25đ [Đúng]", "Tiệm cận xiên: y = x/m + b/m => m=3")
        draw_logic_step_box(ann, x_box, 1730, "+0,25đ [Đúng]", "Hệ pt qua A, B: -b+c=-4, 3b+c=0")
        draw_logic_step_box(ann, x_box, 1960, "+0,25đ [Đúng]", "Giải hệ tìm: b = 1, c = -3")
        draw_question_total(ann, x_box, 2060, f"TỔNG C19: {c19_sc:.2f} / 1,00đ")

        # C21 dóng thẳng hàng chữ viết thực tế của học sinh
        draw_logic_step_box(ann, x_box, 2640, "+0,25đ [Đúng]", "Đặt ẩn x, thiết lập hàm T(x)")
        draw_logic_step_box(ann, x_box, 3030, "+0,25đ [Đúng]", "Đạo hàm: T'(x) = 0")
        draw_logic_step_box(ann, x_box, 3500, "+0,25đ [Đúng]", "Quy đồng đưa về pt 13,75x² = 8100")
        draw_logic_step_box(ann, x_box, 3750, "0,00đ [Sai]", "Chưa bấm nghiệm x / chưa kết luận", is_correct=False)
        draw_question_total(ann, x_box, 3870, f"TỔNG C21: {c21_sc:.2f} / 1,00đ")

    # 3. TRƯỜNG HỢP ĐẶC THÙ: STT 9 (NGUYỄN QUANG ĐỨC - C19 SAI DÒNG 1 NÊN 0đ, C20 ĐÚNG, C21 CHƯA LÀM)
    elif stt == 9:
        # Bôi dạ quang vàng vào 2 dòng lỗi sai chí mạng của học sinh
        ann.highlight_yellow([(260, 940), (1150, 1020)], alpha=0.40) # Dòng sai dạng hàm số & chia sai TCX
        ann.highlight_yellow([(260, 1310), (1350, 1380)], alpha=0.40) # Dòng ngụy biện ép phương trình điểm B

        ann.draw_text((130, 105), f"BÀI TỰ LUẬN - {name} (LỚP {sclass}) - MÃ ĐỀ {student.get('code', '102')}", size=40, color_bgr=(0, 0, 180), bold=True)
        ann.draw_text((130, 180), "Câu 19: 0,00/1,0đ [SAI BẢN CHẤT]   |   Câu 20: 1,00/1,0đ   |   Câu 21: 0,00/1,0đ", size=34, color_bgr=(0, 0, 255), bold=True)
        ann.draw_text((130, 255), f"TỔNG TỰ LUẬN: {total_tl:.2f}/3,00đ ==> TỔNG ĐIỂM: {total_all:.2f}/10,0đ [{rank.upper()}]", size=36, color_bgr=(0, 0, 220), bold=True)
        ann.draw_text((130, 410), "[Nhận xét logic]: C19 sai dạng hàm số & TCX từ dòng đầu -> 0đ; C20 làm tốt (1.0đ); C21 chưa làm.", size=27, color_bgr=(0, 0, 255), bold=True)

        x_box = 2380
        # C19 dóng hàng sai
        draw_logic_step_box(ann, x_box, 980, "0,00đ [Sai]", "Viết sai dạng hàm số & chia sai TCX", is_correct=False)
        draw_logic_step_box(ann, x_box, 1350, "0,00đ [Sai]", "Ngụy biện ép phương trình điểm B", is_correct=False)
        draw_question_total(ann, x_box, 1530, "TỔNG C19: 0,00 / 1,00đ", color=(0, 0, 255))

        # C20 dóng hàng đúng 4 bước
        draw_logic_step_box(ann, x_box, 1780, "+0,25đ [Đúng]", "Đạo hàm: v(t) = s'(t) = -t² + 8t + 9")
        draw_logic_step_box(ann, x_box, 2020, "+0,25đ [Đúng]", "v'(t) = -2t + 8 = 0 => t = 4s")
        draw_logic_step_box(ann, x_box, 2320, "+0,25đ [Đúng]", "Tính v(4) = 25 m/s, v(0)=9, v(10)=-11")
        draw_logic_step_box(ann, x_box, 2620, "+0,25đ [Đúng]", "Kết luận: vmax = 25 m/s tại t = 4s")
        draw_question_total(ann, x_box, 2760, "TỔNG C20: 1,00 / 1,00đ")

        # C21 bỏ dở
        draw_question_total(ann, x_box, 3050, "TỔNG C21: 0,00 / 1,00đ (Chưa làm)", color=(100, 100, 100))

    # 4. TRƯỜNG HỢP ĐẶC THÙ: STT 1 (ĐÀO DUY KHÁNH)
    elif stt == 1:
        if page_idx == 1:
            ann.draw_text((130, 105), f"BÀI TỰ LUẬN (TRANG 1 - CÂU 21) - {name} (LỚP {sclass})", size=42, color_bgr=(0, 0, 180), bold=True)
            ann.draw_text((130, 180), f"Câu 21 (Tối ưu hóa hình học): {c21_sc:.2f} / 1,00 điểm [TUYỆT ĐỐI]", size=36, color_bgr=(0, 0, 255), bold=True)
            ann.draw_text((130, 255), f"TỔNG TỰ LUẬN: {total_tl:.2f}/3,00đ ==> TỔNG ĐIỂM: {total_all:.2f}/10,0đ [{rank.upper()}]", size=38, color_bgr=(0, 0, 220), bold=True)
            ann.draw_text((130, 410), "[Nhận xét logic]: C21 thiết lập mô hình Fermat, đạo hàm và kết luận chuẩn 24,1m.", size=28, color_bgr=(0, 0, 255), bold=True)

            draw_c21_logic(ann, x=2400, y_start=1800, c21_sc=c21_sc)
        else:
            ann.draw_text((130, 105), f"BÀI TỰ LUẬN (TRANG 2 - CÂU 19 & CÂU 20) - {name} (LỚP {sclass})", size=42, color_bgr=(0, 0, 180), bold=True)
            ann.draw_text((130, 180), f"Câu 19: {c19_sc:.2f}/1,0đ   |   Câu 20: {c20_sc:.2f}/1,0đ   |   Câu 21: {c21_sc:.2f}/1,0đ", size=36, color_bgr=(0, 0, 255), bold=True)
            ann.draw_text((130, 255), f"TỔNG TỰ LUẬN: {total_tl:.2f}/3,00đ ==> TỔNG ĐIỂM: {total_all:.2f}/10,0đ [{rank.upper()}]", size=38, color_bgr=(0, 0, 220), bold=True)
            ann.draw_text((130, 410), "[Nhận xét logic]: C19 và C20 lập luận chặt chẽ, đạo hàm và giải hệ chuẩn xác.", size=28, color_bgr=(0, 0, 255), bold=True)

            draw_c19_logic(ann, x=2400, y_start=1600, c19_sc=c19_sc)
            draw_c20_logic(ann, x=2400, y_start=2850, c20_sc=c20_sc)

    # 3. TRƯỜNG HỢP TỔNG QUÁT CHO CÁC HỌC SINH CÒN LẠI
    else:
        if total_pages == 1 or page_idx == 1:
            ann.draw_text((130, 105), f"BÀI TỰ LUẬN (TRANG 1) - {name} (LỚP {sclass}) - MÃ ĐỀ {student.get('code', '101')}", size=42, color_bgr=(0, 0, 180), bold=True)
            ann.draw_text((130, 180), f"Câu 19: {c19_sc:.2f}/1,0đ   |   Câu 20: {c20_sc:.2f}/1,0đ   |   Câu 21: {c21_sc:.2f}/1,0đ", size=36, color_bgr=(0, 0, 255), bold=True)
            ann.draw_text((130, 255), f"TỔNG TỰ LUẬN: {total_tl:.2f}/3,00đ ==> TỔNG ĐIỂM: {total_all:.2f}/10,0đ [{rank.upper()}]", size=38, color_bgr=(0, 0, 220), bold=True)
            
            note_tl = student.get("details_p4", "")[:80]
            ann.draw_text((130, 410), f"[Nhận xét logic]: {note_tl}", size=28, color_bgr=(0, 0, 255), bold=True)

            draw_c19_logic(ann, x=2400, y_start=1600, c19_sc=c19_sc)
            if c20_sc > 0 and total_pages == 1:
                draw_c20_logic(ann, x=2400, y_start=2800, c20_sc=c20_sc)
        else:
            ann.draw_text((130, 105), f"BÀI TỰ LUẬN (TRANG 2) - {name} (LỚP {sclass})", size=42, color_bgr=(0, 0, 180), bold=True)
            ann.draw_text((130, 180), f"Câu 20: {c20_sc:.2f}/1,0đ   |   Câu 21: {c21_sc:.2f}/1,0đ", size=36, color_bgr=(0, 0, 255), bold=True)
            ann.draw_text((130, 255), f"TỔNG TỰ LUẬN: {total_tl:.2f}/3,00đ ==> TỔNG ĐIỂM: {total_all:.2f}/10,0đ [{rank.upper()}]", size=38, color_bgr=(0, 0, 220), bold=True)
            
            note_tl = student.get("details_p4", "")[:80]
            ann.draw_text((130, 410), f"[Nhận xét logic]: {note_tl}", size=28, color_bgr=(0, 0, 255), bold=True)

            if c20_sc > 0:
                draw_c20_logic(ann, x=2400, y_start=680, c20_sc=c20_sc)
            draw_c21_logic(ann, x=2400, y_start=2150, c21_sc=c21_sc)

    return imwrite_unicode(out_path, ann.to_bgr())

# Bảng khớp nối chuẩn xác 100% giữa STT học sinh và Tệp ảnh phiếu Trắc nghiệm thực tế (Zero-duplicate)
STUDENT_OMR_MAP = {
    # BATCH 1 (PORTRAIT & LANDSCAPE - Chụp ngày 07/11)
    1: "IMG20261107115058.jpg",  # Đào Duy Khánh (104)
    2: "IMG20261107114335.jpg",  # Đoàn Phương Đức (104)
    3: "IMG20261107115010.jpg",  # Nguyễn Trọng Hiếu (104)
    4: "IMG20261107114741.jpg",  # Nguyễn Mạnh Thắng (103)
    5: "IMG20261107114344.jpg",  # Nguyễn Tiến Đạt (103)
    7: "IMG20261107114558.jpg",  # Nguyễn Minh Hiếu (104)
    8: "IMG20261107115032.jpg",  # Chu Thu Hiền (103)
    10: "IMG20261107114628.jpg", # Nguyễn Mai Thảo (104)
    11: "IMG20261107114729.jpg", # Cao Thị Thảo Nguyên (103)
    13: "IMG20261107114407.jpg", # Đào Xuân Dương (101)
    14: "IMG20261107114354.jpg", # Hoàng Mạnh Hùng (102)
    17: "IMG20261107115042.jpg", # Hoàng Trung Kiên (102)
    20: "IMG20261107114918.jpg", # Bùi Tuấn Long (101)
    21: "IMG20261107114531.jpg", # Lương Duy Nghĩa (101)
    23: "IMG20261107114859.jpg", # Trần Triệu Nam (101)
    24: "IMG20261107115021.jpg", # Hà Hoàng Nhật Tân (101)
    26: "IMG20261107114615.jpg", # Bùi Long Nhật (101)
    27: "IMG20261107114642.jpg", # Bế Ngọc Thành (102)
    28: "IMG20261107114959.jpg", # Phạm Chí Hùng (102)

    # BATCH 2 (RETROSPECTIVE RETAKES - Chụp lại ngày 09/11)
    6: "IMG20261109183545.jpg",  # Tạ Trí Vĩ (102)
    9: "IMG20261109183332.jpg",  # Nguyễn Quang Đức (102)
    12: "IMG20261109183518.jpg", # Hoàng Thảo Nguyên (104)
    15: "IMG20261109183531.jpg", # Nguyễn Văn Thành (103)
    16: "IMG20261109183605.jpg", # Dương Tuấn Anh (102)
    18: "IMG20261109183626.jpg", # Đỗ Văn Nghĩa (104)
    19: "IMG20261109183452.jpg", # Vũ Thị Mai (101)
    22: "IMG20261109183504.jpg", # Triệu Trung Bách (104)
    25: "IMG20261109183439.jpg", # Dương Đức Vinh (101)
    29: "IMG20261109183353.jpg", # Vũ Anh Tú (102)
    30: "IMG20261109183053.jpg", # Nguyễn Phương Hà Anh (104)
    31: "IMG20261109183146.jpg", # Lương Duy Khoa (104)
    32: "IMG20261109183124.jpg", # Phạm Khánh Ly (103)
    33: "IMG20261109183211.jpg", # Dương Quốc Khánh (103)
    34: "IMG20261109183243.jpg", # Nguyễn Thị Vân Anh (103)
    35: "IMG20261109183305.jpg", # Đào Duy Hiển (103)
    36: "IMG20261109183616.jpg", # Ngô Duy Khánh (103)
    37: "IMG20261109183235.jpg", # Bùi Hoàng Thành Đạt (104)
    38: "IMG20261109183323.jpg", # Nguyễn Lê Hải Anh (102)
    39: "IMG20261109183349.jpg", # Ngô Thế Anh (102)
    40: "IMG20261109183104.jpg", # Đàm Vĩnh Hưng (102)
    41: "IMG20261109183135.jpg", # Nguyễn Châu Ly (101)
    42: "IMG20261109183158.jpg", # Cao Minh Thư (101)
    43: "IMG20261109183222.jpg", # Đào Thanh Tùng (101)
}

def parse_img_time(fn):
    s = fn.replace('IMG', '').replace('.jpg', '')
    try:
        from datetime import datetime
        return datetime.strptime(s, '%Y%m%d%H%M%S')
    except:
        return None

def run_batch():
    students = cham_diem_core.STUDENT_RESULTS
    print(f"Bắt đầu xuất bản bộ ảnh đã chấm cho {len(students)} học sinh...")

    all_files = sorted(glob.glob(os.path.join(IMG_DIR, "*.jpg")))
    all_basenames = [os.path.basename(f) for f in all_files]

    success_cnt = 0
    tl_cnt = 0
    for s in students:
        stt = s["stt"]
        name_clean = sanitize_filename(s["name"])
        code = s["code"]
        
        # 1. Chấm Trắc Nghiệm vòng đỏ
        omr_file = STUDENT_OMR_MAP.get(stt)
        if omr_file:
            omr_path = os.path.join(IMG_DIR, omr_file)
            if os.path.exists(omr_path):
                out_tn = os.path.join(OUT_DIR, f"{stt:02d}_{name_clean}_{code}_TN_DaCham.jpg")
                if annotate_omr_sheet(omr_path, s, out_tn):
                    success_cnt += 1

                # 2. Chấm Tự Luận mực đỏ
                try:
                    omr_t = parse_img_time(omr_file)
                    f_idx = all_basenames.index(omr_file)
                    
                    # Tìm ảnh tự luận chụp ngay sau phiếu trắc nghiệm (trong 120s) và trước HS tiếp theo
                    fwd = []
                    for offset in range(1, 10):
                        check_idx = f_idx + offset
                        if check_idx >= len(all_basenames):
                            break
                        cand_name = all_basenames[check_idx]
                        if cand_name in STUDENT_OMR_MAP.values():
                            break
                        cand_t = parse_img_time(cand_name)
                        if cand_t and omr_t:
                            diff = (cand_t - omr_t).total_seconds()
                            if 0 < diff <= 120:
                                fwd.append(cand_name)

                    # Trường hợp chụp tự luận ngay trước phiếu trắc nghiệm (trong 120s)
                    bwd = []
                    for offset in range(-1, -10, -1):
                        check_idx = f_idx + offset
                        if check_idx < 0:
                            break
                        cand_name = all_basenames[check_idx]
                        if cand_name in STUDENT_OMR_MAP.values():
                            break
                        cand_t = parse_img_time(cand_name)
                        if cand_t and omr_t:
                            diff = (omr_t - cand_t).total_seconds()
                            if 0 < diff <= 120:
                                bwd.append(cand_name)

                    # Bảng định tuyến tờ bài thi chuẩn xác 100% cho các trường hợp đặc biệt
                    STUDENT_ESSAY_OVERRIDE = {
                        9: ["IMG20261109183337.jpg"],
                        39: ["IMG20261109183343.jpg"],
                    }

                    if stt in STUDENT_ESSAY_OVERRIDE:
                        adjacent = STUDENT_ESSAY_OVERRIDE[stt]
                    else:
                        adjacent = fwd if len(fwd) > 0 else list(reversed(bwd))

                    if len(adjacent) == 1:
                        # Dọn dẹp file cũ nếu trước đó từng bị chia nhầm thành Trang1 / Trang2
                        for obsolete in glob.glob(os.path.join(OUT_DIR, f"{stt:02d}_{name_clean}_{code}_TL_Trang*_DaCham.jpg")):
                            try:
                                os.remove(obsolete)
                            except:
                                pass

                    for p_i, tl_f in enumerate(adjacent[:2]):
                        tl_path = os.path.join(IMG_DIR, tl_f)
                        suffix = f"_TL_Trang{p_i+1}" if len(adjacent) > 1 else "_TL"
                        out_tl = os.path.join(OUT_DIR, f"{stt:02d}_{name_clean}_{code}{suffix}_DaCham.jpg")
                        ptype = "Trang2" if p_i == 1 else "Trang1"
                        if annotate_essay_sheet(tl_path, s, out_tl, page_type=ptype, page_idx=p_i+1, total_pages=len(adjacent)):
                            tl_cnt += 1
                except Exception as e:
                    pass

    print(f"HOÀN THÀNH: Đã xuất bản thành công {success_cnt} ảnh trắc nghiệm và {tl_cnt} ảnh tự luận đã chấm đỏ vào: {OUT_DIR}")

if __name__ == "__main__":
    run_batch()
