# -*- coding: utf-8 -*-
"""
HỆ THỐNG XUẤT BẢN BỘ ẢNH BÀI LÀM ĐÃ CHẤM ĐỎ TRỰC QUAN - LỚP 12B6
Giáo viên: ThS. Nguyễn Văn Sang - Khoa Cơ bản - Trường Cao đẳng Nghề số 1 - BQP
- Bilinear 4 góc trên 6 hàng mốc vuông thực sự (khử triệt để trôi phối cảnh & watermark).
- Phần III căn trái col 0..3 hỗ trợ đầy đủ '-', ',' và các chữ số '0'..'9'.
- Bảng Đối Soát Sư Phạm lề phải (không đè lên lưới tô hay mốc đen).
- Chấm Tự Luận phân bước 0,25đ đúng từng trang bài làm trong tl_pages, chống chồng lấn 100%.
"""

import glob
import os
import sys
import unicodedata
import cv2
import numpy as np
from PIL import Image, ImageDraw, ImageFont
import data_12b6

FONT_PATH_BOLD = r"C:\Windows\Fonts\arialbd.ttf"
FONT_PATH_REG = r"C:\Windows\Fonts\arial.ttf"
MASTER_KEYS = data_12b6.MASTER_KEYS


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
            except Exception:
                self._fonts[key] = ImageFont.load_default()
        return self._fonts[key]

    def _sanitize_font_text(self, text):
        s = str(text)
        for src, dst in (
            ("✓", "ĐÚNG"),
            ("✗", "SAI"),
            ("∈", "thuộc"),
            ("∉", "không thuộc"),
            ("₀", "0"),
            ("₁", "1"),
            ("₂", "2"),
            ("₃", "3"),
            ("₄", "4"),
        ):
            s = s.replace(src, dst)
        return s

    def text_width(self, text, size=24, bold=True):
        return self.draw.textlength(self._sanitize_font_text(text), font=self.get_font(size, bold))

    def wrap_text(self, text, max_width, size=24, bold=True, max_lines=2):
        words = self._sanitize_font_text(text).split()
        if not words:
            return [""]
        lines = []
        curr = words[0]
        for w in words[1:]:
            cand = curr + " " + w
            if self.text_width(cand, size=size, bold=bold) <= max_width:
                curr = cand
            else:
                lines.append(curr)
                curr = w
                if len(lines) == max_lines - 1:
                    break
        if len(lines) < max_lines:
            used = sum(len(l.split()) for l in lines)
            rem = " ".join(words[used:])
            while len(rem) > 4 and self.text_width(rem, size=size, bold=bold) > max_width:
                rem = rem[:-4].rstrip() + "..."
            lines.append(rem)
        return lines

    def draw_text(self, xy, text, size=32, color_bgr=(0, 0, 220), bold=True):
        color_rgb = (color_bgr[2], color_bgr[1], color_bgr[0])
        self.draw.text(xy, self._sanitize_font_text(text), font=self.get_font(size, bold), fill=color_rgb)

    def draw_rect(self, box, fill_bgr=None, outline_bgr=(0, 0, 220), width=3):
        outline_rgb = (outline_bgr[2], outline_bgr[1], outline_bgr[0]) if outline_bgr else None
        fill_rgb = (fill_bgr[2], fill_bgr[1], fill_bgr[0]) if fill_bgr else None
        self.draw.rectangle(box, fill=fill_rgb, outline=outline_rgb, width=width)

    def draw_cross_red(self, box, color_bgr=(0, 0, 220), width=6):
        color_rgb = (color_bgr[2], color_bgr[1], color_bgr[0])
        (x1, y1), (x2, y2) = box
        self.draw.line([(x1, y1), (x2, y2)], fill=color_rgb, width=width)
        self.draw.line([(x1, y2), (x2, y1)], fill=color_rgb, width=width)

    def to_bgr(self):
        return cv2.cvtColor(np.array(self.pil_img), cv2.COLOR_RGB2BGR)


def sanitize_filename(name):
    nfkd = unicodedata.normalize("NFKD", name)
    no_dia = "".join(c for c in nfkd if not unicodedata.combining(c))
    no_dia = no_dia.replace("đ", "d").replace("Đ", "D")
    clean = ""
    for c in no_dia:
        if c.isalnum():
            clean += c
        elif c.isspace() or c in ("_", "-"):
            clean += "_"
    while "__" in clean:
        clean = clean.replace("__", "_")
    return clean.strip("_")


def imread_unicode(path):
    with open(path, "rb") as f:
        arr = np.asarray(bytearray(f.read()), dtype=np.uint8)
        return cv2.imdecode(arr, cv2.IMREAD_COLOR)


def imwrite_unicode(path, img):
    ext = os.path.splitext(path)[1]
    ok, buf = cv2.imencode(ext, img, [cv2.IMWRITE_JPEG_QUALITY, 92])
    if not ok:
        return False
    with open(path, "wb") as f:
        f.write(buf)
    return True


def get_true_square_markers(img):
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    h, w = gray.shape[:2]
    blur = cv2.GaussianBlur(gray, (5, 5), 0)
    adapt = cv2.adaptiveThreshold(blur, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY_INV, 181, 28)
    num_labels, _, stats, centroids = cv2.connectedComponentsWithStats(adapt)
    markers = []
    for i in range(1, num_labels):
        x, y, cw, ch, area = stats[i]
        if int(0.013 * w) <= cw <= int(0.028 * w) and int(0.013 * w) <= ch <= int(0.028 * w):
            aspect = float(cw) / ch
            fill_ratio = float(area) / (cw * ch)
            if 0.78 <= aspect <= 1.28 and fill_ratio >= 0.76:
                c1 = adapt[y:y + max(2, ch // 5), x:x + max(2, cw // 5)]
                c2 = adapt[y:y + max(2, ch // 5), x + cw - max(2, cw // 5):x + cw]
                c3 = adapt[y + ch - max(2, ch // 5):y + ch, x:x + max(2, cw // 5)]
                c4 = adapt[y + ch - max(2, ch // 5):y + ch, x + cw - max(2, cw // 5):x + cw]
                corner_fill = (np.mean(c1) + np.mean(c2) + np.mean(c3) + np.mean(c4)) / (4.0 * 255.0)
                if corner_fill >= 0.45:
                    cx, cy = int(centroids[i][0]), int(centroids[i][1])
                    if cx >= int(0.24 * w) and cy >= int(0.32 * h):
                        markers.append((cx, cy, int(cw), int(ch)))
    return markers


def cluster_rows(markers):
    if not markers:
        return []
    pts = sorted(markers, key=lambda b: (b[1], b[0]))
    rows = []
    for p in pts:
        matched = False
        for r in rows:
            if abs(p[1] - r["mean_y"]) < 45:
                r["pts"].append(p)
                r["mean_y"] = sum(x[1] for x in r["pts"]) / len(r["pts"])
                matched = True
                break
        if not matched:
            rows.append({"mean_y": float(p[1]), "pts": [p]})
    return [(r["mean_y"], sorted(r["pts"], key=lambda p: p[0])) for r in sorted(rows, key=lambda r: r["mean_y"])]


def reconstruct_row_from_ref(pts, ref_pts, expected_len):
    if len(pts) >= expected_len:
        return pts[:expected_len]
    pairs = []
    for p in pts:
        r_near = min(ref_pts, key=lambda r: abs(r[0] - p[0]))
        if abs(r_near[0] - p[0]) < 80:
            pairs.append((p[0] - r_near[0], p[1] - r_near[1]))
    if not pairs:
        return pts
    dx = int(round(np.median([x[0] for x in pairs])))
    dy = int(round(np.median([x[1] for x in pairs])))
    out = []
    for r in ref_pts[:expected_len]:
        near = min(pts, key=lambda p: abs(p[0] - (r[0] + dx)))
        if abs(near[0] - (r[0] + dx)) < 45:
            out.append(near)
        else:
            out.append((r[0] + dx, r[1] + dy, r[2], r[3]))
    return out


def extract_6_grid_rows(img):
    all_rows = cluster_rows(get_true_square_markers(img))
    main_rows = [(cy, pts) for cy, pts in all_rows if len(pts) >= 4]
    if len(main_rows) < 5:
        return None
    t1_y, t1 = main_rows[0]
    b1_y, b1 = main_rows[1]
    t2_y, t2 = main_rows[2]
    b2_y, b2 = main_rows[3]
    t3_y, t3 = main_rows[4]

    if len(b1) < 6 and len(t1) == 6:
        b1 = reconstruct_row_from_ref(b1, t1, 6)
    if len(t1) < 6 and len(b1) == 6:
        t1 = reconstruct_row_from_ref(t1, b1, 6)
    if len(t3) < 6:
        diffs = [t3[i + 1][0] - t3[i][0] for i in range(len(t3) - 1)]
        med = int(round(np.median(diffs))) if diffs else 295
        if len(t3) == 5:
            if max(diffs) > 1.65 * med:
                idx = diffs.index(max(diffs))
                t3.insert(idx + 1, ((t3[idx][0] + t3[idx + 1][0]) // 2, (t3[idx][1] + t3[idx + 1][1]) // 2, 45, 45))
            elif abs(t3[0][0] - t2[0][0]) < 80:
                t3.append((t3[-1][0] + med, t3[-1][1], 45, 45))
            else:
                t3.insert(0, (t3[0][0] - med, t3[0][1], 45, 45))

    b3_cand = [(cy, pts) for cy, pts in all_rows if 500 <= (cy - t3_y) <= 850 and len(pts) >= 2]
    if b3_cand:
        _, b3_raw = max(b3_cand, key=lambda x: len(x[1]))
        b3 = reconstruct_row_from_ref(b3_raw, t3, 6)
    else:
        dx = b1[0][0] - t1[0][0]
        dy = int(round(1.16 * (b1_y - t1_y)))
        b3 = [(p[0] + dx, p[1] + dy, p[2], p[3]) for p in t3]

    return t1, b1, t2, b2, t3, b3


def bilinear(tl, tr, bl, br, u, v):
    x = (1 - u) * (1 - v) * tl[0] + u * (1 - v) * tr[0] + (1 - u) * v * bl[0] + u * v * br[0]
    y = (1 - u) * (1 - v) * tl[1] + u * (1 - v) * tr[1] + (1 - u) * v * bl[1] + u * v * br[1]
    return int(round(x)), int(round(y))


def match_pair(row_top, row_bot, i_left, i_right):
    tl = row_top[i_left]
    tr = row_top[i_right]
    bl = min(row_bot, key=lambda p: abs(p[0] - tl[0]))
    br = min(row_bot, key=lambda p: abs(p[0] - tr[0]))
    return tl, tr, bl, br


def p3_v_for_char(ch):
    if ch == "-":
        return 0.068
    if ch in (",", "."):
        return 0.148
    if ch.isdigit():
        return 0.229 + int(ch) * 0.0804
    return None


def annotate_omr_sheet(img_path, student, out_path):
    img = imread_unicode(img_path)
    if img is None:
        return False
    h, w = img.shape[:2]
    if w > h:
        img = cv2.rotate(img, cv2.ROTATE_90_COUNTERCLOCKWISE)
        h, w = img.shape[:2]

    code = str(student["code"])
    key = MASTER_KEYS[code]
    grid = extract_6_grid_rows(img)
    assert grid is not None, f"Không định vị được 6 hàng mốc OMR cho STT {student['stt']} ({student['name']})"
    t1, b1, t2, b2, t3, b3 = grid

    COLOR_GREEN = (40, 180, 20)   # BGR: Tô đúng (Xanh lá)
    COLOR_AMBER = (10, 140, 245)  # BGR: Tô sai (Cam)
    COLOR_RED = (0, 0, 240)       # BGR: Viền đỏ đáp án chuẩn
    COLOR_BLUE = (240, 160, 40)   # BGR: Mã đề (Xanh dương)

    st_p1 = student["answers_p1"]
    st_p2 = student["answers_p2"]
    st_p3 = student["answers_p3"]

    overlay = img.copy()
    ring_ops = []  # Vẽ sau addWeighted để viền luôn sắc nét

    # 0. Mã đề
    tl_md, tr_md, bl_md, br_md = match_pair(t1, b1, 1, 2)
    for c_md, ch_md in enumerate(code[:3]):
        if ch_md.isdigit():
            u = 0.235 + c_md * 0.175
            v = 0.094 + int(ch_md) * 0.0908
            pt = bilinear(tl_md, tr_md, bl_md, br_md, u, v)
            cv2.circle(overlay, pt, 18, COLOR_BLUE, -1)
            ring_ops.append((pt, 21, COLOR_BLUE, 3))

    # 1. Phần I (12 câu ABCD)
    tl1, tr1, bl1, br1 = match_pair(t1, b1, 2, 3)
    tl2, tr2, bl2, br2 = match_pair(t1, b1, 3, 4)
    col_ratios = [0.240, 0.423, 0.606, 0.789]
    for q in range(12):
        quad = (tl1, tr1, bl1, br1) if q < 10 else (tl2, tr2, bl2, br2)
        row_idx = q if q < 10 else (q - 10)
        v = 0.094 + row_idx * 0.0908

        ans_key = key["p1"][q]
        pt_key = bilinear(*quad, col_ratios[ord(ans_key) - ord("A")], v)
        st_ans = st_p1[q] if q < len(st_p1) else ""

        if st_ans == ans_key:
            cv2.circle(overlay, pt_key, 22, COLOR_GREEN, -1)
            ring_ops.append((pt_key, 25, COLOR_GREEN, 3))
        else:
            if st_ans in ("A", "B", "C", "D"):
                pt_st = bilinear(*quad, col_ratios[ord(st_ans) - ord("A")], v)
                cv2.circle(overlay, pt_st, 22, COLOR_AMBER, -1)
                ring_ops.append((pt_st, 24, COLOR_AMBER, 3))
            ring_ops.append((pt_key, 27, COLOR_RED, 4))

    # 2. Phần II (Câu 13, 14)
    y_p2 = [0.203, 0.391, 0.590, 0.790]
    key_p2 = key["p2_13"] + key["p2_14"]
    for c_idx in range(2):
        quad = match_pair(t2, b2, c_idx, c_idx + 1)
        for sub in range(4):
            p2_idx = c_idx * 4 + sub
            k_ds = key_p2[p2_idx]
            s_ds = st_p2[p2_idx] if p2_idx < len(st_p2) else ""
            u_k = 0.241 if k_ds == "Đ" else 0.543
            pt_k = bilinear(*quad, u_k, y_p2[sub])

            if s_ds == k_ds:
                cv2.circle(overlay, pt_k, 22, COLOR_GREEN, -1)
                ring_ops.append((pt_k, 25, COLOR_GREEN, 3))
            else:
                if s_ds in ("Đ", "S"):
                    u_s = 0.241 if s_ds == "Đ" else 0.543
                    pt_s = bilinear(*quad, u_s, y_p2[sub])
                    cv2.circle(overlay, pt_s, 22, COLOR_AMBER, -1)
                    ring_ops.append((pt_s, 24, COLOR_AMBER, 3))
                ring_ops.append((pt_k, 27, COLOR_RED, 4))

    # 3. Phần III (4 câu điền số căn trái col 0..3)
    p3_badges = []
    for qi in range(4):
        quad = match_pair(t3, b3, qi, qi + 1)
        ans_str = str(key["p3"][qi]).strip()
        st_str = str(st_p3[qi]).strip() if qi < len(st_p3) else ""
        norm_st = st_str[:-2] if st_str.endswith(".0") else st_str
        is_correct = (norm_st == ans_str)

        for col_i, ch in enumerate(st_str[:4]):
            vv = p3_v_for_char(ch)
            if vv is not None:
                u = 0.248 + col_i * 0.157
                pt_s = bilinear(*quad, u, vv)
                col_c = COLOR_GREEN if is_correct else COLOR_AMBER
                cv2.circle(overlay, pt_s, 20, col_c, -1)
                ring_ops.append((pt_s, 23, col_c, 3))

        if not is_correct:
            for col_i, ch in enumerate(ans_str[:4]):
                vv = p3_v_for_char(ch)
                if vv is not None:
                    u = 0.248 + col_i * 0.157
                    pt_k = bilinear(*quad, u, vv)
                    ring_ops.append((pt_k, 25, COLOR_RED, 4))

        p3_badges.append((quad[0], quad[1], ans_str, st_str, is_correct))

    # Blend overlay bán trong suốt, sau đó vẽ toàn bộ viền sắc nét
    img = cv2.addWeighted(overlay, 0.42, img, 0.58, 0)
    for pt, rad, col, thick in ring_ops:
        cv2.circle(img, pt, rad, col, thick)

    ann = VietnameseAnnotator(img)

    # Nhãn ĐA phía trên từng cột Phần III
    for tlq, trq, ans_str, st_str, is_ok in p3_badges:
        bx1 = tlq[0] + 6
        bx2 = trq[0] - 6
        top_y = min(tlq[1], trq[1]) - 74
        bcol = (0, 160, 0) if is_ok else (0, 0, 230)
        ann.draw_rect([(bx1, top_y), (bx2, top_y + 46)], fill_bgr=(255, 255, 255), outline_bgr=bcol, width=3)
        lbl = f"ĐA: {ans_str}" if is_ok else f"{st_str or 'Trống'} | ĐA: {ans_str}"
        fsize = 24 if ann.text_width(lbl, 24, True) <= (bx2 - bx1 - 12) else 20
        ann.draw_text((bx1 + 8, top_y + 10), lbl, size=fsize, color_bgr=bcol, bold=True)

    # Banner tiêu đề đầu trang (y <= 345)
    ann.draw_rect([(70, 28), (w - 70, 345)], fill_bgr=(255, 255, 255), outline_bgr=(0, 0, 220), width=4)
    name_str = f"STT {student['stt']:02d}: {student['name'].upper()} - LỚP: {student['class']} - MÃ ĐỀ: {code}"
    ann.draw_text((95, 42), name_str, size=42, color_bgr=(0, 0, 220), bold=True)
    ann.draw_text((95, 100), f"PHẦN I (12 câu ABCD): {student['p1']:.2f} / 3,00 điểm", size=31, color_bgr=(0, 0, 200), bold=True)
    ann.draw_text(
        (95, 142),
        f"PHẦN II (Đúng/Sai):    {student['p2']:.2f} / 2,00 điểm   (C13: {student['p2_c13']:.2f}đ | C14: {student['p2_c14']:.2f}đ)",
        size=31, color_bgr=(0, 0, 200), bold=True,
    )
    ann.draw_text((95, 184), f"PHẦN III (Điền số):    {student['p3']:.2f} / 2,00 điểm", size=31, color_bgr=(0, 0, 200), bold=True)
    sc_str = f"TỔNG TN: {student['total_tn']:.2f}/7,00đ  |  TỰ LUẬN: {student['p4']:.2f}/3,00đ  |  TỔNG ĐIỂM: {student['total']:.2f}/10,0đ ({student['rank']})"
    ann.draw_text((95, 234), sc_str, size=34, color_bgr=(0, 0, 220), bold=True)
    ann.draw_text(
        (95, 290),
        "[XANH LÁ = TÔ ĐÚNG | CAM = TÔ SAI | VIỀN ĐỎ = ĐÁP ÁN CHUẨN | XANH DƯƠNG = MÃ ĐỀ]",
        size=25, color_bgr=(160, 0, 0), bold=True,
    )

    # BẢNG ĐỐI SOÁT SƯ PHẠM LỀ PHẢI (Cách ly an toàn khỏi lưới tô & mốc đen)
    panel_x2 = w - 32
    p1_x1 = max(w - 640, max(t1[4][0], b1[4][0]) + 32)
    p1_y1, p1_y2 = 1420, 2250
    ann.draw_rect([(p1_x1, p1_y1), (panel_x2, p1_y2)], fill_bgr=(255, 255, 255), outline_bgr=(0, 0, 200), width=3)
    ann.draw_rect([(p1_x1, p1_y1), (panel_x2, p1_y1 + 52)], fill_bgr=(230, 240, 255), outline_bgr=(0, 0, 200), width=2)
    hdr_p1 = f"ĐỐI SOÁT P.I (MÃ {code}): {student['p1']:.2f}/3.00đ"
    ann.draw_text((p1_x1 + 14, p1_y1 + 12), hdr_p1, size=24, color_bgr=(180, 0, 0), bold=True)

    col_w = (panel_x2 - p1_x1 - 16) // 2
    for q in range(12):
        ci = 0 if q < 6 else 1
        ri = q if q < 6 else (q - 6)
        cx = p1_x1 + 10 + ci * col_w
        cy = p1_y1 + 66 + ri * 124
        ans_k = key["p1"][q]
        ans_s = st_p1[q] if q < len(st_p1) else ""
        is_ok = (ans_s == ans_k)
        fg = (0, 140, 0) if is_ok else (0, 0, 220)
        bg = (235, 255, 235) if is_ok else (255, 235, 235)
        ann.draw_rect([(cx, cy), (cx + col_w - 8, cy + 106)], fill_bgr=bg, outline_bgr=fg, width=2)
        if is_ok:
            l1 = f"C{q+1:02d}: {ans_s} [ĐÚNG]"
            l2 = "+0.25đ"
        else:
            l1 = f"C{q+1:02d}: {ans_s or '-'} [SAI]"
            l2 = f"ĐA: {ans_k}"
        fs = 21 if ann.text_width(l1, 21, True) <= (col_w - 22) else 18
        ann.draw_text((cx + 10, cy + 18), l1, size=fs, color_bgr=fg, bold=True)
        ann.draw_text((cx + 10, cy + 56), l2, size=21, color_bgr=fg, bold=True)

    # Panel Phần II & Phần III (nằm dưới hàng b1, bên phải t2/t3)
    p23_x1 = max(w - 690, max(t3[4][0], b3[4][0]) + 36)
    p2_y1, p2_y2 = 2295, 2745
    ann.draw_rect([(p23_x1, p2_y1), (panel_x2, p2_y2)], fill_bgr=(255, 255, 255), outline_bgr=(0, 0, 200), width=3)
    ann.draw_rect([(p23_x1, p2_y1), (panel_x2, p2_y1 + 52)], fill_bgr=(230, 240, 255), outline_bgr=(0, 0, 200), width=2)
    ann.draw_text((p23_x1 + 14, p2_y1 + 12), f"ĐỐI SOÁT PHẦN II: {student['p2']:.2f}/2.00đ", size=24, color_bgr=(180, 0, 0), bold=True)

    sub_names = ["a", "b", "c", "d"]
    sub_pts = [student["p2_c13"], student["p2_c14"]]
    box_w = (panel_x2 - p23_x1 - 36) // 4
    for c_i, c_num in enumerate((13, 14)):
        cy_c = p2_y1 + 66 + c_i * 188
        ann.draw_text((p23_x1 + 14, cy_c), f"Câu {c_num} (+{sub_pts[c_i]:.2f}đ):", size=24, color_bgr=(0, 0, 180), bold=True)
        for s_i in range(4):
            idx_sub = c_i * 4 + s_i
            k_ds = key_p2[idx_sub]
            s_ds = st_p2[idx_sub] if idx_sub < len(st_p2) else ""
            is_ok = (s_ds == k_ds)
            bx = p23_x1 + 14 + s_i * (box_w + 4)
            by = cy_c + 38
            fg = (0, 145, 0) if is_ok else (0, 0, 220)
            bg = (235, 255, 235) if is_ok else (255, 235, 235)
            ann.draw_rect([(bx, by), (bx + box_w, by + 112)], fill_bgr=bg, outline_bgr=fg, width=2)
            ann.draw_text((bx + 10, by + 16), f"{sub_names[s_i]}) {s_ds or '-'}", size=23, color_bgr=fg, bold=True)
            tag2 = "[Đ]" if is_ok else f"[S>{k_ds}]"
            ann.draw_text((bx + 10, by + 60), tag2, size=22, color_bgr=fg, bold=True)

    p3_y1, p3_y2 = 2785, 3265
    ann.draw_rect([(p23_x1, p3_y1), (panel_x2, p3_y2)], fill_bgr=(255, 255, 255), outline_bgr=(0, 0, 200), width=3)
    ann.draw_rect([(p23_x1, p3_y1), (panel_x2, p3_y1 + 52)], fill_bgr=(230, 240, 255), outline_bgr=(0, 0, 200), width=2)
    ann.draw_text((p23_x1 + 14, p3_y1 + 12), f"ĐỐI SOÁT PHẦN III: {student['p3']:.2f}/2.00đ", size=24, color_bgr=(180, 0, 0), bold=True)

    for i in range(4):
        c_num = 15 + i
        k_val = str(key["p3"][i]).strip()
        s_val = str(st_p3[i]).strip() if i < len(st_p3) else ""
        norm_s = s_val[:-2] if s_val.endswith(".0") else s_val
        is_ok = (norm_s == k_val)
        cy_q = p3_y1 + 66 + i * 102
        fg = (0, 145, 0) if is_ok else (0, 0, 220)
        bg = (235, 255, 235) if is_ok else (255, 235, 235)
        ann.draw_rect([(p23_x1 + 14, cy_q), (panel_x2 - 14, cy_q + 84)], fill_bgr=bg, outline_bgr=fg, width=2)
        if is_ok:
            txt = f"Câu {c_num}: Tô '{s_val}' [ĐÚNG] +0.50đ"
        else:
            txt = f"Câu {c_num}: Tô '{s_val or 'Trống'}' [SAI] -> ĐA: {k_val}"
        fs = 24 if ann.text_width(txt, 24, True) <= (panel_x2 - p23_x1 - 44) else 20
        ann.draw_text((p23_x1 + 24, cy_q + 26), txt, size=fs, color_bgr=fg, bold=True)

    return imwrite_unicode(out_path, ann.to_bgr())


def _rects_overlap(r1, r2):
    x1 = max(r1[0], r2[0])
    y1 = max(r1[1], r2[1])
    x2 = min(r1[2], r2[2])
    y2 = min(r1[3], r2[3])
    return x2 > x1 and y2 > y1


def annotate_essay_sheet_pedagogical(img_path, student, out_path, page_info, page_idx=1, total_pages=1, missing_on_final=()):
    img = imread_unicode(img_path)
    if img is None:
        return False
    h, w = img.shape[:2]
    if w > h:
        img = cv2.rotate(img, cv2.ROTATE_90_COUNTERCLOCKWISE)
        h, w = img.shape[:2]

    ann = VietnameseAnnotator(img)
    drawn_rects = []

    def add_rect_checked(box, fill_bgr=None, outline_bgr=(0, 0, 220), width=3):
        (bx1, by1), (bx2, by2) = box
        rc = (int(bx1), int(by1), int(bx2), int(by2))
        for prev in drawn_rects:
            assert not _rects_overlap(rc, prev), f"STT {student['stt']} P{page_idx}: Chồng chéo khung {rc} vs {prev}"
        drawn_rects.append(rc)
        ann.draw_rect(box, fill_bgr=fill_bgr, outline_bgr=outline_bgr, width=width)

    # Banner đầu trang (cao 220px -> y <= 248)
    add_rect_checked([(70, 28), (w - 70, 250)], fill_bgr=(255, 255, 255), outline_bgr=(0, 0, 220), width=4)
    page_tag = f" (Trang {page_idx}/{total_pages})" if total_pages > 1 else ""
    header_title = f"BÀI THI TỰ LUẬN (PHẦN IV) - {student['name'].upper()} - LỚP {student['class']} - MÃ ĐỀ {student['code']}{page_tag}"
    f_hdr = 38 if ann.text_width(header_title, 38, True) <= (w - 190) else 32
    ann.draw_text((95, 42), header_title, size=f_hdr, color_bgr=(0, 0, 220), bold=True)

    c19_sc, c20_sc, c21_sc = student["answers_p4"]
    sub_title = f"Câu 19: {c19_sc:.2f}/1.00đ  |  Câu 20: {c20_sc:.2f}/1.00đ  |  Câu 21: {c21_sc:.2f}/1.00đ"
    ann.draw_text((95, 100), sub_title, size=31, color_bgr=(0, 0, 200), bold=True)

    summary_title = f"TỔNG TỰ LUẬN: {student['p4']:.2f} / 3,00 điểm  ===>  TỔNG KẾT TOÀN BÀI: {student['total']:.2f} / 10,0 điểm ({student['rank']})"
    ann.draw_text((95, 154), summary_title, size=34, color_bgr=(0, 0, 220), bold=True)

    # Các câu thực sự nằm trên trang này (sắp theo y_start từ trên xuống)
    y_bands = page_info.get("y_bands", {})
    questions = sorted(page_info.get("questions", []), key=lambda q: y_bands.get(q, [0.5, 0.9])[0])
    if not questions and missing_on_final:
        questions = []

    p4_det = student["p4_detail"]
    card_x1 = w - 1140
    card_x2 = w - 75
    text_max_w = (card_x2 - card_x1) - 36

    reserve_bottom = len(missing_on_final) * 76 + (16 if missing_on_final else 0)
    avail_bottom = h - 45 - reserve_bottom
    cursor_y = 278

    for q_idx, q_key in enumerate(questions):
        q_num = q_key.replace("C", "")
        q_info = p4_det[q_key]
        sc = q_info["score"]
        steps = q_info["steps"]
        err_note = (q_info.get("error_note") or "").strip()
        has_note_bar = bool(err_note and sc < 1.0)

        rem_q = len(questions) - q_idx
        max_for_this = cursor_y + int((avail_bottom - cursor_y - (rem_q - 1) * 20) / rem_q)

        # Anchor theo y0 thực tế của câu trên ảnh nhưng bảo toàn tối thiểu không gian & không đè câu dưới
        y0_norm, y1_norm = y_bands.get(q_key, [0.25, 0.85])
        desired_top = max(cursor_y, int(y0_norm * h))
        needed_min = 4 * 78 + 90 + (62 if has_note_bar else 0)
        if desired_top + needed_min > max_for_this:
            desired_top = max(cursor_y, max_for_this - needed_min)

        avail_h = max_for_this - desired_top
        total_box_h = min(84, max(62, int(avail_h * 0.16)))
        note_h = min(64, max(48, int(avail_h * 0.12))) if has_note_bar else 0
        steps_budget = avail_h - total_box_h - note_h - 32
        step_pitch = min(130, max(64, steps_budget // 4))
        step_box_h = min(98, max(54, step_pitch - 10))
        step_font = 23 if step_box_h >= 78 else 19

        y_top_card = desired_top
        for s_idx, st in enumerate(steps[:4]):
            sy1 = desired_top + s_idx * step_pitch
            sy2 = sy1 + step_box_h
            is_ok = st["ok"]
            fg = (0, 145, 0) if is_ok else (0, 0, 220)
            bg = (240, 255, 240) if is_ok else (255, 240, 240)
            tag = "+0,25đ [ĐÚNG]" if is_ok else "0,00đ [SAI]"
            full_txt = f"C{q_num}.{s_idx+1} {tag}: {st['text']}"
            lines = ann.wrap_text(full_txt, text_max_w, size=step_font, bold=True, max_lines=2 if step_box_h >= 74 else 1)

            add_rect_checked([(card_x1, sy1), (card_x2, sy2)], fill_bgr=bg, outline_bgr=fg, width=3)
            if len(lines) == 1:
                ann.draw_text((card_x1 + 16, sy1 + (step_box_h - step_font) // 2 - 2), lines[0], size=step_font, color_bgr=fg, bold=True)
            else:
                ann.draw_text((card_x1 + 16, sy1 + 8), lines[0], size=step_font, color_bgr=fg, bold=True)
                ann.draw_text((card_x1 + 16, sy1 + 12 + step_font + 4), lines[1], size=step_font, color_bgr=fg, bold=True)

        sum_y1 = desired_top + 4 * step_pitch + 4
        sum_y2 = sum_y1 + total_box_h
        sum_fg = (0, 145, 0) if sc >= 0.75 else ((180, 100, 0) if sc > 0 else (0, 0, 220))
        sum_bg = (232, 255, 232) if sc >= 0.75 else ((255, 248, 225) if sc > 0 else (255, 232, 232))
        add_rect_checked([(card_x1, sum_y1), (card_x2, sum_y2)], fill_bgr=sum_bg, outline_bgr=sum_fg, width=4)
        extra_ess = " — [SAI BẢN CHẤT -> 0,00đ]" if q_info.get("wrong_in_essence") else ""
        sum_txt = f"[ TỔNG CÂU {q_num}: {sc:.2f}đ / 1,00đ ]{extra_ess}"
        s_f = 30 if ann.text_width(sum_txt, 30, True) <= text_max_w else 24
        ann.draw_text((card_x1 + 18, sum_y1 + (total_box_h - s_f) // 2 - 2), sum_txt, size=s_f, color_bgr=sum_fg, bold=True)

        if q_info.get("wrong_in_essence"):
            # Gạch chéo đỏ cảnh báo lên vùng bài làm bên trái khung điểm
            ann.draw_cross_red([(140, y_top_card + 20), (card_x1 - 50, sum_y2 - 10)], color_bgr=(0, 0, 225), width=6)

        end_y = sum_y2
        if has_note_bar:
            ny1 = sum_y2 + 8
            ny2 = ny1 + note_h
            add_rect_checked([(card_x1, ny1), (card_x2, ny2)], fill_bgr=(255, 245, 230), outline_bgr=(0, 90, 220), width=2)
            n_lines = ann.wrap_text(f"Lỗi: {err_note}", text_max_w, size=20, bold=True, max_lines=2 if note_h >= 56 else 1)
            if len(n_lines) == 1:
                ann.draw_text((card_x1 + 14, ny1 + (note_h - 20) // 2 - 2), n_lines[0], size=20, color_bgr=(0, 60, 200), bold=True)
            else:
                ann.draw_text((card_x1 + 14, ny1 + 6), n_lines[0], size=19, color_bgr=(0, 60, 200), bold=True)
                ann.draw_text((card_x1 + 14, ny1 + 29), n_lines[1], size=19, color_bgr=(0, 60, 200), bold=True)
            end_y = ny2

        cursor_y = end_y + 22

    # Thẻ báo câu bỏ trống hoàn toàn ở cuối trang cuối
    if missing_on_final:
        by = max(cursor_y, h - 45 - len(missing_on_final) * 74)
        for mq in missing_on_final:
            q_num = mq.replace("C", "")
            add_rect_checked([(card_x1, by), (card_x2, by + 62)], fill_bgr=(255, 235, 235), outline_bgr=(0, 0, 220), width=3)
            ann.draw_text((card_x1 + 18, by + 16), f"[ CÂU {q_num}: 0,00đ / 1,00đ — Học sinh bỏ trống không làm ]", size=26, color_bgr=(0, 0, 220), bold=True)
            by += 72

    return imwrite_unicode(out_path, ann.to_bgr())


def main():
    base_in = r"D:\TOAN\05_CONG_CU_VA_ENGINE\Cham_Diem_Hoc_Sinh\01_ANH_CHUA_CHAM\Lop_12B6"
    base_out = r"D:\TOAN\05_CONG_CU_VA_ENGINE\Cham_Diem_Hoc_Sinh\02_ANH_DA_CHAM\Lop_12B6"
    os.makedirs(base_out, exist_ok=True)

    for f in glob.glob(os.path.join(base_out, "*.*")):
        os.remove(f)

    students = data_12b6.STUDENTS_12B6
    print(f"Bắt đầu xuất bản bộ ảnh chấm bút đỏ Lớp 12B6 ({len(students)} học sinh)...")

    omr_ok = 0
    tl_ok = 0
    for s in students:
        stt = s["stt"]
        name_clean = sanitize_filename(s["name"])
        code = s["code"]

        in_omr = os.path.join(base_in, s["file_omr"])
        out_omr_name = f"12B6_{stt:02d}_{name_clean}_{code}_DaCham.jpg"
        out_omr = os.path.join(base_out, out_omr_name)
        assert annotate_omr_sheet(in_omr, s, out_omr), f"OMR failed for STT {stt}"
        omr_ok += 1

        tl_pages = s["tl_pages"]
        total_tl_pages = len(tl_pages)
        present_qs = {q for pg in tl_pages for q in pg.get("questions", [])}
        missing_qs = [q for q in ("C19", "C20", "C21") if q not in present_qs]

        for p_idx, pg in enumerate(tl_pages, start=1):
            in_tl = os.path.join(base_in, pg["file"])
            if total_tl_pages == 1:
                out_tl_name = f"12B6_{stt:02d}_{name_clean}_{code}_TuLuan_DaCham.jpg"
            else:
                out_tl_name = f"12B6_{stt:02d}_{name_clean}_{code}_TuLuan_Trang{p_idx}_DaCham.jpg"
            out_tl = os.path.join(base_out, out_tl_name)
            m_final = missing_qs if (p_idx == total_tl_pages) else ()
            assert annotate_essay_sheet_pedagogical(
                in_tl, s, out_tl, page_info=pg, page_idx=p_idx, total_pages=total_tl_pages, missing_on_final=m_final
            ), f"TL failed for STT {stt} page {p_idx}"
            tl_ok += 1

    print(f"HOÀN THÀNH: Đã xuất bản {omr_ok} ảnh OMR và {tl_ok} ảnh Tự luận -> {base_out}")

    print("\n>>> TỰ ĐỘNG CHẠY HỆ THỐNG DOUBLE CHECK 2 VÒNG ĐỘC LẬP...")
    from double_check_cham_diem import run_double_check
    if not run_double_check("Lop_12B6", expected_count=34):
        sys.exit(1)


if __name__ == "__main__":
    main()
