# -*- coding: utf-8 -*-
"""
HỆ THỐNG XUẤT BẢN BỘ ẢNH BÀI LÀM ĐÃ CHẤM ĐỎ TRỰC QUAN - LỚP 10A17 (TOÁN 10)
Giáo viên: ThS. Nguyễn Văn Sang - Khoa Cơ bản - Trường Cao đẳng Nghề số 1 - BQP
"""

import glob
import os
import tempfile
import cv2
import numpy as np
import data_10a17
import tao_anh_da_cham_12b6 as engine

engine.MASTER_KEYS = data_10a17.MASTER_KEYS


def ensure_upright_omr_file(in_path, tmp_dir):
    img = engine.imread_unicode(in_path)
    if img is None:
        return in_path
    h, w = img.shape[:2]
    if h >= w:
        return in_path
    for rot_code in (cv2.ROTATE_90_COUNTERCLOCKWISE, cv2.ROTATE_90_CLOCKWISE):
        cand = cv2.rotate(img, rot_code)
        grid = engine.extract_6_grid_rows(cand)
        if grid is not None:
            t1, b1, t2, b2, t3, b3 = grid
            ch, cw = cand.shape[:2]
            if t1[0][0] >= 0.25 * cw and 0.32 * ch <= t1[0][1] <= 0.48 * ch and b3[0][1] >= 0.83 * ch:
                if rot_code == cv2.ROTATE_90_COUNTERCLOCKWISE:
                    return in_path
                tmp_path = os.path.join(tmp_dir, os.path.basename(in_path))
                engine.imwrite_unicode(tmp_path, cand)
                return tmp_path
    return in_path


def ensure_upright_tl_file(in_path, tmp_dir):
    img = engine.imread_unicode(in_path)
    if img is None:
        return in_path
    h, w = img.shape[:2]
    if h >= w:
        return in_path
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
        return in_path
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    if np.mean(gray[:, :int(0.10 * w)]) > np.mean(gray[:, int(0.90 * w):]):
        cand = cv2.rotate(img, cv2.ROTATE_90_CLOCKWISE)
        tmp_path = os.path.join(tmp_dir, "tl_" + os.path.basename(in_path))
        engine.imwrite_unicode(tmp_path, cand)
        return tmp_path
    return in_path


def annotate_blank_code_omr(in_path, student, out_path):
    img = engine.imread_unicode(in_path)
    if img is None:
        return False
    ann = engine.VietnameseAnnotator(img)
    h, w = img.shape[:2]
    ann.draw_rect([(60, 24), (w - 60, 220)], fill_bgr=(255, 243, 220), outline_bgr=(0, 0, 220), width=4)
    ann.draw_text((90, 42), f"STT {student['stt']:02d}: {student['name'].upper()} - LỚP 10A17 - CHƯA TÔ MÃ ĐỀ", size=34, color_bgr=(0, 0, 220), bold=True)
    ann.draw_text((90, 102), "CẢNH BÁO: Thí sinh không ghi và không tô Mã đề thi -> Tạm tính 0.00đ phần Trắc nghiệm.", size=28, color_bgr=(0, 70, 200), bold=True)
    ann.draw_text((90, 152), f"PHẦN IV (Tự luận): {student['p4']:.2f}/3.00đ  |  TỔNG TẠM TÍNH: {student['total']:.2f}/10.0đ", size=28, color_bgr=(0, 0, 200), bold=True)
    return engine.imwrite_unicode(out_path, ann.to_bgr())


def annotate_essay_10a17(img_path, student, out_path, page_info, page_idx=1, total_pages=1, missing_on_final=()):
    img = engine.imread_unicode(img_path)
    if img is None:
        return False
    h, w = img.shape[:2]
    if w > h:
        img = cv2.rotate(img, cv2.ROTATE_90_COUNTERCLOCKWISE)
        h, w = img.shape[:2]

    ann = engine.VietnameseAnnotator(img)
    drawn_rects = []

    def add_rect_checked(box, fill_bgr=None, outline_bgr=(0, 0, 220), width=3):
        (bx1, by1), (bx2, by2) = box
        rc = (int(bx1), int(by1), int(bx2), int(by2))
        for prev in drawn_rects:
            if max(rc[0], prev[0]) < min(rc[2], prev[2]) and max(rc[1], prev[1]) < min(rc[3], prev[3]):
                return
        drawn_rects.append(rc)
        ann.draw_rect(box, fill_bgr=fill_bgr, outline_bgr=outline_bgr, width=width)

    add_rect_checked([(70, 28), (w - 70, 250)], fill_bgr=(255, 255, 255), outline_bgr=(0, 0, 220), width=4)
    page_tag = f" (Trang {page_idx}/{total_pages})" if total_pages > 1 else ""
    header_title = f"BÀI THI TỰ LUẬN (TOÁN 10) - {student['name'].upper()} - LỚP 10A17 - MÃ ĐỀ {student['code']}{page_tag}"
    f_hdr = 34 if ann.text_width(header_title, 34, True) <= (w - 190) else 28
    ann.draw_text((95, 42), header_title, size=f_hdr, color_bgr=(0, 0, 210), bold=True)

    p4_det = student["p4_detail"]
    c19_sc, c19_mx = p4_det["C19"]["score"], p4_det["C19"]["max_score"]
    c20_sc, c20_mx = p4_det["C20"]["score"], p4_det["C20"]["max_score"]
    c21_sc, c21_mx = p4_det["C21"]["score"], p4_det["C21"]["max_score"]
    ann.draw_text((95, 100), f"C19: {c19_sc:.2f}/{c19_mx:.2f}đ   |   C20: {c20_sc:.2f}/{c20_mx:.2f}đ   |   C21: {c21_sc:.2f}/{c21_mx:.2f}đ", size=30, color_bgr=(180, 0, 0), bold=True)
    ann.draw_text((95, 158), f"TỔNG TỰ LUẬN: {student['p4']:.2f} / 3,00đ   |   TỔNG KẾT TOÀN BÀI: {student['total']:.2f} / 10,0đ  [{student['rank']}]", size=30, color_bgr=(0, 110, 0) if student["total"] >= 5.0 else (0, 0, 210), bold=True)

    questions = page_info.get("questions", [])
    if not questions:
        add_rect_checked([(int(0.45 * w), 310), (w - 50, 410)], fill_bgr=(255, 248, 225), outline_bgr=(0, 90, 220), width=3)
        ann.draw_text((int(0.45 * w) + 20, 342), "TRANG NHÁP / BỔ SUNG (Không tính điểm tự luận mới)", size=24, color_bgr=(0, 70, 200), bold=True)
        return engine.imwrite_unicode(out_path, ann.to_bgr())

    card_x1 = int(0.47 * w)
    card_x2 = w - 46
    text_max_w = card_x2 - card_x1 - 30
    y_bands = page_info.get("y_bands", {})
    cursor_y = 268
    avail_bottom = h - 38 - (len(missing_on_final) * 76 if missing_on_final else 0)

    for q_idx, q_key in enumerate(questions):
        q_num = q_key.replace("C", "")
        q_info = p4_det[q_key]
        sc, mx = q_info["score"], q_info["max_score"]
        steps = q_info.get("steps", [])
        n_steps = max(1, min(4, len(steps)))
        err_note = (q_info.get("error_note") or "").strip()
        has_note_bar = bool(err_note and sc < mx)

        rem_q = len(questions) - q_idx
        max_for_this = cursor_y + int((avail_bottom - cursor_y - (rem_q - 1) * 20) / rem_q)
        y0_norm, _ = y_bands.get(q_key, [0.25, 0.85])
        desired_top = max(cursor_y, int(y0_norm * h))
        needed_min = n_steps * 76 + 90 + (60 if has_note_bar else 0)
        if desired_top + needed_min > max_for_this:
            desired_top = max(cursor_y, max_for_this - needed_min)

        avail_h = max_for_this - desired_top
        total_box_h = min(80, max(58, int(avail_h * 0.16)))
        note_h = min(62, max(46, int(avail_h * 0.12))) if has_note_bar else 0
        steps_budget = avail_h - total_box_h - note_h - 32
        step_pitch = min(124, max(62, steps_budget // n_steps))
        step_box_h = min(94, max(52, step_pitch - 10))
        step_font = 22 if step_box_h >= 74 else 18

        y_top_card = desired_top
        for s_idx, st in enumerate(steps[:n_steps]):
            sy1 = desired_top + s_idx * step_pitch
            sy2 = sy1 + step_box_h
            is_ok = st.get("ok", False)
            pts = st.get("pts", 0.25 if is_ok else 0.0)
            fg = (0, 145, 0) if is_ok else (0, 0, 220)
            bg = (240, 255, 240) if is_ok else (255, 240, 240)
            tag = f"+{pts:.2f}đ [ĐÚNG]" if is_ok else "0,00đ [SAI]"
            full_txt = f"C{q_num}.{s_idx+1} {tag}: {st.get('text', '')}"
            lines = ann.wrap_text(full_txt, text_max_w, size=step_font, bold=True, max_lines=2 if step_box_h >= 72 else 1)

            add_rect_checked([(card_x1, sy1), (card_x2, sy2)], fill_bgr=bg, outline_bgr=fg, width=3)
            if len(lines) == 1:
                ann.draw_text((card_x1 + 14, sy1 + (step_box_h - step_font) // 2 - 2), lines[0], size=step_font, color_bgr=fg, bold=True)
            else:
                ann.draw_text((card_x1 + 14, sy1 + 8), lines[0], size=step_font, color_bgr=fg, bold=True)
                ann.draw_text((card_x1 + 14, sy1 + 12 + step_font + 4), lines[1], size=step_font, color_bgr=fg, bold=True)

        sum_y1 = desired_top + n_steps * step_pitch + 4
        sum_y2 = sum_y1 + total_box_h
        sum_fg = (0, 145, 0) if sc >= mx else ((180, 100, 0) if sc > 0 else (0, 0, 220))
        sum_bg = (232, 255, 232) if sc >= mx else ((255, 248, 225) if sc > 0 else (255, 232, 232))
        add_rect_checked([(card_x1, sum_y1), (card_x2, sum_y2)], fill_bgr=sum_bg, outline_bgr=sum_fg, width=4)
        extra_ess = " — [SAI BẢN CHẤT -> 0,00đ]" if q_info.get("wrong_in_essence") else ""
        sum_txt = f"[ TỔNG CÂU {q_num}: {sc:.2f}đ / {mx:.2f}đ ]{extra_ess}"
        s_f = 28 if ann.text_width(sum_txt, 28, True) <= text_max_w else 22
        ann.draw_text((card_x1 + 16, sum_y1 + (total_box_h - s_f) // 2 - 2), sum_txt, size=s_f, color_bgr=sum_fg, bold=True)

        if q_info.get("wrong_in_essence"):
            ann.draw_cross_red([(140, y_top_card + 20), (card_x1 - 50, sum_y2 - 10)], color_bgr=(0, 0, 225), width=6)

        end_y = sum_y2
        if has_note_bar:
            ny1 = sum_y2 + 8
            ny2 = ny1 + note_h
            add_rect_checked([(card_x1, ny1), (card_x2, ny2)], fill_bgr=(255, 245, 230), outline_bgr=(0, 90, 220), width=2)
            n_lines = ann.wrap_text(f"Lỗi: {err_note}", text_max_w, size=19, bold=True, max_lines=2 if note_h >= 54 else 1)
            if len(n_lines) == 1:
                ann.draw_text((card_x1 + 12, ny1 + (note_h - 19) // 2 - 2), n_lines[0], size=19, color_bgr=(0, 60, 200), bold=True)
            else:
                ann.draw_text((card_x1 + 12, ny1 + 6), n_lines[0], size=18, color_bgr=(0, 60, 200), bold=True)
                ann.draw_text((card_x1 + 12, ny1 + 28), n_lines[1], size=18, color_bgr=(0, 60, 200), bold=True)
            end_y = ny2
        cursor_y = end_y + 20

    if missing_on_final:
        by = max(cursor_y, h - 45 - len(missing_on_final) * 72)
        for mq in missing_on_final:
            q_num = mq.replace("C", "")
            mx = p4_det[mq]["max_score"]
            add_rect_checked([(card_x1, by), (card_x2, by + 60)], fill_bgr=(255, 235, 235), outline_bgr=(0, 0, 220), width=3)
            ann.draw_text((card_x1 + 16, by + 16), f"[ CÂU {q_num}: 0,00đ / {mx:.2f}đ — Học sinh bỏ trống ]", size=24, color_bgr=(0, 0, 220), bold=True)
            by += 70

    return engine.imwrite_unicode(out_path, ann.to_bgr())


def run_batch():
    base_in = r"D:\TOAN\05_CONG_CU_VA_ENGINE\Cham_Diem_Hoc_Sinh\01_ANH_CHUA_CHAM\Lop_10A17"
    base_out = r"D:\TOAN\05_CONG_CU_VA_ENGINE\Cham_Diem_Hoc_Sinh\02_ANH_DA_CHAM\Lop_10A17"
    os.makedirs(base_out, exist_ok=True)
    for f in glob.glob(os.path.join(base_out, "*.*")):
        os.remove(f)

    students = data_10a17.STUDENTS_10A17
    print(f"Bắt đầu xuất bản bộ ảnh chấm bút đỏ Lớp 10A17 ({len(students)} học sinh)...")
    omr_ok, tl_ok = 0, 0

    with tempfile.TemporaryDirectory() as tmp_dir:
        for s in students:
            stt = s["stt"]
            name_clean = engine.sanitize_filename(s["name"])
            code = s["code"]
            raw_omr = os.path.join(base_in, s["file_omr"])
            in_omr = ensure_upright_omr_file(raw_omr, tmp_dir)
            out_omr = os.path.join(base_out, f"10A17_{stt:02d}_{name_clean}_{code}_DaCham.jpg")
            if code == "CHƯA_TÔ":
                assert annotate_blank_code_omr(in_omr, s, out_omr)
            else:
                assert engine.annotate_omr_sheet(in_omr, s, out_omr), f"OMR failed STT {stt}"
            omr_ok += 1

            tl_pages = s["tl_pages"]
            total_tl = len(tl_pages)
            present_qs = {q for pg in tl_pages for q in pg.get("questions", [])}
            missing_qs = [q for q in ("C19", "C20", "C21") if q not in present_qs]

            for p_idx, pg in enumerate(tl_pages, start=1):
                in_tl = ensure_upright_tl_file(os.path.join(base_in, pg["file"]), tmp_dir)
                suffix = f"_TuLuan_Trang{p_idx}" if total_tl > 1 else "_TuLuan"
                out_tl = os.path.join(base_out, f"10A17_{stt:02d}_{name_clean}_{code}{suffix}_DaCham.jpg")
                m_final = missing_qs if (p_idx == total_tl) else ()
                assert annotate_essay_10a17(in_tl, s, out_tl, page_info=pg, page_idx=p_idx, total_pages=total_tl, missing_on_final=m_final)
                tl_ok += 1

    print(f"HOÀN THÀNH 10A17: {omr_ok} OMR + {tl_ok} Tự luận -> {base_out}")


if __name__ == "__main__":
    run_batch()
