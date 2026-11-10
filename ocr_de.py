# -*- coding: utf-8 -*-
import cv2
import os
import pytesseract

folder = r'D:\TOAN\05_CONG_CU_VA_ENGINE\Cham_Diem_Hoc_Sinh\anh_bai_lam\De kiem tra giua ki 1'
files = sorted([f for f in os.listdir(folder) if f.endswith('.jpg')])

for idx, f in enumerate(files):
    p = os.path.join(folder, f)
    img = cv2.imread(p)
    rot = cv2.rotate(img, cv2.ROTATE_90_COUNTERCLOCKWISE)
    h, w = rot.shape[:2]
    foot = rot[int(h*0.85):, :]
    head = rot[:int(h*0.15), :]
    txt_f = pytesseract.image_to_string(foot)
    txt_h = pytesseract.image_to_string(head)
    combined = txt_h + '\n' + txt_f
    summary = []
    for line in combined.splitlines():
        line = line.strip()
        if any(k in line.lower() for k in ['mã đề', 'trang', 'phần', '101', '102', '103', '104']):
            summary.append(line)
    res = " ; ".join(summary)
    print(f"{idx+1:2d} | {f} | {res}")

