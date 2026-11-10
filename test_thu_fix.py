# -*- coding: utf-8 -*-
import os
os.environ['OPENBLAS_NUM_THREADS'] = '1'
import cv2

# Exact coordinates for Batch 2 (11/09)
PARAMS_B2_EXACT = {
    "p1_x1": 1740, "p1_y1": 1675, "p1_dx": 90, "p1_dy": 68,
    "p1_x11": 2220,
    "p2_c13_d": 1050, "p2_c13_s": 1130,
    "p2_c14_d": 1290, "p2_c14_s": 1370,
    "p2_y": [2540, 2605, 2670, 2735],
    "r": 28
}

KEY_101 = {
    "p1": ["D", "B", "D", "D", "D", "B", "C", "A", "B", "D", "B", "A"],
    "p2_13": ["S", "Đ", "Đ", "S"],
    "p2_14": ["Đ", "Đ", "S", "S"]
}

img_path = r'D:\TOAN\05_CONG_CU_VA_ENGINE\Cham_Diem_Hoc_Sinh\anh_bai_lam\IMG20261109183158.jpg'
img = cv2.imread(img_path)
params = PARAMS_B2_EXACT

# Draw P1
for q in range(12):
    ans = KEY_101["p1"][q]
    col_idx = {"A": 0, "B": 1, "C": 2, "D": 3}[ans]
    if q < 10:
        cx = params["p1_x1"] + col_idx * params["p1_dx"]
        cy = params["p1_y1"] + q * params["p1_dy"]
    else:
        cx = params["p1_x11"] + col_idx * params["p1_dx"]
        cy = params["p1_y1"] + (q - 10) * params["p1_dy"]
    cv2.circle(img, (cx, cy), params["r"], (0, 0, 255), 4)

# Draw P2
for i, ans in enumerate(KEY_101["p2_13"]):
    x = params["p2_c13_d"] if ans == "Đ" else params["p2_c13_s"]
    y = params["p2_y"][i]
    cv2.circle(img, (x, y), params["r"], (0, 0, 255), 4)

for i, ans in enumerate(KEY_101["p2_14"]):
    x = params["p2_c14_d"] if ans == "Đ" else params["p2_c14_s"]
    y = params["p2_y"][i]
    cv2.circle(img, (x, y), params["r"], (0, 0, 255), 4)

# Crop P1 to check
p1_crop = img[1550:2450, 1600:2550]
cv2.imwrite(r'C:\Users\ADMIN\.gemini\antigravity\brain\9fd41f44-9d04-4d5f-a91c-0f6d49e69d05\check_p1_perfect.jpg', p1_crop)
print("Saved check_p1_perfect.jpg")
