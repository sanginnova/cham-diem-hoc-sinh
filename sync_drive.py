# -*- coding: utf-8 -*-
import urllib.request
import os
import re
import sys
import json

out_dir = r"D:\TOAN\05_CONG_CU_VA_ENGINE\Cham_Diem_Hoc_Sinh\01_ANH_CHUA_CHAM"
os.makedirs(out_dir, exist_ok=True)

if len(sys.argv) > 1 and "drive.google.com" in sys.argv[1]:
    url = sys.argv[1]
else:
    url = "https://drive.google.com/drive/folders/1Ig3MJyBVrLcgN8Yym9oeLurR0XzZc6ue"

print(f"Đang quét Google Drive: {url}")
req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
with urllib.request.urlopen(req) as resp:
    html = resp.read().decode('utf-8', errors='ignore')

# Bóc tách tất cả các file ảnh
files = {}

# Pattern 1: Tìm trong khối data-initial-data hoặc JSON của Google Drive
# Google Drive embed data: [id, [filename,...], mimeType, ...]
for m in re.finditer(r'\"([a-zA-Z0-9_\-]{28,40})\"[,\s]*\[\"([^\"]+\.(?:jpg|jpeg|png|JPG|PNG))\"', html):
    fid, fn = m.group(1), m.group(2)
    if "icon" not in fn.lower() and "logo" not in fn.lower():
        files[fn] = fid

# Pattern 2: Tìm theo tên file và chunk xung quanh
for m in re.finditer(r'(IMG[a-zA-Z0-9_\-]+\.(?:jpg|jpeg|png|JPG|PNG))', html):
    fn = m.group(1)
    if fn not in files:
        chunk = html[max(0, m.start() - 350):min(len(html), m.end() + 350)]
        id_matches = re.findall(r'\"([a-zA-Z0-9_\-]{33})\"', chunk)
        if id_matches:
            files[fn] = id_matches[0]

print(f"Tổng số ảnh quét được trên Drive: {len(files)}")
for fn, fid in sorted(files.items()):
    print(f"  + {fn}: {fid}")

# Tiến hành tải các ảnh mới về out_dir
success = 0
for i, (fn, fid) in enumerate(sorted(files.items()), 1):
    dest = os.path.join(out_dir, fn)
    if os.path.exists(dest) and os.path.getsize(dest) > 100000:
        print(f"[{i}/{len(files)}] Đã có sẵn: {fn} ({os.path.getsize(dest)} bytes)")
        success += 1
        continue
    down_url = f'https://drive.google.com/uc?export=download&id={fid}'
    try:
        r = urllib.request.Request(down_url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(r) as response, open(dest, 'wb') as f:
            f.write(response.read())
        print(f"[{i}/{len(files)}] TẢI THÀNH CÔNG: {fn} ({os.path.getsize(dest)} bytes)")
        success += 1
    except Exception as e:
        print(f"[{i}/{len(files)}] Lỗi khi tải {fn}: {e}")

print(f"\n=> HOÀN TẤT ĐỒNG BỘ: {success}/{len(files)} ảnh!")
