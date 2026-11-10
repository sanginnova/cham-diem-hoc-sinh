# -*- coding: utf-8 -*-
import urllib.request
import os
import re
import sys

out_dir = r"D:\TOAN\05_CONG_CU_VA_ENGINE\Cham_Diem_Hoc_Sinh\anh_bai_lam"
os.makedirs(out_dir, exist_ok=True)

url = 'https://drive.google.com/drive/folders/1cULmlbvqnVm1jjXrioGwnTi-vb1Jrmvi'
req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
print("Đang quét Google Drive...")
with urllib.request.urlopen(req) as resp:
    html = resp.read().decode('utf-8', errors='ignore')

# Bóc tách tất cả các file IMG2026*.jpg
files = {}
for m in re.finditer(r'(IMG2026\d+\.jpg)', html):
    fn = m.group(1)
    chunk = html[max(0, m.start() - 350):m.end() + 350]
    # Tìm chuỗi file ID 33 ký tự
    id_matches = re.findall(r'\"([a-zA-Z0-9_\-]{33})\"', chunk)
    if id_matches:
        files[fn] = id_matches[0]

print(f"Tổng số ảnh quét được trên Drive: {len(files)}")

success = 0
for i, (fn, fid) in enumerate(sorted(files.items()), 1):
    dest = os.path.join(out_dir, fn)
    if os.path.exists(dest) and os.path.getsize(dest) > 100000:
        print(f"[{i}/{len(files)}] Đã có sẵn: {fn}")
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
