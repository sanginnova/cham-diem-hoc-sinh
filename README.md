# HỆ THỐNG SUB-AGENT CHẤM ĐIỂM BÀI THI TOÁN HỌC (THPT, CĐN01)
> **Chủ nhiệm hệ thống**: ThS. Nguyễn Văn Sang - Khoa Cơ bản - Trường Cao đẳng Nghề số 1 - BQP  
> **Sub-agent**: `cham-diem-hoc-sinh` (Gemini Multimodal Vision & Bilinear 4-Corner Optical OMR)  
> **Tiêu chuẩn khảo thí**: Chương trình GDPT 2018 — Cấu trúc đề 4 phần (Thang điểm 10)

---

## 🌟 TỔNG QUAN & CÔNG CỤ ĐÍNH KÈM

Hệ thống cung cấp giải pháp khép kín từ đồng bộ ảnh bài thi Google Drive, nhận diện kép (**Gemini Vision + Optical Bilinear 4 góc**), chấm bút đỏ sư phạm trực quan, kiểm toán 2 vòng bắt buộc, đến xuất bản báo cáo Excel:

| Công cụ / Script đính kèm | Chức năng chính |
| :--- | :--- |
| **`CHAY_CHAM_DIEM.bat`** | Script chạy trọn gói 1-click trên Windows (Đồng bộ -> Xuất Excel -> Vẽ ảnh bút đỏ -> Double Check -> Dọn rác) |
| **`sync_drive.py`** | Tải và đồng bộ ảnh bài làm từ thư mục Google Drive |
| **`cham_diem_core.py`** | Engine chấm điểm & xuất Excel tổng hợp Lớp 12 (`Ket_Qua_Cham_Diem_Giua_Ky_1_Lop12.xlsx`) |
| **`export_excel_12b6.py`** / **`12b4.py`** / **`10a17.py`** | Xuất báo cáo Excel chi tiết từng lớp kèm Sheet `Đối Chiếu` & `Ảnh Cần Xác Nhận` |
| **`tao_anh_da_cham_12b6.py`** / **`12b4.py`** / **`10a17.py`** / **`tao_bo_anh_da_cham.py`** | Engine vẽ vòng tròn xanh/cam/đỏ Bilinear 4 góc (OMR) + khung barem phân bước 0,25đ (Tự luận) |
| **`double_check_cham_diem.py`** | Hệ thống Kiểm Toán Kép 2 Vòng độc lập (`Vòng 1`: Sĩ số, đáp án, kho ảnh; `Vòng 2`: Pixel phân vùng OMR & Tự luận) |
| **`soat_chep_bai_tu_luan.py`** | Công cụ rà soát phát hiện bài tự luận có dấu hiệu chép nhau / trùng lập luận |
| **`don_dep_he_thong.py`** | Tự động dọn sạch file rác/cache (`__pycache__`, lock Excel, ảnh crop tạm) |

---

## ⚙️ HƯỚNG DẪN CÀI ĐẶT

### 1. Yêu cầu môi trường
- **Python**: `3.10+` (khuyến nghị `3.12`)
- **Git**: Đã cài đặt `git` trên máy
- **Font hệ thống (Windows)**: `C:\Windows\Fonts\arial.ttf` và `arialbd.ttf` (mặc định có sẵn trên Windows để hiển thị tiếng Việt có dấu trên ảnh chấm đỏ)

### 2. Cài đặt nhanh từ GitHub
```bash
# 1. Clone repository về máy
git clone https://github.com/sanginnova/cham-diem-hoc-sinh.git
cd cham-diem-hoc-sinh

# 2. (Khuyến nghị) Tạo môi trường ảo Python
python -m venv .venv
# Windows PowerShell / CMD:
.venv\Scripts\activate

# 3. Cài đặt toàn bộ thư viện & công cụ đính kèm
pip install --upgrade pip
pip install -r requirements.txt
```

*(Hoặc cài nhanh thủ công từng gói: `pip install opencv-python numpy pillow openpyxl gdown python-docx`)*

### 3. Cài đặt Skill cho Antigravity / AI Agent (Tùy chọn)
Sao chép file skill `SKILL.md` vào thư mục `.agents/skills/cham-diem-hoc-sinh/` trong workspace để kích hoạt lệnh `/cham-diem-hoc-sinh`.

---

## 📂 CẤU TRÚC THƯ MỤC CHUẨN

```text
Cham_Diem_Hoc_Sinh/
├── 00_DE_VA_DAP_AN/
│   ├── Khoi_10_GHK1_2026_2027/     # Bộ đề & đáp án Toán 10 (101-104, 201-204)
│   └── Khoi_12_GHK1_2026_2027/     # Bộ đề & đáp án Toán 12 (101-104, 201-204)
├── 01_ANH_CHUA_CHAM/               # Ảnh gốc tải từ Google Drive (ignored trên Git)
├── 02_ANH_DA_CHAM/                 # Ảnh bài làm đã chấm bút đỏ (ignored trên Git)
├── 03_KET_QUA_EXCEL/               # Bảng điểm & đối chiếu chi tiết (.xlsx)
├── data_12b1.py / data_12b4.py / data_12b6.py / data_10a17.py
├── requirements.txt                # Danh sách thư viện Python cần thiết
├── CHAY_CHAM_DIEM.bat              # Chạy 1-click trên Windows
└── docs/samples/                   # Ảnh mẫu bài làm đã chấm đỏ trực quan
```

---

## 🚀 HƯỚNG DẪN SỬ DỤNG

### 1. Chạy 1-Click tự động (Windows)
Nhấp đúp trực tiếp vào:
```cmd
CHAY_CHAM_DIEM.bat
```

### 2. Chạy từng công cụ qua Command Line
```bash
# 1. Đồng bộ ảnh bài làm từ Google Drive (thay link thư mục Drive tương ứng)
python sync_drive.py "https://drive.google.com/drive/folders/<FOLDER_ID>"

# 2. Xuất báo cáo Excel theo từng lớp
python cham_diem_core.py          # Lớp 12B1 + file tổng hợp Khối 12
python export_excel_12b4.py       # Lớp 12B4
python export_excel_12b6.py       # Lớp 12B6
python export_excel_10a17.py      # Lớp 10A17

# 3. Xuất bộ ảnh bài làm đã chấm bút đỏ
python tao_bo_anh_da_cham.py      # Lớp 12B1
python tao_anh_da_cham_12b4.py    # Lớp 12B4
python tao_anh_da_cham_12b6.py    # Lớp 12B6
python tao_anh_da_cham_10a17.py   # Lớp 10A17

# 4. Chạy Kiểm toán Kép 2 Vòng Độc Lập (Double-Check Gate)
python double_check_cham_diem.py Lop_12B1 43
python double_check_cham_diem.py Lop_12B4 36
python double_check_cham_diem.py Lop_12B6 34

# 5. Soát trùng bài tự luận & dọn dẹp rác hệ thống
python soat_chep_bai_tu_luan.py
python don_dep_he_thong.py
```

---

## ⚖️ BẢN QUYỀN
Hệ thống được thiết kế và phát triển bởi **ThS. Nguyễn Văn Sang** - Khoa Cơ bản - Trường Cao đẳng Nghề số 1 - BQP.
