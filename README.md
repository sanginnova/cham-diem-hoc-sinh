# Hệ Thống Chấm Điểm Bài Thi Đa Môn Chuẩn GDPT 2018 (OMR + Tự Luận + Excel)

Bộ công cụ tự động hóa chấm phiếu trắc nghiệm (Phần I, II, III), chấm bút đỏ sư phạm trên ảnh bài làm, sắp xếp kết quả theo bảng chữ cái tiếng Việt, xuất báo cáo Excel chi tiết và kiểm toán kép 2 vòng. Áp dụng được cho **mọi môn học và mọi bộ đề** theo cấu trúc GDPT 2018.

---

## 1. Công cụ đi kèm trong Repo

| Tệp / Công cụ | Vai trò chính |
| :--- | :--- |
| `CHAY_CHAM_DIEM.bat` | Kịch bản chạy trọn gói 1-click trên Windows (đồng bộ Drive → xuất Excel → vẽ ảnh bút đỏ → double-check → dọn rác). |
| `data_template.py` | **File mẫu cấu hình**: điền đáp án chuẩn `MASTER_KEYS` và danh sách bài làm cho môn học / mã đề của bạn. |
| `common_utils.py` | Thư viện dùng chung: hàm sắp xếp tên tiếng Việt (`vn_name_sort_key`) và cấu hình đáp án mẫu. |
| `sync_drive.py` | Đồng bộ ảnh bài làm từ Google Drive về thư mục `01_ANH_CHUA_CHAM/`. |
| `cham_diem_core.py` | Engine tính điểm theo thang GDPT 2018, sắp xếp alphabe và xuất Excel (Bảng điểm + Đối chiếu chi tiết từng câu). |
| `double_check_cham_diem.py` | Kiểm toán kép 2 vòng: tái chấm chéo đáp án, đối soát sĩ số/ảnh gốc & kiểm tra phân vùng ảnh chấm đỏ. |
| `don_dep_he_thong.py` | Dọn dẹp file tạm, `__pycache__` sau khi xuất bản. |
| `index.html` | Giao diện web **SmartOMR** chấm nhanh qua camera trình duyệt. |
| `requirements.txt` | Danh sách thư viện Python cần cài đặt. |

---

## 2. Yêu cầu môi trường & Cài đặt

- **Python 3.10+** (khuyến nghị Python 3.11 / 3.12), **Git**
- Hệ điều hành Windows (có sẵn font `arial.ttf` / `arialbd.ttf` trong `C:\Windows\Fonts`)

```bash
# 1. Tải mã nguồn về máy
git clone https://github.com/sanginnova/cham-diem-hoc-sinh.git
cd cham-diem-hoc-sinh

# 2. (Khuyên dùng) Tạo môi trường ảo Python
python -m venv .venv
.venv\Scripts\activate

# 3. Cài đặt toàn bộ thư viện phụ thuộc
pip install -r requirements.txt
```

---

## 3. Cách sử dụng cho môn học / đề thi khác

1. **Chuẩn bị thư mục dữ liệu cục bộ** (các thư mục này được `.gitignore` giữ riêng trên máy bạn, không đẩy lên GitHub):
   - `00_DE_VA_DAP_AN/`: Lưu đề thi và đáp án gốc của bạn.
   - `01_ANH_CHUA_CHAM/<Ten_Lop>/`: Chứa ảnh phiếu OMR và bài tự luận học sinh.
2. **Cấu hình đáp án của bạn**:
   - Sao chép `data_template.py` thành `data_<ten_lop>.py`, điền bảng đáp án `MASTER_KEYS` cho các mã đề của môn bạn dạy (Toán, Lý, Hóa, Sinh, Sử, Địa, GDKTPL...).
3. **Chạy chấm điểm & kiểm toán**:
   ```bash
   python cham_diem_core.py
   python double_check_cham_diem.py <Ten_Lop> <Si_So>
   python don_dep_he_thong.py
   ```
   Hoặc chạy trực tiếp `CHAY_CHAM_DIEM.bat`. Kết quả xuất vào `03_KET_QUA_EXCEL/` và ảnh chấm bút đỏ lưu trong `02_ANH_DA_CHAM/`.
