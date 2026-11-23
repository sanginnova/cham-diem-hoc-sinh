@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ===============================================================================
echo   HỆ THỐNG TỰ ĐỘNG CHẤM ĐIỂM BÀI THI TOÁN 12 - THS. NGUYỄN VĂN SANG
echo ===============================================================================
echo.
echo [1/5] Đang đồng bộ ảnh bài làm mới nhất từ Google Drive vào 01_ANH_CHUA_CHAM...
python sync_drive.py

echo.
echo [2/5] Đang chấm điểm và kết xuất Báo cáo Excel vào 03_KET_QUA_EXCEL...
python cham_diem_core.py
python export_excel_12b4.py
python export_excel_12b6.py
python export_excel_10a17.py

echo.
echo [3/5] Đang xuất bản bộ ảnh chấm bút đỏ vào 02_ANH_DA_CHAM...
python tao_bo_anh_da_cham.py
python tao_anh_da_cham_12b4.py
python tao_anh_da_cham_12b6.py
python tao_anh_da_cham_10a17.py

echo.
echo [4/5] Đang chạy Kiểm Toán Kép 2 Vòng Độc Lập (Double-Check Gate)...
python double_check_cham_diem.py Lop_12B1 43
if errorlevel 1 (
    echo [CHẶN BÀN GIAO] Kiểm toán Lop_12B1 thất bại! Vui lòng kiểm tra lỗi ở trên.
    pause
    exit /b 1
)
python double_check_cham_diem.py Lop_12B4 36
if errorlevel 1 (
    echo [CHẶN BÀN GIAO] Kiểm toán Lop_12B4 thất bại! Vui lòng kiểm tra lỗi ở trên.
    pause
    exit /b 1
)
python double_check_cham_diem.py Lop_12B6 34
if errorlevel 1 (
    echo [CHẶN BÀN GIAO] Kiểm toán Lop_12B6 thất bại! Vui lòng kiểm tra lỗi ở trên.
    pause
    exit /b 1
)

echo.
echo [5/5] Đang dọn dẹp toàn bộ file rác và giải phóng bộ nhớ...
python don_dep_he_thong.py

echo.
echo ===============================================================================
echo   CHẤM ĐIỂM & KIỂM TOÁN THÀNH CÔNG 100%%! ĐANG MỞ FILE BÁO CÁO EXCEL...
echo ===============================================================================
start "" "03_KET_QUA_EXCEL\Ket_Qua_Cham_Diem_Giua_Ky_1_Lop12.xlsx"
pause
