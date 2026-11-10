@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ===============================================================================
echo   HỆ THỐNG TỰ ĐỘNG CHẤM ĐIỂM BÀI THI TOÁN 12 - THS. NGUYỄN VĂN SANG
echo ===============================================================================
echo.
echo [1/3] Đang đồng bộ ảnh bài làm mới nhất từ Google Drive...
python sync_drive.py

echo.
echo [2/3] Đang chấm điểm và kết xuất Báo cáo Excel...
python cham_diem_core.py

echo.
echo [3/3] Hoàn tất! Đang mở file Excel kết quả...
start "" "Ket_Qua_Cham_Diem_Giua_Ky_1_Lop12.xlsx"

echo.
echo ===============================================================================
echo   CHẤM ĐIỂM THÀNH CÔNG! BÁO CÁO ĐÃ ĐƯỢC CẬP NHẬT.
echo ===============================================================================
pause
