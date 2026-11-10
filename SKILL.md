---
name: cham-diem-hoc-sinh
description: >-
  Kỹ năng & Sub-agent chuyên sâu tự động hóa Chấm điểm bài thi Toán học (THPT, CĐN01) từ ảnh chụp bài làm học sinh lưu trữ trên Google Drive: Tự động tải ảnh, phân tích nhận diện Họ tên, Lớp, Mã đề, chấm trắc nghiệm Phần I (12 câu ABCD), Phần II (Đúng/Sai 2 câu 8 ý), Phần III (Trả lời ngắn 4 câu), chấm điểm tự luận Phần IV (3 câu), tính tổng điểm thang 10, phân loại học lực và xuất Báo cáo Excel chuyên nghiệp (.xlsx) kèm lỗi sai chi tiết từng câu. Tự động đồng bộ và cập nhật khi giáo viên tải thêm ảnh vào thư mục Drive.
---

# KỸ NĂNG TỰ ĐỘNG CHẤM ĐIỂM BÀI THI TOÁN TỪ GOOGLE DRIVE (CHAM-DIEM-HOC-SINH)
> **Chủ nhiệm hệ thống**: ThS. Nguyễn Văn Sang - Khoa Cơ bản - Trường Cao đẳng Nghề số 1 - BQP.  
> **GitHub Repository**: [https://github.com/sanginnova/cham-diem-hoc-sinh](https://github.com/sanginnova/cham-diem-hoc-sinh)  
> **Lệnh kích hoạt**: `/cham-diem`, `/cham-bai`, hoặc câu lệnh: *"chấm điểm học sinh từ link drive"*, *"tổng hợp điểm và lỗi sai"*, *"cập nhật thêm bài làm mới"*.

---

## 1. TỔNG QUAN HỆ THỐNG

Hệ thống cung cấp giải pháp khép kín 100% từ khâu thu thập ảnh bài thi từ Google Drive đến khâu phân tích dữ liệu, đối chiếu trực tiếp với 4 bộ đề in gốc thực tế (101, 102, 103, 104) và xuất bản Báo cáo Excel chuẩn Sư phạm.

### Thang điểm chuẩn GDPT 2018:
- **Phần I (3.0 điểm)**: 12 câu trắc nghiệm 4 lựa chọn ABCD (0.25đ / câu).
- **Phần II (2.0 điểm)**: 2 câu Đúng/Sai (Câu 13, 14), mỗi câu 4 ý:
  * Đúng 1 ý: 0.25 điểm.
  * Đúng 2 ý: 0.50 điểm.
  * Đúng 3 ý: 0.75 điểm.
  * Đúng 4 ý: 1.00 điểm.
- **Phần III (2.0 điểm)**: 4 câu Trả lời ngắn (Câu 15, 16, 17, 18), mỗi câu 0.50 điểm.
- **Phần IV (3.0 điểm)**: 3 câu Tự luận (Câu 19, 20, 21), mỗi câu 1.00 điểm theo barem phân bước.

---

## 2. BẢNG ĐÁP ÁN CHUẨN XÁC 100% CỦA 4 BỘ ĐỀ IN THỰC TẾ

### Mã đề 101:
- **Phần I**: 1.D, 2.B, 3.D, 4.D, 5.D, 6.B, 7.C, 8.A, 9.B, 10.D, 11.B, 12.A
- **Phần II**: C13: a.S, b.Đ, c.Đ, d.S | C14: a.Đ, b.Đ, c.S, d.S
- **Phần III**: C15: 1 | C16: 18 | C17: 2 | C18: 7
- **Phần IV**: C19 (m=3, b=1, c=-3); C20 (v_max=25 m/s khi t=4s); C21 (x ≈ 24.3m)

### Mã đề 102:
- **Phần I**: 1.B, 2.B, 3.C, 4.D, 5.C, 6.B, 7.C, 8.B, 9.B, 10.C, 11.D, 12.D
- **Phần II**: C13: a.Đ, b.S, c.Đ, d.Đ | C14: a.Đ, b.S, c.Đ, d.S
- **Phần III**: C15: 7 | C16: 18 | C17: 1 | C18: 2
- **Phần IV**: C19, C20, C21 tương tự mã 101.

### Mã đề 103:
- **Phần I**: 1.B, 2.A, 3.B, 4.D, 5.C, 6.A, 7.A, 8.D, 9.B, 10.C, 11.A, 12.C
- **Phần II**: C13: a.Đ, b.Đ, c.S, d.S | C14: a.S, b.Đ, c.Đ, d.S
- **Phần III**: C15: 18 | C16: 2 | C17: 1 | C18: 7
- **Phần IV**: C19, C20, C21 tương tự mã 101.

### Mã đề 104:
- **Phần I**: 1.D, 2.A, 3.B, 4.A, 5.D, 6.C, 7.A, 8.D, 9.C, 10.C, 11.D, 12.D
- **Phần II**: C13: a.Đ, b.Đ, c.S, d.S | C14: a.S, b.Đ, c.Đ, d.S
- **Phần III**: C15: 2 | C16: 1 | C17: 7 | C18: 18
- **Phần IV**: C19, C20, C21 tương tự mã 101.

---

## 3. QUY TRÌNH KIỂM SOÁT THỊ GIÁC BẮT BUỘC BẰNG GEMINI VISION (ZERO ERROR GATE)

Để loại trừ 100% sai sót của các thuật toán OpenCV cổ điển (bị lệch tọa độ, bóng mờ, vết tẩy chì), hệ thống **BẮT BUỘC SỬ DỤNG GEMINI MULTIMODAL VISION** để đọc trực tiếp từng ảnh bài thi gốc phân giải cao:

1. **Nhận diện Thông tin thí sinh & Mã đề thi**:
   - Gemini Vision đọc trực tiếp vùng tiêu đề viết tay: Họ và tên, Lớp, Số báo danh.
   - Đọc đối chiếu kép cả ô số viết tay và 3 cột ô tròn OMR mã đề (`0-9`).
   - ⚠️ **KỶ LUẬT KHẢO THÍ SỐ 1**: **Học sinh viết và tô mã đề nào thì BẮT BUỘC chấm duy nhất và tuyệt đối theo đúng mã đề đó**. Tuyệt đối không được tự ý sửa, thay đổi hay suy đoán mã đề của học sinh!

2. **Soi chiếu Phần I (12 câu trắc nghiệm ABCD)**:
   - Gemini Vision trực tiếp phân tích độ tương phản quang học giữa vết chì tô thật và vết tẩy xóa mờ.
   - Trích xuất chính xác 100% mảng phương án học sinh chọn (ví dụ: `1-A, 2-D, 3-A, 4-D, 5-D, 6-D, 7-D, 8-D, 9-C, 10-B, 11-A, 12-C`).
   - Đối chiếu với đáp án chuẩn của Mã đề học sinh tô, tính điểm chính xác 0.25đ / câu đúng.

3. **Soi chiếu Phần II (Đúng/Sai) & Phần III (Điền số)**:
   - Phần II: Đọc chính xác 8 ý Đúng/Sai (Câu 13: a,b,c,d; Câu 14: a,b,c,d), tính điểm phân bước chuẩn Bộ: 1 ý = 0.1đ; 2 ý = 0.25đ; 3 ý = 0.5đ; 4 ý = 1.0đ.
   - Phần III: Đọc cả ô số viết tay và các ô tròn tô bên dưới cho 4 câu (C15, C16, C17, C18).

4. **Chấm tự luận Phần IV (Câu 19, 20, 21)**:
   - Gemini Vision đọc từng bước biến đổi toán học viết tay trên giấy thi, chấm chuẩn theo barem điểm từng bước (0.25đ, 0.5đ, 0.75đ, 1.0đ).

5. **Xuất bản Bảng đối chiếu chi tiết từng câu trong Excel**:
   - File Excel báo cáo [Ket_Qua_Cham_Diem_Giua_Ky_1_Lop12.xlsx](file:///D:/TOAN/05_CONG_CU_VA_ENGINE/Cham_Diem_Hoc_Sinh/Ket_Qua_Cham_Diem_Giua_Ky_1_Lop12.xlsx) bắt buộc phải có **Sheet `Đối Chiếu Bài Làm Chi Tiết`**:
     * Hiển thị rõ từng phương án học sinh chọn ở từng câu (C1..C12, 13a..14d, C15..C18, C19..C21).
     * Tự động highlight màu: 🟢 Xanh lá nếu học sinh làm đúng; 🔴 Hồng đỏ nếu học sinh làm sai.
     * Cung cấp hàng đáp án chuẩn 4 mã đề ở đầu bảng để Giáo viên đối soát trực quan ngay lập tức.

---

## 4. QUY CHUẨN XỬ LÝ ĐA TỜ & BẢO TOÀN 100% ẢNH CHỤP (ZERO-LEAK & UNKNOWN HANDLER)

Một học sinh có thể làm nhiều tờ và giáo viên có thể chụp tách rời từng tờ. Hệ thống thiết lập quy chuẩn xử lý bảo toàn tuyệt đối:

1. **Ghép tự động nhiều tờ của cùng một học sinh (Multi-page Grouping)**:
   - Dựa vào **Họ tên viết tay**, **Lớp**, **Số báo danh** và **Dấu thời gian chụp ảnh (Timestamp)** để tự động gộp các ảnh rời rạc vào một hồ sơ học sinh hoàn chỉnh:
     * **Tờ 1**: Phiếu trắc nghiệm OMR (Phần I: 12 câu ABCD; Phần II: Đúng/Sai; Phần III: Điền số).
     * **Tờ 2**: Tờ giấy thi tự luận trang 1 (Câu 19: Khảo sát hàm số; Câu 20: Chuyển động vật lý).
     * **Tờ 3 (nếu có)**: Tờ giấy thi tự luận trang 2 (Câu 21: Tối ưu hóa thực tế).

2. **Kỷ luật "Quét sạch 100% - Tuyệt đối không bỏ sót bất kỳ ảnh nào"**:
   - Nếu thư mục Drive có 50 ảnh thì hệ thống phải **kiểm kê và xử lý đủ 50 ảnh**. Có 74 ảnh thì phải **xử lý đủ 74 ảnh**.
   - Cấm mọi hành vi bỏ qua, đếm thiếu hoặc chỉ chấm các bài dễ nhận diện mà bỏ sót các ảnh còn lại.

3. **Cơ chế quản lý ảnh không rõ danh tính (Ảnh không biết bài của ai)**:
   - Đối với các ảnh bị mờ họ tên, không điền SBD, hoặc ảnh chụp trang 2 tự luận rời mà học sinh không ghi tên đầu trang:
     * **Tuyệt đối KHÔNG ĐƯỢC BỎ QUA hay xóa bỏ**.
     * **Bắt buộc xuất thành mục riêng trong File Báo cáo Excel**:
       - Liệt kê chính xác **Tên file ảnh gốc** (ví dụ: `IMG20261107115121.jpg`).
       - Ghi rõ nội dung toán học Gemini Vision đọc được trên tờ giấy (ví dụ: *"Tự luận C20 giải đúng vmax=25, C21 giải đúng x=24.3m, điểm thành phần 2.0đ"*).
       - Đánh dấu trạng thái nổi bật màu cam/đỏ: ⚠️ **`ẢNH CẦN XÁC NHẬN CHỦ SỞ HỮU`** để Giáo viên đối chiếu trực tiếp với học sinh trên lớp.

---

## 5. QUY CHUẨN CHẤM BÚT ĐỎ SƯ PHẠM TRỰC QUAN (VISUAL RED-PEN GRADING STANDARD)

Nhằm đảm bảo tính minh bạch, trực quan 100% giúp Giáo viên và Học sinh kiểm tra đối soát bài làm chỉ trong 1 giây, hệ thống bắt buộc áp dụng **Bộ quy chuẩn Bút Đỏ Sư Phạm** trên toàn bộ ảnh bài làm:

1. **Phiếu Trắc nghiệm (Phần I, II, III)**:
   - **Khoanh tròn đỏ rực rỡ (BGR: `0, 0, 255`, thickness 4-5px)** vào **TẤT CẢ các đáp án đúng** của bộ đề (101, 102, 103, 104) tương ứng với mã đề học sinh tô.
   - **Cơ chế nhận diện tức thì**:
     * Câu ĐÚNG: Vòng tròn đỏ ôm khít lấy ô chì học sinh tô đen.
     * Câu SAI: Vòng tròn đỏ nằm ở ô đáp án chuẩn, còn ô chì học sinh tô nằm ở ô khác $\implies$ Phát hiện ngay lỗi sai mà không cần tra cứu bảng đáp án rời.
   - **Ghi điểm thành phần bằng MỰC ĐỎ** ngay trên đầu phiếu hoặc lề bài thi:
     * `Phần I: X.XX / 3.00đ` (Số câu đúng/12).
     * `Phần II: X.XX / 2.00đ` (Chi tiết số ý đúng từng câu C13, C14).
     * `Phần III: X.XX / 2.00đ` (Số câu điền đúng/4).
     * `TỔNG TRẮC NGHIỆM: X.XX / 7.00đ`.

2. **Bài thi Tự luận (Phần IV - Câu 19, 20, 21)**:
   - **Chấm như Thầy Cô giáo chấm tay trên lớp**: Soi từng dòng biến đổi toán học viết tay.
   - **Barem phân bước đến 0,25đ ngay cạnh bài làm**:
     * Ghi rõ `+0,25đ [✓]` bằng mực đỏ ngay cạnh dòng/bước làm bài ở lề phải cho từng bước đạt điểm chuẩn.
     * Nếu sai hoặc thiếu: ghi `0,00đ [✗]` kèm chỉ dẫn cụ thể (ví dụ: `0,00đ [✗] (Thiếu vmax)`).
     * Đóng khung đỏ tổng điểm từng câu ngay cạnh lề bài làm: `[ TỔNG C19: 1,00đ ]`, `[ TỔNG C20: 1,00đ ]`, `[ TỔNG C21: 1,00đ ]`.
   - **Chỉ rõ vị trí sai bằng MỰC ĐỎ**:
     * Gạch chân màu đỏ / Khoanh tròn đỏ tại vị trí học sinh viết sai (sai dấu khi đạo hàm, tính nhầm giá trị $y$, thiếu điều kiện $t \ge 0$, chưa kết luận $v_{\max}$, chưa làm bài).
     * Viết ghi chú ngắn gọn bằng chữ màu đỏ bên cạnh dòng sai: `[Sai dấu]`, `[Tính nhầm]`, `[Thiếu KL vmax]`, `[Chưa làm C21]`.
   - **Ghi điểm thành phần từng câu bằng MỰC ĐỎ**:
     * `C19: X.X đ / 1.0đ  |  C20: X.X đ / 1.0đ  |  C21: X.X đ / 1.0đ`
     * `TỔNG TỰ LUẬN: X.XX / 3.00đ  ==>  TỔNG ĐIỂM THI: X.XX / 10.00đ  [XẾP LOẠI]`.

3. **Bàn giao bộ ảnh đã chấm (Atomic Delivery)**:
   - Toàn bộ ảnh bài làm đã chấm đỏ được tự động lưu vào thư mục:  
     `D:\TOAN\05_CONG_CU_VA_ENGINE\Cham_Diem_Hoc_Sinh\anh_da_cham\`
   - Cú pháp đặt tên file chuẩn:
     * Trắc nghiệm: `[STT]_[Họ Tên]_[Mã Đề]_TN_DaCham.jpg`
     * Tự luận: `[STT]_[Họ Tên]_[Mã Đề]_TL_DaCham.jpg`

---

## 6. SUB-AGENT CHUYÊN TRÁCH: GEMINI-VISION-GRADER

Hệ thống đã đóng gói toàn diện 5 Tầng Giải Pháp thành Sub-agent chuyên trách độc lập: **`gemini-vision-grader`**.
* **Cơ chế hoạt động**:
  1. Khi nhận lệnh chấm thi, Sub-agent được kích hoạt trong không gian làm việc cô lập để chống hoàn toàn mọi ảo giác ngữ cảnh.
  2. Sub-agent trực tiếp nhận ảnh phân giải cao, sử dụng mắt thần thị giác Gemini Vision bóc tách từng ô tô chì và chữ viết tay.
  3. Khóa cứng mã đề học sinh tô, chạy bộ lọc kiểm toán kép (Assertion Check: $\text{Total} \equiv P1 + P2 + P3 + P4$).
  4. Gom nhóm đa tờ (trắc nghiệm + tự luận) và kiểm kê 100% số lượng ảnh, không bỏ sót bất kỳ tệp nào.
  5. Xuất bản File Báo cáo Excel 4 Sheet chuẩn Sư phạm kèm Sheet Đối Chiếu Chi Tiết tô màu Xanh/Đỏ từng câu.

---

## 6. FILE CHẠY TỰ ĐỘNG 1-CLICK DÀNH CHO GIÁO VIÊN
Giáo viên nhấp đúp trực tiếp vào file:
`D:\TOAN\05_CONG_CU_VA_ENGINE\Cham_Diem_Hoc_Sinh\CHAY_CHAM_DIEM.bat`
Hệ thống sẽ tự động đồng bộ ảnh bài làm mới từ Drive, kích hoạt Sub-agent Gemini Vision kiểm toán và cập nhật file Excel [Ket_Qua_Cham_Diem_Giua_Ky_1_Lop12.xlsx](file:///D:/TOAN/05_CONG_CU_VA_ENGINE/Cham_Diem_Hoc_Sinh/Ket_Qua_Cham_Diem_Giua_Ky_1_Lop12.xlsx).
