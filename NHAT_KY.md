# Nhật ký thực hiện – Xuất PDF theo file tổng hợp (KSDT MEP)

Dự án: xuất PDF dự toán theo `…\231108_KSDT_MEP_Du toan den ngay 17.3.2023`, đúng thứ tự file tổng.

## Lần 1 – Tạo tính năng xuất theo file tổng (V6.2)
- Đọc link ngoại trong cột G của sheet `TH M&E` (file tổng) → file trung gian → file chi tiết; dòng `THKP hạng mục!$H$n` xác định hạng mục.
- Đặt tên PDF `NN_<STT>_<tên>.pdf`, dùng chung PDF cho mục trỏ cùng hạng mục, ghi `_THU_TU_XUAT.txt`.
- Thêm nút xanh trên sheet + macro `XuatTheoFileTong`; thêm `_gen.ps1` và sửa `BUILD_V6.ps1` để build không cần Python.
- Kết quả chạy đầu: 37 PDF, 4 mục không xuất (không có file chi tiết).

## Lần 2 – Sửa theo phản hồi (V6.3)
- Phát hiện: PDF cùng file nguồn nặng bằng nhau → Giá tổng hợp / Đơn giá chi tiết bị in nguyên file. Nguyên nhân: hai sheet này không có dòng "HẠNG MỤC:", engine không nhận ra khối.
- Sửa: nhận khối theo dòng tiêu đề lặp lại; khớp hạng mục với THKP (tên chính xác → khớp duy nhất → loại TB/XD → theo thứ tự).
- Giá tháng: nhiều file ẩn sheet này nên bị bỏ → tự hiện lại; đưa vào PDF của mọi hạng mục.
- Bỏ in cột Tổng cộng ở Giá tổng hợp.
- Thêm bộ lọc theo thư mục `API_SetTongFilter`. Xuất thử ĐHKK: 9 PDF (02: 8 MB → 1,4 MB; 03: 8 MB → 2,2 MB).
- Xuất các thư mục còn lại: ĐIỆN 7, ĐIỆN NHẸ 16, CTN 5, Chống cháy lan 0. Cảnh báo ở ĐIỆN: 4 mục không có khối Đơn giá chi tiết cho hạng mục thứ 3.

## Lần 3 – Bỏ Đơn giá chi tiết (V6.4)
- Thêm `API_SetTongNoDGCT`. Xuất lại toàn bộ vào `PDF\KHONG_DGCT`: 37 PDF (0,3–3,5 MB/file), cảnh báo ở ĐIỆN hết.
- 4 mục vẫn không xuất: ngăn cháy lan (23), xử lý nước thải (37), Onsen (38), ống lồng xuyên sàn (39).

## Lần 4 – Đưa lên GitHub
- Cài Git 2.56 và GitHub CLI 2.103.3, đăng nhập `gh`, tạo repo public `tonghang0111-rgb/In-an-G8`, đẩy 11 file (mã nguồn, build script, bản V6.4).

## Chưa làm / hạn chế
- Chưa mở PDF để kiểm tra từng trang (môi trường không có trình xem PDF) – cần người dùng duyệt mẫu.
- Tuỳ chọn bỏ Đơn giá chi tiết và lọc theo thư mục chưa có nút trên giao diện.
- Form `frmInDuToan` giữ nguyên (VBE Designer báo lỗi 75 khi dựng lại bằng code).
