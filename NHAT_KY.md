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

## Lần 5 – GĐ1: link nhiều cấp + thư mục con (V6.5, 04/11/2026)
Dự án thử: `M:\DUTOAN\2026\01. BT2.MAG\...\260724. THDT LT\0. TH LUONG TAI.xlsx` (sheet `XD TB`, cột H, 53 dòng link).
- Phân tích: V6.4 chỉ đọc cột F/E của dòng trong file trung gian nên bỏ toàn bộ mục đi qua ô
  `SUBTOTAL`/`SUM` (Điện, TXLNT, trạm bơm, cảnh quan) và mục link vào `TH chi phí TB`/sheet báo giá.
- Viết lại phần xuất theo file tổng: đọc đúng ô được link, đi xuống `SUBTOTAL`/`SUM`/`+` tới 6 cấp,
  phân loại THKP / TH chi phí TB / TH chi phí XD (theo tên hạng mục) / in nguyên sheet.
- 1 dòng file tổng → nhiều PDF `NN.kk_`; link vào dòng tổng của file nhiều hạng mục → `_HMk.pdf`.
- PDF đặt trong thư mục con giống thư mục nguồn (`6. TH Dien\`, `TXLNT LUONG TAI\Excel\`, …).
- Thêm `API_SetTongDryRun` (chỉ dò link, ~15–30 giây).
- Kết quả chạy đầy đủ vào `PDF_TEST`: 53 mục → 103 mục in, xuất 102 PDF (+ file `_HMk`), 1 mục dùng chung PDF,
  0 lỗi, ~30 phút. File 6 hạng mục (`2c/2d Hoan tra - DZ`) tách đúng từng hạng mục.

## Lần 6 – Tên dự án, chọn sheet in (V6.5, 04/11/2026)
- TH chi phí TB chỉ in khi tổng thiết bị khác 0 (San nền, Kè… không còn trang TB trống).
- THKP: bỏ ẩn CỐ ĐỊNH dòng 2, 5, 25 — đã làm mất dòng Tên dự án (khối 1), dòng tiêu đề STT ở file
  9.1 (tiêu đề ở dòng 5), dòng LÀM TRÒN / Bằng chữ. Nay chỉ ẩn dòng "Thời điểm lập" khi trống.
- Thay tên dự án / tên công trình khi in (`API_SetTongTen`), áp vào mọi sheet, mọi khối hạng mục, chỉ trong bộ nhớ.
- Kiểm tra chức năng đổi tên cũ: hàm `ApplyRename` có sẵn nhưng KHÔNG được gọi ở đâu (chết); còn gộp
  "Tên công trình" vào "Dự án" và đổi nhãn thành "DỰ ÁN:" — không dùng lại, thay bằng `TongApplyNames`.
- Tên hạng mục ở Giá tổng hợp / Đơn giá CT: mỗi khối hạng mục trong file mẫu (4. TNT, 2c Điện) có tiêu đề
  riêng; khối 1 lấy `'Công trình'!F6` = hạng mục đầu → in tách theo khối vẫn đúng tên.
- Sheet cấu hình `IN THEO FILE TONG` + `API_SetTongSheets`: chọn sheet + thứ tự cho file 1 hạng mục / nhiều hạng mục.
- Thử: TNT (2 HM) chỉ THKP + Giá TH + Giá tháng: 83 → 23 trang.

## Lần 7 – GĐ2: THKP nhận form theo tiêu đề (04/11/2026)
- Viết HOA dòng Tên dự án / công trình / hạng mục ở THKP; Giá tháng ẩn cột Mã số (2 tuỳ chọn mặc định Có).
- `THKPLayout`: dòng tiêu đề = dòng có cột A "STT"; cột cuối = ô có chữ cuối cùng trên dòng tiêu đề.
  Form `.xls` DM38 → in tới cột **I "KÝ HIỆU"** (trước bị cắt ở H); form `.xlsx` cây xanh → H.
- Dòng lặp tiêu đề theo dòng STT thực tế (file 9.1 ở dòng 5, trước cố định `$6:$6`).
- Lấy thêm dòng "Bằng chữ" ngay dưới bảng THKP.
- Áp cho cả nhánh 1 hạng mục và nhiều hạng mục (mỗi khối nhận riêng). Kiểm tra trên San nền, Cây xanh,
  9.1, 4. TNT (2 HM), 2c Điện (6 HM) – đúng; xuất thử 11 PDF vào `PDF_TEST\_GD2`.

## Lần 8 – Chạy đủ 53 mục qua nút, ẩn số 0, hộp thoại tiếng Việt (04/11/2026)
- Chạy qua macro nút ▶ BẮT ĐẦU XUẤT (Excel hiển thị) vào `PDF_V6.5`: 53 mục → 103 mục in, 102 PDF,
  0 lỗi, 33 phút (bản trước khi có ẩn số 0).
- Ô giá trị 0 luôn in ra để trống (`DisplayZeros = False` trên mọi sheet) ở cả 2 hàm xuất chung →
  áp cho in theo file tổng lẫn form. `API_SetTongAnSo0 False` để giữ số 0 (không có ô trên giao diện).
- `MsgBox` của VBA chỉ hiện ANSI → tiếng Việt thành "?". Thay bằng `MsgBoxU` (MessageBoxW, Unicode)
  ở mọi hộp thoại của `modInAnMain`.
- Đẩy nhánh lên fork `QuangHuyQS/In-an-du-toan-G8` (tài khoản QuangHuyQS không có quyền ghi repo gốc).

## Lần 9 – Bằng chữ, áp các sửa cho FORM in ấn (04/11/2026)
- Ẩn dòng "Bằng chữ: …" trên mọi sheet ngay trước khi xuất (gõ tay, dễ sai khi số liệu đổi).
- Người dùng: in theo file tổng khó dùng vì file tổng mỗi người tổ chức khác nhau → ưu tiên FORM.
- Form (Căn chỉnh / Xem trước / In / PDF / cả thư mục) nay có đủ: THKP viết HOA + đổi tên dự án/công trình,
  cột I, dòng tiêu đề theo STT; bỏ TH chi phí TB khi = 0; ẩn số 0; ẩn Bằng chữ; Giá tháng ẩn Mã số.
- Tuỳ chọn cho form đọc từ sheet `IN THEO FILE TONG` (C8, C9, C12, C13) — form không dựng lại được để thêm ô.
- Thử qua `API_Run1HM` (đúng hàm nút XUẤT PDF) trên file 9.1: OK.

## Lần 10 – Sắp thứ tự hạng mục trên form, file .bat mở form (04/11/2026)
- Mỗi dòng hạng mục trên form có ô số thứ tự (mặc định 1, 2, 3…). Gõ lại số để quyết định HM nào in là HM1, HM2…:
  tên PDF tách (`_HM1 Xây dựng`, `_HM2 Lắp đặt`) và thứ tự trang khi ghép theo số này.
- Ô số tạo lúc chạy (`MakeOrderBoxes`), không cần dựng lại giao diện. BUILD_V6.ps1 `-KeepForm` nay thay code form từ `gen\`.
- Engine: `gHmOrd`, `HmSequence`, `API_SetHmOrder("2,1")`; in theo file tổng luôn theo thứ tự trong file.
- `MO_FORM_IN_AN.bat` (+ `.vbs`): mở thẳng form, không phải mở .xlsm. Dùng Excel đang mở nếu có;
  file tool mở chỉ đọc, ẩn cửa sổ; đóng form thì tự đóng file tool.
- Thử: file 1d (2 HM) thứ tự "3,1" → `_HM1 Xây dựng.pdf`, `_HM2 Lắp đặt.pdf`. File .bat mở được form.

## Lần 11 – Đóng gói gửi người khác, ẩn "Thời điểm lập" (04/11/2026)
- Rà mã nguồn + file xlsm: không có đường dẫn cứng, không link ngoài, chỉ tham chiếu chuẩn của Office,
  nút gọi macro theo tên; cài đặt form lưu registry từng máy. Declare có `#If VBA7` (Excel 32/64 bit).
- `MO_FORM_IN_AN.bat` gỡ dấu "tải từ Internet" (Unblock-File) trong thư mục của nó để Excel không chặn macro;
  `.vbs` dùng file tool trùng tên đang mở (Excel không mở được 2 file trùng tên).
- Gói `dist\In_an_du_toan_G8_V6.5.zip` = xlsm + .bat + .vbs + `DOC_TRUOC.txt`. Thử giải nén vào thư mục
  có dấu/cách, gắn dấu Internet: .bat gỡ dấu, mở form, đóng form → Excel tự thoát.
- "Thời điểm lập…" luôn ẩn khi in (mọi sheet, kể cả file nhiều HM — trước đây bị hiện lại khi tách khối).

## Chưa làm / hạn chế
- Máy tắt VBScript (Windows sau này / chính sách công ty): .bat không chạy → mở thẳng file .xlsm.
  Chức năng in theo file tổng cũng dùng VBScript.RegExp.
- Đổi số thứ tự HM không tự dời tick Giá tháng (mặc định HM cuối) / TH chi phí TB (HM đầu) – tick lại tay nếu cần.
- Chưa mở PDF để kiểm tra từng trang (môi trường không có trình xem PDF) – cần người dùng duyệt mẫu.
- Giá tháng vẫn ẩn E, G:J theo chữ cái cột (cùng layout ở mọi file đã gặp); cột Mã số đã tìm theo tiêu đề.
- Tuỳ chọn trên form `frmInDuToan` (in từng file) chưa có ô Tên dự án / viết hoa – mới có ở in theo file tổng.
- Tuỳ chọn bỏ Đơn giá chi tiết và lọc theo thư mục chưa có nút trên giao diện.
- Form `frmInDuToan` giữ nguyên (VBE Designer báo lỗi 75 khi dựng lại bằng code).
