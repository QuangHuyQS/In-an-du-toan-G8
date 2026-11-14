# Hướng dẫn sử dụng – Phần mềm In ấn dự toán G8 (bản 7.1)

> Mở bằng **`InAnG8.exe`** (để chung thư mục với `FILE_TONG_IN_G8_V7.1.xlsm`). Không cần mở file .xlsm.
> Form cũ (nền xanh, nút cam, `MO_FORM_IN_AN.bat`) đã ngừng dùng từ bản 7.1.
> Thao tác từng tab, bố cục, phím tắt: xem [DOC_TRUOC.txt](DOC_TRUOC.txt). File này nói **cách phần mềm in**.

## 1. Hai cách in
- **Tab "Danh sách file dự toán"**: kéo thả file / thư mục dự toán G8. Phần mềm tự nhận file 1 hạng mục
  hay nhiều hạng mục, xem trước đúng như sẽ in, chọn sheet + thứ tự cho từng file / từng hạng mục.
- **Tab "In theo file tổng hợp"**: chọn file tổng, quét link, mỗi mục chọn sheet, thứ tự sheet, hướng
  dọc / ngang; xuất mỗi mục 1 PDF đúng thứ tự file tổng.

Chuẩn bị: các file dự toán chi tiết phải còn đúng đường dẫn mà file tổng đang link (ổ M:, W:…).
Phần mềm mở mọi file **chỉ đọc**, dùng Excel chạy ngầm riêng, không đụng Excel bạn đang làm.

## 2. In theo file tổng hợp: phần mềm đi theo link thế nào
Ở tab **In theo file tổng hợp**, phần mềm đọc file tổng hợp (ví dụ `0. TH LUONG TAI.xlsx`), đi theo các công thức link
tới từng file dự toán chi tiết, rồi xuất mỗi hạng mục thành 1 file PDF, đúng thứ tự trong file tổng.

### Link nhiều cấp
Phần mềm đi theo link qua bao nhiêu file trung gian cũng được (tối đa 6 cấp):
- Ô trong file trung gian là `SUBTOTAL(...)` / `SUM(...)` / `=E11+E12` → phần mềm đi xuống từng ô con
  (với `SUBTOTAL` bỏ qua ô con cũng là SUBTOTAL, đúng như Excel tính) và theo từng link ra file khác.
- Link vào `THKP hạng mục` → in dự toán của đúng hạng mục đó.
- Link vào `TH chi phí TB` của file dự toán → chỉ in sheet TH chi phí TB.
- Link vào `TH chi phí XD` của file dự toán → in hạng mục ghi ở dòng đó; nếu là dòng tổng của file nhiều
  hạng mục thì xuất mỗi hạng mục 1 PDF `..._HM1.pdf`, `..._HM2.pdf`…
- Ô không còn link nào xuống dưới (bảng báo giá, `OUTDOOR FITNESS`…) → in nguyên sheet đó.

Một dòng ở file tổng có thể ra nhiều PDF: đánh số `NN.01_`, `NN.02_`…

### PDF nằm trong thư mục con giống thư mục nguồn
PDF được đặt vào thư mục con **giống vị trí file nguồn so với file tổng**, để người dùng tự ghép theo thư mục:
```
PDF\
  01_A-1.1 San nền.pdf                         (file nguồn ở cùng thư mục file tổng)
  6. TH Dien\12.01_A-a Đường dây trung thế ... .pdf
  Hoan tra tram bom\09.01_A-3.2 Cổng, hàng rào.pdf
  TXLNT LUONG TAI\Excel\20.01_A-8 San nền ... .pdf
```
File nguồn nằm ngoài thư mục file tổng → `_NGOAI THU MUC TONG\`.

## 3. Luật chung khi in
Ô có giá trị 0 **luôn in ra để trống** (tắt "Show a zero in cells that have zero value" trên mọi sheet trước
khi xuất) — áp dụng cho cả 2 tab.

Dòng **"Bằng chữ: …"** (số tiền đọc thành chữ, gõ tay nên dễ sai khi số liệu đổi) được **ẩn trên mọi sheet**
trước khi xuất.

Tên dự án / công trình chỉ thay trên bản in — file nguồn mở chỉ đọc, **không bị lưu đè**.

- **TH chi phí TB** chỉ mặc định in kèm hạng mục có tên chứa "Lắp đặt" hoặc "Thiết bị", và chỉ khi tổng
  thiết bị khác 0. Tick tay thêm được (có cảnh báo).
- Số nguyên không hiện dấu phẩy thừa. Ô số dài hơn cột (`###`): cột nới có giới hạn, số quá dài thì thu
  nhỏ chữ riêng ô đó; mỗi chỗ xử lý ghi vào Nhật ký và báo cáo xuất.

## 4. Mỗi PDF gồm những gì (thứ tự mặc định, đổi được trên phần mềm)
TH chi phí TB (theo luật ở mục 3) → THKP hạng mục → Giá tổng hợp → Đơn giá chi tiết → Giá tháng.
- File có nhiều hạng mục: Giá tổng hợp và Đơn giá chi tiết được **tách theo từng hạng mục** đã tách ở THKP; Giá tháng đi kèm mỗi hạng mục.
- Sheet Giá tháng bị ẩn trong file nguồn sẽ được tự hiện lại để in.
- Tên file: `NN_<STT> <tên hạng mục>.pdf` hoặc `NN.kk_<STT> <tên mục con>.pdf` (NN = số thứ tự trong file tổng).

## 5. Đọc kết quả
- Mỗi lần chạy ghi `_THU_TU_XUAT.txt` trong thư mục đích, dạng cây: mục file tổng → các mục con →
  đường dẫn PDF (tương đối) → file nguồn (dòng THKP / sheet TB / in nguyên sheet / cả file).
- `KHONG XUAT`: mục ở file tổng là link nội bộ / số nhập tay, không có file dự toán chi tiết (ví dụ: ngăn cháy lan, xử lý nước thải, Onsen, ống lồng xuyên sàn).
- `(chung PDF voi muc NN)`: nhiều mục cùng trỏ vào một hạng mục nên dùng chung PDF (B.2 dùng chung với B.1).
- `Canh bao`: PDF vẫn xuất nhưng thiếu một phần (ví dụ file nguồn không có khối Đơn giá chi tiết cho hạng mục đó).

## 6. Xử lý sự cố
- **Không mở được file:** kiểm tra ổ mạng M:/W: còn kết nối, file nguồn không bị khoá.
- **Dừng giữa chừng:** chạy lại, chỉ tick các mục chưa xuất.
- **Lỗi / cảnh báo:** xem ở Nhật ký (góc dưới phải, Ctrl+L) và file `BAO_CAO_XUAT_*.txt` trong thư mục PDF.
- File tổng nhiều mục chạy lâu (nhiều chục phút): bỏ tick Đơn giá chi tiết để nhanh hơn.
