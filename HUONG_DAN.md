# Hướng dẫn sử dụng – Xuất PDF theo FILE TỔNG HỢP (V6.5)

> **Mở nhanh form in:** bấm đúp `MO_FORM_IN_AN.bat` (để cùng thư mục với `FILE_TONG_IN_G8_V*.xlsm`).
> **Thứ tự hạng mục:** file nhiều HM, gõ số ở ô đầu mỗi dòng hạng mục (1, 2, 3…) để chọn HM nào in là HM1, HM2.

Tool đọc file tổng hợp (ví dụ `0. TH LUONG TAI.xlsx`), đi theo các công thức link
tới từng file dự toán chi tiết, rồi xuất mỗi hạng mục thành 1 file PDF, đúng thứ tự trong file tổng.

### Link nhiều cấp (từ V6.5)
Tool đi theo link qua bao nhiêu file trung gian cũng được (tối đa 6 cấp):
- Ô trong file trung gian là `SUBTOTAL(...)` / `SUM(...)` / `=E11+E12` → tool đi xuống từng ô con
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

## 1. Chuẩn bị
- Máy có Excel. Mở file `FILE_TONG_IN_G8_V6.5.xlsm`, bấm **Enable Content** (cho phép macro).
- Các file dự toán chi tiết phải còn ở đúng đường dẫn mà file tổng đang link (ổ M:, W:…).
- Đóng các file Excel khác đang mở cùng tên, để tool mở được file ở chế độ chỉ đọc.

## 2. Cách chạy bằng nút trên sheet
1. Trên sheet đầu của file `.xlsm`, bấm nút xanh **▶ XUẤT PDF THEO FILE TỔNG HỢP** → chọn file tổng hợp.
2. Tool chuyển sang sheet **IN THEO FILE TONG** (cấu hình, được nhớ cho lần sau):
   | Ô | Ý nghĩa |
   |---|---|
   | File tổng hợp | đổi bằng nút *Chọn file tổng hợp…* |
   | Thư mục lưu PDF | trống = thư mục `PDF` cạnh file tổng |
   | Chạy từ / đến mục số | chạy tiếp khi bị ngắt, hoặc chạy một đoạn |
   | Tên dự án in ra | thay dòng `Tên dự án:` ở THKP (mọi khối hạng mục). Trống = giữ như file nguồn |
   | Tên công trình in ra | thay `Tên công trình:` / `CÔNG TRÌNH :` ở mọi sheet. Trống = giữ |
   | In cột Tổng cộng | cột Tổng cộng ở Giá tổng hợp |
   | Chỉ dò link | không xuất PDF, chỉ ghi `_THU_TU_XUAT.txt` để xem trước |
   | Viết HOA dòng tên ở THKP | `TÊN DỰ ÁN: … / TÊN CÔNG TRÌNH: … / TÊN HẠNG MỤC: …` (mặc định Có) |
   | Giá tháng: ẩn cột Mã số | mặc định Có (cột tìm theo tiêu đề "Mã số") |
3. Bảng chọn sheet: mỗi sheet có 2 ô **thứ tự in** — cột *File 1 hạng mục* và cột *File 2, 3… hạng mục*.
   Để trống = không in; số nhỏ in trước. Mặc định: TH chi phí TB (1), THKP (2), Giá tổng hợp (3),
   Đơn giá chi tiết (4), Giá tháng (5).
   - File 1 hạng mục: cả file 1 PDF, chọn được thêm Công trình, Tổng hợp VT, TH chi phí XD.
   - File nhiều hạng mục: mỗi hạng mục 1 PDF; THKP / Giá tổng hợp / Đơn giá CT / Công trình được tách
     đúng phần của hạng mục đó; Giá tháng in kèm mọi hạng mục; TH chi phí TB đi kèm hạng mục đầu.
   - TH chi phí TB chỉ in khi tổng thiết bị khác 0.
4. Bấm **▶ BẮT ĐẦU XUẤT**. Xong, tool báo số PDF và mở thư mục kết quả.

Ô có giá trị 0 **luôn in ra để trống** (tắt "Show a zero in cells that have zero value" trên mọi sheet trước
khi xuất) — áp dụng cho cả in theo file tổng lẫn in từng file trên form.

Dòng **"Bằng chữ: …"** (số tiền đọc thành chữ, gõ tay nên dễ sai khi số liệu đổi) được **ẩn trên mọi sheet**
trước khi xuất. `API_SetGiuBangChu True` để giữ.

Tên dự án / công trình chỉ thay trên bản in — file nguồn mở chỉ đọc, **không bị lưu đè**.

## 3. Mỗi PDF gồm những gì (theo thứ tự trang)
TH chi phí TB (chỉ ở hạng mục đầu của file) → THKP hạng mục → Giá tổng hợp → Đơn giá chi tiết → Giá tháng.
- File có nhiều hạng mục: Giá tổng hợp và Đơn giá chi tiết được **tách theo từng hạng mục** đã tách ở THKP; Giá tháng đi kèm mỗi hạng mục.
- Sheet Giá tháng bị ẩn trong file nguồn sẽ được tự hiện lại để in.
- Tên file: `NN_<STT> <tên hạng mục>.pdf` hoặc `NN.kk_<STT> <tên mục con>.pdf` (NN = số thứ tự trong file tổng).

## 4. Chạy không có Đơn giá chi tiết, hoặc chỉ một thư mục (dùng lệnh)
Các tuỳ chọn này hiện chưa có nút, chạy qua COM/PowerShell hoặc trong VBE (Immediate window):

```vb
API_SetTongFilter "\1. DHKK\"     ' chỉ xuất mục có file nguồn nằm trong thư mục này ("" = tất cả)
API_SetTongNoDGCT True            ' True = KHÔNG xuất Đơn giá chi tiết
API_SetTongDryRun True            ' True = chỉ dò link + ghi _THU_TU_XUAT.txt, KHÔNG xuất PDF (xem trước, ~30 giây)
API_SetTongTen "TÊN DỰ ÁN MỚI", ""  ' tên dự án / tên công trình in ra ("" = giữ nguyên)
API_SetTongSheets "8,1,2,3,5", "6,1,2,3,5"  ' sheet + thứ tự: file 1 HM (1 THKP,2 GTH,3 DGCT,4 CT,5 GT,6 THVT,7 THXD,8 TB) / nhiều HM (1..5 như trên, 6 TB)
Dim rep As String
n = API_RunTong("M:\...\0.KSDT DD D.TOAN T.HOP 26.01.2022.xlsx", "M:\...\PDF\KHONG_DGCT", False, rep)
```
Tham số thứ 3 của `API_RunTong`: `True` = in cột Tổng cộng, `False` = không in.

## 5. Đọc kết quả
- Mỗi lần chạy ghi `_THU_TU_XUAT.txt` trong thư mục đích, dạng cây: mục file tổng → các mục con →
  đường dẫn PDF (tương đối) → file nguồn (dòng THKP / sheet TB / in nguyên sheet / cả file).
- `KHONG XUAT`: mục ở file tổng là link nội bộ / số nhập tay, không có file dự toán chi tiết (ví dụ: ngăn cháy lan, xử lý nước thải, Onsen, ống lồng xuyên sàn).
- `(chung PDF voi muc NN)`: nhiều mục cùng trỏ vào một hạng mục nên dùng chung PDF (B.2 dùng chung với B.1).
- `Canh bao`: PDF vẫn xuất nhưng thiếu một phần (ví dụ file nguồn không có khối Đơn giá chi tiết cho hạng mục đó).

## 6. Xử lý sự cố
- **Không mở được file:** kiểm tra ổ mạng M:/W: còn kết nối, file nguồn không bị khoá.
- **Dừng giữa chừng:** chạy lại và nhập số mục bắt đầu ở bước 4.
- **Còn tiến trình Excel nền:** đóng bằng Task Manager nếu file báo "đang được dùng".
- Lần chạy đầy đủ có Đơn giá chi tiết mất khoảng 45 phút; bỏ Đơn giá chi tiết còn khoảng 25 phút.

## 7. Build lại tool từ mã nguồn
```powershell
.\BUILD_V6.ps1 -KeepForm -Template .\FILE_TONG_IN_G8_V6.4.xlsm -Out "$PWD\FILE_TONG_IN_G8_V6.5.xlsm"
```
Cần Excel bật "Trust access to the VBA project object model". Mã nguồn ở `src\`, thư mục `gen\` tự sinh.
