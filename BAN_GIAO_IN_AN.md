# BÀN GIAO – TOOL IN ẤN DỰ TOÁN G8 (đọc file này đầu tiên)

> Cập nhật: **15/11/2026** (phiên 12–15/11 – tên trên bản in, thứ tự + hướng từng mục file tổng, ô ###; **bỏ hẳn form VBA cũ**, file .xlsm chỉ còn là bộ máy; tick cha–con cây file tab 1; chi tiết NHAT_KY Lần 20–28). Mỗi lần bàn giao: ghi đè file này, giữ đúng tên.
>
> **Mở phiên mới:** đặt thư mục làm việc `C:\Users\DT-06\Documents\TOOL - RND7\In an du toan\In-an-G8`, gõ:
> *"Đọc BAN_GIAO_IN_AN.md rồi làm tiếp"*.

Gốc dự án (gọi tắt **G8**): `C:\Users\DT-06\Documents\TOOL - RND7\In an du toan\In-an-G8\`.

| Cần gì | Ở đâu |
|---|---|
| File bàn giao này | `G8\BAN_GIAO_IN_AN.md` |
| Nhật ký thay đổi (mới nhất ở cuối) | `G8\NHAT_KY.md` |
| Mã nguồn VBA (sửa ở đây, KHÔNG sửa `gen\`) | `G8\src\` |
| Phần mềm desktop V7.1 (C#) | `G8\app\InAnG8.cs` + `*Ui.cs`, build `G8\app\build_app.ps1` (mặc định ra gốc `G8\`, cạnh `FILE_TONG_IN_G8_V7.1.xlsm`) |
| Bản build mới nhất | `G8\InAnG8.exe` + `G8\FILE_TONG_IN_G8_V7.1.xlsm` (có trong git; xlsm cũng là file khung khi build) |
| Gói gửi người dùng | `G8\dist\In_an_du_toan_G8_V7.1_261115(.zip)`, cũng ở GitHub Releases `QuangHuyQS/In-an-du-toan-G8` (tag `v7.1-261115`; bản trước `v7.1-261112` giữ lại). Chỉ phát hành gói phần mềm, **không còn gói form .bat V6.x** |
| Form cũ + gói cũ (lưu trữ, không dùng) | `G8\_LUU_TRU_FORM_CU\` (gitignore): MO_FORM_IN_AN.bat/.vbs, xlsm V6.4–6.6, video HD cũ, dist V6.5–V7.1 cũ |
| Thiết kế UI/UX (Antigravity – **bắt buộc**, Claude chỉ dựng theo) | `In an du toan\UI_PHAN_MEM_IN_G8\` – `BRIEF*.md` → `THIET_KE*.md`, `mockup*.html`, `XONG*.txt` (bản 1: tổng thể; 2: tab file tổng + luật TB; 3: tách vùng / co giãn) |
| File thử | `In an du toan\Test\` (5 file: TXL 4 HM, 2. ĐM38 3 HM, 9. ĐM36 3 HM, 10. ĐM38 1 HM, 9.2b 1 HM) |
| File tổng hợp mẫu | `M:\DUTOAN\2026\01. BT2.MAG\02. OUTPUT\04. DU LIEU\HUY\261103\260724. THDT LT\6. TH Dien\260709.TH DZ TBA - LT.xlsx` (sheet `THC`, 34 link) |

---

## 1. Kiến trúc

```
InAnG8.exe (WinForms, .NET Framework 4.x có sẵn trong Windows)
   │  COM (Excel.Application ẩn, riêng của phần mềm, 1 luồng STA + hàng đợi việc)
   ▼
FILE_TONG_IN_G8_V7.1.xlsm  ── lõi VBA inanDTG8_V6.bas (+ modInAnMain); KHÔNG còn form VBA / nút trên sheet
   │  mở file dự toán chỉ đọc → dàn trang → ExportAsFixedFormat PDF
   ▼
PDF (xem trước: %TEMP%\InAnG8_Preview\..., xóa khi đóng phần mềm ; xuất thật: thư mục người dùng chọn)
   ▼
Windows.Data.Pdf (có sẵn Win 10/11) vẽ trang PDF thành ảnh trong khung xem trước
```

- Không cần cài gì thêm. **Không dùng WebView2 / NuGet** (máy không có .NET SDK; tải NuGet cần xin phép). PDF hiển thị bằng `Windows.Data.Pdf`, gọi WinRT thủ công (không có `Windows.winmd` gộp → không `await`, chờ `IAsyncOperation.Status` rồi `GetResults()`; đọc ảnh bằng `DataReader`). Không dùng `Windows.UI.Color` / `Windows.Foundation.Size` (cần System.Runtime.WindowsRuntime → lỗi biên dịch).
- Form VBA cũ (`frmInDuToan`, `clsInAnBtn`, `MO_FORM_IN_AN.bat`) **đã bỏ từ 12/11/2026**. Sheet 1 của xlsm chỉ ghi "mở bằng InAnG8.exe"; sheet `IN THEO FILE TONG` ẩn hẳn (VeryHidden) nhưng giữ giá trị vì lõi còn đọc C12/C13 (`TongLoadUiOptions`). Macro `MoBangDieuKhien` còn nhưng chỉ báo "form cũ đã ngừng dùng".

### 1.1 Hàm lõi cho phần mềm ngoài (trong `src\inanDTG8_V6.bas`, chỉ trả String, không ByRef)
| Hàm | Việc |
|---|---|
| `APP_Detect(path)` | `kind⇥nHM⇥tên HM (|)⇥ex1 (8 ký tự 0/1)⇥ex2 (6)⇥lỗi`; kind 0 không phải dự toán / 1 = 1 HM / 2 = nhiều HM |
| `APP_Run(path, outFolder, isPdf, sel1, selN, custN, hmOrd, ghep, tenHM, tongCot)` | xuất 1 file → `ok⇥kind⇥thông báo`. `custN` = sheet riêng từng HM `"1,2,3;0;1,4"` (0 = bỏ HM) → `gCustomN` → `ApplyCustomN` (gọi cuối `ApplyTemplateSel`) |
| `APP_TongSheets(sumPath)` | sheet hiện của file tổng + số link, sắp giảm dần (`tên⇥số\n`) |
| `APP_Tong(sumPath, out, sheet, dry, sel1, selN, rowSel, inclTong)` | dò link (dry) / xuất theo file tổng → `số PDF⇥báo cáo`; danh sách mục: `API_TongItems` (mảng n×13: TT, nhãn, tên, giá trị, file, vị trí, chọn-sheet-được, đường dẫn, dòng THKP, loại 0/1/2, tên HM trung gian, dòng bảng tổng, chuỗi dòng cha `dòng␟STT␟nội dung␟giá trị` nối `␞`). `rowSel` "7=1,2;8=0" – `0` = không xuất mục đó |
| `APP_TongOne(sumPath, out, sheet, seq, sel1, selN, rowSel, inclTong)` | xem trước 1 mục = xuất thật riêng mục `seq` (≈10–20 s) |
| `APP_HmMap(path, "dòngTHKP,..", "tênHM|..")` | `nHM⇥h1,h2..` mục nào thuộc HM thứ mấy (0 = cả file / file 1 HM) |
| `HmIsTB(tên)` | luật TB: tên HM có "LẮP ĐẶT" / "THIẾT BỊ" (dùng ở `ApplyTemplateSel`, `TongDoFile`) |
| `APP_DbgBlocks(wbName, idx)` | chẩn đoán ranh giới khối HM của 1 loại sheet (giữ lại để soát lỗi) |

Mã sheet: hệ 1 HM `1 THKP, 2 GTH, 3 ĐGCT, 4 CT, 5 GT, 6 THVT, 7 THXD, 8 TB`; hệ nhiều HM `1 THKP, 2 GTH, 3 ĐGCT, 4 CT, 5 GT, 6 TB`.

### 1.2 Build
```bash
powershell -ExecutionPolicy Bypass -File BUILD_V6.ps1 -Out "$TEMP\FILE_TONG_IN_G8_V7.1.xlsm"   # khung mặc định = G8\FILE_TONG_IN_G8_V7.1.xlsm; -Out phải khác khung
cp "$TEMP/FILE_TONG_IN_G8_V7.1.xlsm" .                                                    # chép đè sau khi kiểm
powershell -ExecutionPolicy Bypass -File app/build_app.ps1                                 # ra G8\InAnG8.exe
```
- `BUILD_V6.ps1` chỉ còn chế độ bộ máy: bỏ form cũ nếu khung còn, gọi `BuildInAnForm_Sheet` rồi `BuildInAnForm_Engine`. Nó sinh `gen\` (ASCII, chuỗi tiếng Việt → `U("\uXXXX")`) rồi nạp vào xlsm qua COM (cần "Trust access to the VBA project"). **Chú thích (`'`) trong `src\*.bas` phải ASCII thuần.**
- `build_app.ps1` dùng `csc.exe` của .NET Framework 4 (C# 5: không `$""`, không `?.`, không `nameof`).
- Gói 7.0 = `InAnG8.exe` + `FILE_TONG_IN_G8_V7.0.xlsm` + `DOC_TRUOC.txt` cùng thư mục (phần mềm tìm `FILE_TONG_IN_G8_V*.xlsm` cạnh exe, bản số cao nhất, tự gỡ Zone.Identifier).

### 1.3 Chạy thử không giao diện (mẫu)
```powershell
$xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false; $xl.AutomationSecurity=1
$tool=$xl.Workbooks.Open("<...>\FILE_TONG_IN_G8_V7.0.xlsm",0,$true)
$xl.Run("'FILE_TONG_IN_G8_V7.0.xlsm'!APP_Run", $file, $out, $true, "8,1,2,3,4,5", "6,1,2,3,4,5", "", "", $false, $false, $false)
$tool.Close($false); $xl.Quit()
```
- **Luôn gọi macro kèm tên file lõi** (`'tên.xlsm'!APP_Run`) – nếu không, khi đang mở file khác Excel báo "Cannot run the macro".
- Kiểm tra PDF: Python `C:\Users\DT-06\AppData\Local\Programs\Python\Python313\python.exe -I` có `fitz` (PyMuPDF) + PIL. Chụp cửa sổ phần mềm mà không giật màn hình người dùng: `PrintWindow` (script mẫu `%TEMP%\x\cap.ps1`).
- Hộp thoại lỗi biên dịch VBA (modal) làm treo tự động hóa → còn sót `EXCEL.EXE /automation` chạy nền; tắt bằng Task Manager.

---

## 2. Phiên 05/11/2026 – bản 7.1 (chi tiết: `NHAT_KY.md` lần 19)
- **Luật TH chi phí TB**: chỉ mặc định in kèm HM có tên "Lắp đặt"/"Thiết bị" – ở lõi (form cũ in thư mục, file tổng) và phần mềm (2 tab, có hiển thị lý do + cảnh báo khi tick tay).
- **Tab In theo file tổng** làm lại theo `THIET_KE_2.md`: cây cha–con có cột (DataGridView vẽ tay), xem trước từng mục (`APP_TongOne`), chọn sheet từng mục, bỏ tick = không xuất, hộp báo cáo.
- **Bố cục** theo `THIET_KE_3.md`: tách vùng ⧉ ra cửa sổ riêng (xem trước nổi dùng chung 2 tab, F9), thu gọn « » (F6/F7/F8), co giãn < 1000 px, lưu bố cục / độ rộng cột (`%APPDATA%\InAnG8\caidat.txt`, khóa `lay_*`).
- **Chống Excel chạy ngầm**: Job Object KILL_ON_JOB_CLOSE + Kill khi đóng + file `excel_pid.txt` dự phòng. Đã thử tắt cưỡng bức phần mềm → Excel nền tắt theo.
- Đã kiểm bằng file tổng THC (34 mục): mục 26 HM Lắp đặt có 15 trang TB, mục 4 HM Xây dựng không TB, xuất thật chỉ mục tick.
- Chạy thử không cần bấm: `InAnG8.exe --tong "<file tổng>" THC 26` (quét + chọn mục 26); chụp cửa sổ bằng PrintWindow (`%TEMP%\x\cap.ps1`) – đỡ tốn token hơn chụp màn hình.

## 3. Còn mở / việc tiếp theo
1. **Ý kiến người dùng ở tab /btw (chuyển sang 05/11), CHỜ người dùng chọn:**
   a) Khoán thêm cho Antigravity viết luôn code UI: tách UI ra partial class `InAnG8.<Tab>Ui.cs` theo "hợp đồng" hàm/dữ liệu do Claude định; Antigravity làm UI tab, hộp báo cáo dạng bảng, làm đẹp tab 1, icon / hướng dẫn / DOC_TRUOC, soát lệch code UI so với THIET_KE_2/3 + mockup. Claude giữ lõi VBA/COM + chạy thử thật.
   b) Bỏ phụ thuộc .xlsm/macro: **Mức 1 (khuyên)** chuyển lõi VBA (~5.000 dòng) sang C# trong InAnG8.exe, vẫn dùng Excel ẩn qua COM để mở .xls + xuất PDF (bản in y hệt), kiểm bằng so PDF cũ/mới trên `Test\`, làm dần (1 file → nhiều HM → file tổng). **Mức 2** không cần Excel (LibreOffice headless lệch bản in; Aspose/Spire tốn phí) – chưa khuyên.
2. (Đã đóng) Form VBA cũ bỏ 12/11/2026, không phát hành gói .bat nữa. `src\BuildInAnForm.bas` còn code dựng form cũ (`BuildInAnForm_Run`, `BuildControls`) không ai gọi, xóa được khi rảnh.
3. Câu hỏi mở của Antigravity (`XONG_2.txt`, `XONG_3.txt`): cache quét file tổng (.g8cache), tùy chọn ghép N HM cho mục "Cả file", tên PDF dài/MAX_PATH. Chưa làm.
4. Chưa có: hủy quét giữa chừng, xem trước tab 2 có cache theo cấu hình (hiện chỉ giữ PDF của lần xem gần nhất mỗi mục).
5. Lỗi 1004 `Application.Calculation` khi file ở Protected View: chưa bọc `On Error`.
6. Git: nhánh `main` của fork `QuangHuyQS/In-an-du-toan-G8` (remote `fork`) là bản chính; `origin` (tonghang0111-rgb/In-an-G8) không đẩy. `InAnG8.exe` + `FILE_TONG_IN_G8_V7.1.xlsm` ở gốc có trong git; `dist\`, `_LUU_TRU_FORM_CU\` thì không.
7. Còn 1 `EXCEL.EXE /automation` sót từ phiên 04/11 (PID 36252) – người dùng tự tắt.

## 4. Lưu ý làm việc với người dùng
- Người dùng là kỹ sư dự toán, nói tiếng Việt; muốn kết quả kiểm chứng bằng file thật (thư mục `Test`, file tổng trên ổ M:).
- Khi sửa lỗi in: tái hiện bằng script trước, xem ảnh trang PDF, rồi mới đóng gói lại gói phần mềm V7.1 (không còn gói V6.x).
- Không sửa thư mục `THAM_KHAO\` của brief Antigravity; giao việc cho Antigravity bằng `BRIEF*.md` + chờ `XONG*.txt`.
- Không dùng sed để sửa script có đường dẫn Windows (mất dấu `\`) và không dùng perl với chuỗi có `$(` (bị nội suy) – dùng Edit / heredoc `'EOF'`.
- UI/UX bắt buộc qua Antigravity: viết `BRIEF_n.md` (kèm dữ liệu thật, ràng buộc WinForms .NET 4), gửi bằng computer-use vào ô chat Antigravity (project UI_PHAN_MEM_IN_G8), rồi chờ `XONG_n.txt` bằng lệnh nền (không chụp màn hình theo dõi).
- Người dùng rất ngại Excel chạy ngầm – mọi thay đổi chạm vào Engine phải thử lại cả đóng thường lẫn tắt cưỡng bức.
