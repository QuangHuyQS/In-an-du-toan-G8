# Phần mềm In ấn dự toán G8 – bản 7.1

Xuất PDF từ file Excel dự toán G8: in từng file / cả thư mục, hoặc **in theo file tổng hợp**
(đi theo link trong file tổng tới từng file dự toán chi tiết, tách theo hạng mục, đúng thứ tự file tổng).

## Tải và dùng

1. Vào mục **[Releases](https://github.com/QuangHuyQS/In-an-du-toan-G8/releases/latest)**, tải
   `In_an_du_toan_G8_V7.1_<ngày>.zip`.
2. Giải nén ra một thư mục (không chạy trong file zip). Giữ 2 file chung thư mục:
   - `InAnG8.exe`: **phần mềm, bấm đúp để mở**
   - `FILE_TONG_IN_G8_V7.1.xlsm`: bộ máy in, phần mềm tự dùng. **Không cần mở.**
3. Đọc `DOC_TRUOC.txt` trong gói, hoặc [HUONG_DAN.md](HUONG_DAN.md).

Yêu cầu: Windows 10/11 có Microsoft Excel 2010 trở lên. Không cần cài thêm gì, không cần mạng.

> Form cũ (nền xanh, nút cam, mở bằng `MO_FORM_IN_AN.bat`, bản V6.x) **đã ngừng dùng**.
> File `FILE_TONG_IN_G8_V7.1.xlsm` không còn form đó. Mọi thao tác làm trên `InAnG8.exe`.

## Thư mục mã nguồn

| Đường dẫn | Nội dung |
|---|---|
| `InAnG8.exe`, `FILE_TONG_IN_G8_V7.1.xlsm` | bản build mới nhất (giống gói trong Releases) |
| `app\` | mã nguồn phần mềm (C# WinForms, .NET Framework 4). Build: `app\build_app.ps1` |
| `src\` | mã nguồn bộ máy VBA (`inanDTG8_V6.bas`). Build: `BUILD_V6.ps1` |
| `DOC_TRUOC.txt` | hướng dẫn nhanh, chép vào gói phát hành |
| `HUONG_DAN.md` | hướng dẫn đầy đủ (cách tool đi theo link, tên PDF, xử lý sự cố) |
| `NHAT_KY.md`, `BAN_GIAO_IN_AN.md` | nhật ký thay đổi, ghi chú bàn giao cho người phát triển |

Build lại (cần Excel bật *Trust access to the VBA project object model*):

```powershell
powershell -ExecutionPolicy Bypass -File BUILD_V6.ps1 -Out "$env:TEMP\FILE_TONG_IN_G8_V7.1.xlsm"
Copy-Item "$env:TEMP\FILE_TONG_IN_G8_V7.1.xlsm" . -Force
powershell -ExecutionPolicy Bypass -File app\build_app.ps1
```
