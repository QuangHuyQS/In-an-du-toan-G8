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

## Lần 12 – Video hướng dẫn, chuyển thư mục dự án (04/11/2026)
- Dự án chuyển về `Documents\TOOL - RND7\In an du toan\In-an-G8` (giữ nguyên git, remote `fork`).
- `MO_FORM_IN_AN.vbs` chọn file tool theo SỐ PHIÊN BẢN cao nhất (trước theo ngày sửa: chép thư mục làm
  V6.4 "mới hơn" → mở nhầm bản không có macro mới).
- 2 video (1920x1080, chữ chú thích tiếng Việt, khung cam + con trỏ chỉ chỗ bấm), dựng từ ảnh chụp form thật
  chạy trên dữ liệu mẫu: `video\HD1_In_theo_file.mp4` (3 phút 13) – 1 HM, nhiều HM, đổi thứ tự HM, Tách/Ghép;
  `video\HD2_In_ca_thu_muc.mp4` (1 phút 20). Có trang PDF thật ở cuối mỗi phần. Đã đưa vào gói zip.
- `video/` không đưa lên git (file lớn).

## Chưa làm / hạn chế
- Máy tắt VBScript (Windows sau này / chính sách công ty): .bat không chạy → mở thẳng file .xlsm.
  Chức năng in theo file tổng cũng dùng VBScript.RegExp.
- Đổi số thứ tự HM không tự dời tick Giá tháng (mặc định HM cuối) / TH chi phí TB (HM đầu) – tick lại tay nếu cần.
- Chưa mở PDF để kiểm tra từng trang (môi trường không có trình xem PDF) – cần người dùng duyệt mẫu.
- Giá tháng vẫn ẩn E, G:J theo chữ cái cột (cùng layout ở mọi file đã gặp); cột Mã số đã tìm theo tiêu đề.
- Tuỳ chọn trên form `frmInDuToan` (in từng file) chưa có ô Tên dự án / viết hoa – mới có ở in theo file tổng.
- Tuỳ chọn bỏ Đơn giá chi tiết và lọc theo thư mục chưa có nút trên giao diện.
- Form `frmInDuToan` giữ nguyên (VBE Designer báo lỗi 75 khi dựng lại bằng code).

## Lần 13 – In theo bảng tổng hợp: chọn sheet, dò ngược link về đúng hạng mục, nhãn cha–con (04/11/2026)
- Lỗi: file `260709.TH DZ TBA - LT.xlsx` có 4 sheet (THC, TK, TMDT HT, TKCS LT); tool tự chọn sheet "nhiều link nhất" = TKCS LT
  → dò sai bảng. Thêm ô **C14 "Sheet của file tổng cần dò link"** (trống = tự chọn như cũ); báo cáo liệt kê sheet + số link nếu gõ sai tên.
- Link vào sheet trung gian của file dự toán (vd `'TH chi phí XD'!C5`) trước đây chỉ in "cả file"/đoán theo tên hạng mục.
  Nay `TongTraceThkpRow` dò công thức nội bộ của ô đó (qua mọi sheet trung gian) đến dòng THKP mà ô lấy giá trị
  → xác định CHÍNH XÁC hạng mục (file nhiều HM chỉ in khối của HM đó, không phụ thuộc tên). Trên THC: 30/30 mục đầu đều ra "THKP dòng N".
- Nhãn phân cấp theo bảng tổng cha–con (`TongRowLabel`, theo SUBTOTAL/SUM bao quanh): `I-2.1-a-1`, `I-2.1-b-4`… (trước là `I-` trùng nhau).
  Thứ tự PDF vẫn theo thứ tự dòng bảng tổng.
- Báo cáo `_THU_TU_XUAT.txt` thêm "giá trị" của ô link để đối chiếu với số trên bảng tổng.
- Thử: chạy khô THC → đủ mục, nhãn đúng; xuất thật mục 7–8 (file 2c nhiều HM) → 2 PDF riêng (16 và 14 trang), mỗi PDF 1 hạng mục.
- Chưa làm: duyệt nội dung PDF bằng mắt; chọn sheet in riêng cho từng dòng bảng tổng (hiện chọn chung theo bảng thứ tự sheet).
- (Lần 13b) Chọn sheet in riêng theo từng dòng bảng tổng: sau CHẠY KHÔ tool dựng sheet `CHON SHEET THEO MUC`
  (mỗi dòng: nhãn, hạng mục, giá trị, file nguồn, vị trí + 6 cột sheet THKP / Giá TH / ĐGCT / Công trình / Giá tháng / TH chi phí TB; gõ `x` để in).
  Ô C2 = Có thì lần xuất thật dùng bảng này; Không thì dùng cấu hình chung. Engine: `API_SetTongRowSel "7=1;8=1,2"`, `API_TongItems`.
  Thử: mục 7 chỉ THKP → 1 trang; mục 8 THKP+Giá TH → 5 trang (trước 16 và 14 trang khi in đủ).
- (Lần 13c) Form `frmInDuToan` có nút "In theo file tổng" (tạo lúc chạy `MakeTongBtn`, cạnh nút Chọn thư mục): gọi `XuatTong_TuForm`
  → hiện cửa sổ file tool (đang ẩn khi mở bằng .bat), sang sheet IN THEO FILE TONG. Đóng form sau đó không tự đóng file tool.
  Đã đóng gói `dist\In_an_du_toan_G8_V6.6(.zip)`.


Lan 14: sua (1) PDF tach hang muc cua sheet Cong trinh tu HM2 tro di thieu tieu de + header cot; (2) TH chi phi TB bi cat sau dong STT dau (ra bang thiet bi trong).
Lan 15: FixHashCells - noi cot khi o so/ngay hien ### (ke ca chua hao 12% cho font khi in PDF).
Lan 16: Cong trinh khong AutoFit cot F nua (o khong xuong dong lam cot F gian 185 -> ca bang co 1 trang); F > 60 dua ve 40 + bat xuong dong o dai (CTCapDanhMuc, CTWrapDanhMuc).
Lan 17: Cong trinh - o Danh muc da xuong dong nhung dong co dinh chieu cao thap (chu bi cat/de) -> AutoFit dong + 10% chua hao khi in.

Lan 16: phan mem desktop V7.0 (app\InAnG8.cs, WinForms + Windows.Data.Pdf, loi VBA qua COM): APP_Detect / APP_Run (gCustomN chon sheet tung HM) / APP_TongSheets / APP_Tong.
Lan 17: Gia tong hop file nhieu HM mat dong TONG HANG MUC / TONG CONG TRINH / LAM TRON - SheetLastRow2 dung End(xlUp) bo qua dong an (khoi HM khac da an) -> doi sang LastRowAnyCol (Find xlFormulas) + xet cot C,I,K,L cho idx 2. APP_Run goi macro kem ten file loi.
Lan 18: Gia tong hop xuat TACH: moi HM dung o dong TONG HANG MUC (bo TONG CONG TRINH / LAM TRON o PDF HM cuoi); xuat GHEP van giu.

## Lần 19 – Bản 7.1: luật TH chi phí TB theo tên HM, tab file tổng có xem trước, tách vùng (05/11/2026)
- Luật TB: sheet TH chi phí TB chỉ mặc định in kèm hạng mục có tên chứa "LẮP ĐẶT" / "THIẾT BỊ" (bỏ dấu, không phân biệt hoa thường).
  Lõi: `HmIsTB` (dùng `KindFromName`), `HmName1` (tên HM file 1 HM ở THKP), `ApplyTemplateSel` cột 6, `TongDoFile` (N HM: mọi HM thiết bị;
  1 HM: theo `HmName1`, không có thì tên dòng bảng tổng; mục "cả file" có chọn riêng: TB chỉ HM thiết bị). `APP_Detect` trả tên HM cho file 1 HM.
  Phần mềm: `Vn.IsTB`, mặc định sheet + "Áp cho mọi file" theo luật; ô TB tô xanh + ● ; tick tay TB cho HM xây dựng → cảnh báo.
- File tổng: `API_TongItems` 13 cột (thêm đường dẫn, dòng THKP, loại, tên HM trung gian, dòng bảng tổng, chuỗi dòng cha `TongParentChain`);
  `APP_TongOne` (xem trước 1 mục = xuất thật riêng mục đó, `API_RunTong` fromLeaf=toLeaf); `APP_HmMap` (dòng THKP → HM k);
  rowSel "seq=0" → `Skip` (không xuất, kể cả mục TB / nguyên sheet). Giá trị số định dạng #,##0 (tránh ###).
- Thiết kế UI do Antigravity làm (BRIEF_2/THIET_KE_2, BRIEF_3/THIET_KE_3). Tab 2: DataGridView vẽ tay thành cây có cột (ô tick 3 trạng thái,
  ▸/▾, chip sheet, trạng thái), xem trước + chọn PDF, bảng tùy chọn mục, hộp báo cáo kết quả + file BAO_CAO_XUAT_*.txt.
- Bố cục: vùng L/C/R mỗi tab có ⧉ tách cửa sổ / « » thu gọn dải 32 px; khung xem trước nổi dùng chung 2 tab (F9); F6/F7/F8, Ctrl+R, F11;
  < 1000 px: thanh trên/dưới tự xuống dòng, ẩn cột File nguồn, tự thu gọn vùng phải; lưu bố cục + độ rộng cột trong caidat.txt.
  Cột cây đổi sang Fill + FillWeight (trước cột tên là Fill duy nhất → kéo hẹp không được).
- Chống sót EXCEL.EXE: Excel nền gắn vào Job Object KILL_ON_JOB_CLOSE (phần mềm văng / bị tắt cưỡng bức → Excel tắt theo; đã thử Stop-Process -Force);
  đóng thường: chờ 8 s rồi Kill; dự phòng file excel_pid.txt. Sửa mục kẹt "Đang tạo…" khi việc xem trước trong hàng đợi bị thay.
- Thử: file tổng THC 34 mục / 23 nhóm; mục 26 (HM LẮP ĐẶT) 26 trang có 15 trang TB; mục 4 (HM XÂY DỰNG) 12 trang không TB;
  xuất thật chỉ tick 4, 23, 26 → 3 PDF. Tham số chạy thử: `InAnG8.exe --tong <file tổng> <sheet> <số mục>`.
- Gói: `dist\In_an_du_toan_G8_V7.1(.zip)`.

## Lần 20 – Bỏ dấu phẩy thừa ở số nguyên (12/11/2026)
- Lỗi: G8 định dạng `#.##0,####` (ĐGCT 5–6 số lẻ) → số nguyên in "1," "4," "8," (Định mức, Hệ số, Số C.K, KL…).
- `FixIntDecimals` (gọi trước AutoFit ở `ApplySettingsToWorksheet`, `ApplyPageSetup2`, và mục "in nguyên sheet" của file tổng):
  đọc cả vùng in 1 lần; ô số có giá trị làm tròn theo số chữ số lẻ của định dạng = số nguyên → gán định dạng bỏ phần `.###`
  (`NfDropHashDecimals`, có bộ nhớ đệm theo định dạng). Số có phần lẻ, định dạng có `0` bắt buộc hoặc chuỗi "…" giữ nguyên. Chỉ trong bộ nhớ.
- Thử: TXL HM1 381 → 0 số thừa dấu phẩy (6 chỗ còn lại là chữ trong diễn giải "mác 100, …"); TXL HM4 193 → 0; 1d HM1 1.247 → 0;
  số có phần lẻ không đổi (618 / 349 / 225). Thời gian xuất không đổi (TXL 56,8 s).
- Phụ: cột số hẹp lại → ĐGCT TXL HM4 chữ to hơn (12,0 → 12,7 pt), thêm 1 trang (19 → 20).
- Đã chép xlsm build mới vào `app\bin`; CHƯA đóng gói lại `dist\` (V7.1 / V6.6).

## Lần 21 – Tăng tốc xem trước / xuất (12/11/2026)
- Đo từng bước (`API_ProfStart` / `API_ProfDump`, `ProfMark` – tắt mặc định, không tốn gì). File TXL 4 HM 55 s, chậm nhất:
  kẻ viền hairline→thin 17 s, gán PageSetup + xoá header/footer 13,6 s, ép hướng giấy 2 lần 7,5 s, xuất PDF 6,8 s.
- Nguyên nhân chính: lõi bật `ScreenUpdating = True` (tránh lỗi 1004 khi Activate) → mỗi lần đổi viền/cột Excel vẽ lại.
  Tắt `ScreenUpdating` trong `HairToThin` và các bước FixIntDecimals / AutoFit / HideColsFinal2 / FixHashCells (không Activate ở đó).
- Header/footer chỉ xoá khi thật sự có chữ (bỏ 12 lệnh EvenPage/FirstPage); nhánh nhiều HM: mỗi sheet xoá 1 lần / file (`mHFDone`).
- `ForceOrientation` / `ForceOri2`: đã đúng A4 + hướng thì không gán lại.
- Viền: `HairToThin` (đọc Weight trước, chỉ hairline mới đọc LineStyle) thay vòng lặp cũ ở 2 nhánh.
  Đã thử và BỎ: hỏi viền theo khối (`Borders(xlInsideHorizontal)` / viền dưới của cả dòng trả giá trị ô đầu, không báo lẫn lộn)
  và Tìm/Thay theo định dạng (không bắt được viền) – cả hai làm mất bước đổi viền mờ.
- Đã thử và BỎ: AutoFit chỉ trong vùng in → đổi độ rộng cột so với bản cũ.
- Kết quả 6 file `Test` (cùng lệnh, đủ sheet): 165,6 s → 81,5 s; TXL 51 → 25 s; 10. ĐM38 12,7 → 8,1 s; 2. ĐM38 38,5 → 14,3 s.
  So từng trang (ảnh 60 dpi): 238/238 trang giống hệt bản trước khi tăng tốc.
- Phần mềm (tab 1): file nhiều HM chỉ dựng HM đang xem (`CustNOnly`: HM khác "0"), xong tự chạy ngầm HM kế tiếp
  (việc "prefetch" trong hàng đợi, bị bỏ khi người dùng bấm việc khác); bản xem trước gộp theo HM, theo `PreviewGen`.
  TXL: HM1 hiện sau ~9 s (trước ~56 s cho cả 4 HM), HM2–4 xong ngầm sau ~40 s.
- Lỗi chưa rõ gốc: trong phần mềm, lần APP_Run thứ 2 ngay sau lần đầu đôi khi làm Excel nền chết (DISP_E_EXCEPTION; ngoài phần mềm
  chạy 24 lần liên tiếp không lỗi; Event Log không có crash). Thêm `Engine.RunRetry`: Excel chết → khởi động lại + thử lại 1 lần
  (không sót EXCEL.EXE; đã kiểm đóng phần mềm → Excel nền tắt). Mất thêm ~3 s khi xảy ra.
- Đã chép `InAnG8.exe` + xlsm vào `app\bin`. CHƯA đóng gói `dist\`, chưa commit.

## Lần 22 – Tên trên bản in: lõi + dữ liệu phần mềm (12/11/2026)
- Quét 6 file Test: dòng tên ở cột A (ô gộp). THKP: "Tên dự án/công trình/hạng mục:" dòng 2–4 mỗi khối (khối ≥2 là `=A2`, `=A3`);
  Giá TH / ĐGCT / Giá tháng / TB: "CÔNG TRÌNH :" = `='Công trình'!A2`, "HẠNG MỤC:" = `='Công trình'!F<dòng HM>`.
- `APP_Detect` thêm cột 7, 8: tên dự án, tên công trình đang có trong file (`ReadLabelText`).
- `API_SetTongTen` nay bật `gNamesFromApi` → `TongLoadUiOptions` không đọc đè bằng ô C8/C9 trống của tool (trước đây phần mềm không đổi tên được).
- Đổi tên HM khi in: `API_SetHmRename "k=tên"` (vbLf). `ApplyHmRenameSheet`: trong vùng in, cột A..F, ô chữ có nội dung (sau 1–4 lớp nhãn "…:")
  trùng tên cũ → tên mới, giữ nhãn, giữ HOA; ô công thức bị thay bằng chữ. Tên cũ = tên THKP + tên ở dòng HM sheet Công trình (`CtHmAlias`).
  Gọi sau dàn trang (nhánh N HM theo từng HM; nhánh 1 HM với HM 1) → nhận khối / luật TB vẫn theo tên gốc. Tên PDF (`HmSuffix`, tick tên HM) dùng tên mới.
- Thử: TXL đổi HM2 + dự án + công trình → mọi sheet đổi đúng, cả "Tổng cộng: HANG MỤC: …"; 9.2b (1 HM) đúng; file 10 không đổi tên → chữ PDF y hệt bản trước.
- Phần mềm (InAnG8.cs): FileItem `DuAnGoc/CongTrinhGoc/DuAnRieng/CongTrinhRieng/HmTen`, `HmRenameString`, `HmGoc`; MainForm `GlobalDuAn/GlobalCongTrinh`
  (lưu caidat `ten_duan`, `ten_congtrinh`), `EffDuAn/EffCongTrinh/ShownDuAn/ShownCongTrinh`, `NamesChanged`, `NameConflicts`, `NamesUiRefresh`;
  `EngSetNames` gửi tên trước mỗi APP_Run (xem trước, xuất) và APP_Tong/APP_TongOne (tên chung, không đổi tên HM).
- Hợp đồng UI mục D; BRIEF_5 giao Antigravity (thanh tên chung, bảng tên lệch – Test có 2 tên công trình: "…LƯƠNG TÀI" 4 file / "…ĐÔNG CỨU 1" 3 file,
  ghi đè từng file, sửa tên HM trên lưới). Đang chờ XONG_5.
- Lưu ý: sửa file .cs bằng Python heredoc làm "\r" thành ký tự CR thật → lỗi biên dịch; viết script ra file (.py) rồi chạy.

## Lần 23 – Giao diện Tên trên bản in (BRIEF_5, Antigravity) (12/11/2026)
- Antigravity: `NamesUi.cs` (mới: thanh "Tên trên bản in" chung 2 tab, hộp "Tên khác nhau giữa các file", vùng "Tên trên bản in (file này)"),
  sửa `Tab1Ui.cs` (sửa tên HM trên lưới / ô tên cho file 1 HM, cây hiện tên mới + "đã đổi tên"), `PreviewUi.cs` (`ResolveTitle`), `LayoutUi.cs` (`lay_namesbar`).
- Biên dịch được ngay. Chạy thật trên 7 file Test: đặt Công trình chung + đổi tên HM2 file TXL trên lưới → xem trước HM2: THKP, Giá TH, Công trình,
  dòng "Tổng cộng" đều tên mới; tiêu đề khung xem trước theo tên mới. Hộp lệch tên: Công trình 2 nhóm (4 / 3 file), Dự án đồng nhất – đúng.
  Đóng phần mềm → Excel nền tắt; tên thử đã xoá (caidat `ten_duan=`, `ten_congtrinh=` trống).
- Còn lệch nhỏ: dòng gợi ý dưới lưới HM bị nút "↺ Gốc tất cả" che chữ cuối; tiêu đề cột lưới HM bị cắt ("ƏGCT", có từ trước).
- Đã chép exe + xlsm vào `app\bin`. CHƯA đóng gói dist, chưa commit.

## Lần 24 – Co kéo cơ động + hiện tên Dự án (BRIEF_5b → LOI_5b → 5c, Antigravity) (12/11/2026)
- 5b: kéo được độ rộng cột lưới HM (nháy đúp = vừa chữ), thanh chia lưới ↔ thứ tự sheet, tỉ lệ 2 ô tên, cột ảnh nhỏ xem trước, chiều cao nhật ký,
  hộp báo cáo / tên lệch; lưu khoá lay_* (lay_names_ratio, lay_gridn_cols, lay_gridn_h, lay_t2_info_h, lay_log_h, lay_prev_thumbs, lay_rep_win,
  lay_rep_split, lay_rep_cols, lay_conf_win). Ô tên trống hiện "Đang in theo file: …"; tiêu đề THKP / ĐGCT hết bị cắt; gợi ý dưới lưới hết bị nút che.
- Lỗi 5b (1440×900): khung "Tên trên bản in (file này)" mở sẵn, ép lưới còn 2/4 HM, thứ tự sheet 3/6, mất nút Xuống / Áp cho mọi file → LOI_5b.md.
- 5c: khung tên mặc định thu gọn (nút tiêu đề bấm thu / mở được), vùng phải tab 1 bọc pnlRightScroll (AutoScroll), lưới tối thiểu đủ số HM (≤ 6),
  thứ tự sheet đủ 6 dòng + 3 nút. Chạy thật 1440×900 (thu gọn / mở rộng → cuộn dọc) và 1000 px: đạt. Ảnh BAO_CAO_2026-11-12\anh\c5_*.png.
- Bản trước 5b lưu ở UI_PHAN_MEM_IN_G8\_BAN_TRUOC_5b\. Đã chép exe vào app\bin (xlsm không đổi). CHƯA đóng gói dist, chưa commit.

## Lần 25 – Việc 4: thứ tự sheet + hướng dọc / ngang riêng từng mục file tổng (12/11/2026)
- Lõi VBA: chuỗi chọn sheet nhận hậu tố hướng "1N,4D,2" (N ngang, D dọc, không hậu tố = mặc định). `TongParseSel` bỏ hậu tố;
  `TongParseOri` đọc hướng; `gOriN(sheet, HM)`, `gOri1(sheet)`, `gOrdNH(sheet, HM)` (thứ tự riêng từng HM), `gCurHm`; `OriPick2/OriPick1`
  trong ForceOri2 / ApplyPageSetup2 / ForceOrientation; `ResetOriOrd` đầu APP_Run, đầu + cuối mỗi file trong file tổng.
  Tab 1 (custN) cũng nhận thứ tự / hậu tố hướng (chưa có UI).
- C#: `TongItem.Ord`, `Ori`, `EffOrder/EffOri/DefOri/HasCustomLayout/ResetLayout`, `SelString()` gửi hậu tố; `ItemLayoutChanged`;
  TongApplyAll chép cả thứ tự + hướng; TongDefaults xoá. HOP_DONG_UI mục E.
- Kiểm: gọi lõi trực tiếp (4 HM: HM1 CT ngang → ĐGCT → THKP; HM2 GTH dọc → THKP; 1 HM: ĐGCT ngang → THKP dọc) – PDF đúng thứ tự, đúng hướng.
  Hồi quy mặc định 7 file Test: 238/238 trang giống hệt bản cũ, ~98 s.
- UI (Antigravity BRIEF_6 + LOI_6 → 6b): danh sách 6 sheet theo thứ tự, nút Dọc/Ngang (khác mặc định tô vàng "Dọc*"), ▲/▼, "★ Đã chỉnh · ↺ Về mặc định",
  cây hiện "THKP·D"; vùng phải tab 2 rộng hơn (DefR 0.23). Chạy thật file tổng 260709.TH DZ TBA - LT.xlsx (34 mục): mục 04 THKP dọc → PDF trang 1 dọc.
- Đã chép exe + xlsm vào app\bin (và xlsm gốc In-an-G8\). CHƯA đóng gói dist, chưa commit.
- Còn nhỏ: khi hiện dòng "★ Đã chỉnh", nút "Cập nhật xem trước" bị đẩy xuống mép dưới (phải cuộn) ở 1440×900.

## Lần 26 – Việc 2: ô hiện ### (12/11/2026)
- Hiện trạng: lõi đã có 3 lớp (AutoFit cột số theo từng sheet → FixHashCells nới cột tới khi hết ### → chừa 12 % hao phông khi in).
  Quét 238 trang PDF 7 file Test + PDF xem trước file tổng: 0 chỗ "###".
- Sửa: (1) `TongExportSheetJob` (mục file tổng "in nguyên sheet" / TB) trước đây KHÔNG xử lý ### → thêm FixHashCells.
  (2) FixHashCells bước 1 nới cột có giới hạn `HASH_MAX_GROW` = +30 % (ít nhất +3); quá giới hạn vẫn ### → tắt xuống dòng + ShrinkToFit RIÊNG ô đó
  (không làm nhỏ cả trang). (3) Nhật ký `gHashLog` / `API_HashLog()`; phần mềm ghi vào Nhật ký sau mỗi file xuất và vào báo cáo xuất file tổng.
- Thử ép: số 98.765.432.109.876 vào ô A9 sheet Công trình (cột 5,7): cột nới 5,7 → 8,7 rồi thu chữ riêng ô → PDF hiện đủ số.
- BÀI HỌC: KHÔNG dùng `Set-Content -Encoding ascii` (PowerShell) để sửa src – làm hỏng 2 dòng có dấu trong HmSuffix ("Lắp đặt" → "L???p")
  → tên PDF có "?" → lỗi 1004 "Document not saved" ở file có HM Lắp đặt / Xây dựng. Đã khôi phục từ git HEAD.
- Hồi quy: 238 trang, nội dung giống; riêng bảng ĐGCT file TXL đôi khi lệch ngang 3 pt GIỮA 2 LẦN CHẠY CÙNG BẢN (nhiễu có sẵn, không do sửa lần này).
- Đã chép exe + xlsm vào app\bin. Chưa đóng gói dist, chưa commit.

## Lần 27 – Bỏ hẳn form VBA cũ, đưa phần mềm ra ngoài (12/11/2026)
- Người dùng: form cũ (nền xanh, nút cam, MO_FORM_IN_AN.bat) là tool cũ, phải trình bản mới ra.
- Lõi: `BuildInAnForm_Engine` (src\BuildInAnForm.bas) bỏ `frmInDuToan` + `clsInAnBtn`, xoá mọi nút trên sheet; sheet 1 đổi tên
  "PHAN MEM IN AN G8", chỉ ghi "mở bằng InAnG8.exe"; sheet `IN THEO FILE TONG` VeryHidden (giữ giá trị, lõi còn đọc C12/C13).
  `MoBangDieuKhien` chỉ báo "form cũ đã ngừng dùng". Xoá src\frmInDuToan_code.txt, src\clsInAnBtn.cls; _gen.py/_gen.ps1 bỏ 2 file đó.
- BUILD_V6.ps1 viết lại: chỉ còn chế độ bộ máy, khung mặc định = FILE_TONG_IN_G8_V7.1.xlsm cạnh script, `-Out` bắt buộc và phải khác khung.
  Build lại từ chính bản bộ máy ra kết quả y hệt (đã so nội dung 2 sheet). build_app.ps1 mặc định ra gốc G8\ (bỏ app\bin).
- Thư mục: InAnG8.exe + FILE_TONG_IN_G8_V7.1.xlsm nằm ở gốc (có trong git). MO_FORM_IN_AN.bat/.vbs, xlsm V6.4/6.5/6.6, video HD cũ,
  gói dist V6.5–V7.1 cũ, xlsm V7.1 còn form → `_LUU_TRU_FORM_CU\` (gitignore, không xoá hẳn).
- README, HUONG_DAN (viết lại cho phần mềm), DOC_TRUOC (bỏ "THỬ NGHIỆM", ghi form cũ ngừng dùng + link Releases).
- Hồi quy 7 file Test: 238/238 trang, số trang giống; chỉ ĐGCT TXL lệch 3 pt như đã biết (nhiễu có sẵn, người dùng không cần sửa).
- GitHub: gộp vào `main` của fork, bản phát hành v7.1-261112 bỏ gói V6.6, thay gói V7.1 mới.

## Lần 28 – Video demo dự án OSI + sửa tick cha–con trên cây file tab 1 (15/11/2026)
- Chép `M:\DUTOAN\2024\1. OSI\2. OUTPUT\03. TKTC\240117. TDT` về `In an du toan\Test_OSI\` (33 MB, 32 file), chỉ chạy trên bản chép.
  Xuất thử 15 file → 22 PDF, 982 trang, 0 ô "###". `2401011 TH KC.xlsx` nhận "1 HM" nhưng không ra PDF (file tổng hợp, không có sheet dự toán chi tiết).
  File tổng `231214. TDT.xlsm` / `2401011 TH KC.xlsx` link ra `240112 TH KC NGAM VA THAN` (ổ M) và ổ D → chưa demo tab file tổng.
- Video `In an du toan\BAO_CAO_2026-11-15\demo_phan_mem_in_an_OSI.mp4` (2'33"): chụp riêng cửa sổ phần mềm bằng PrintWindow (máy đang có UltraViewer
  HUY01 kết nối, khung UltraViewer + thông báo Zalo lọt vào bản quay cả màn hình). Kịch bản ở scratchpad (`rec8.ps1`, `cap_loop.ps1`).
  BÀI HỌC quay: PS 5.1 đọc .ps1 không BOM là ANSI (chữ Việt hỏng) → ghi BOM; Unikey làm hỏng SendKeys chữ (Test_OSI → Tét_ÓI) → dán clipboard;
  chữ WinForms lộ ra UIA là Pane chứ không phải Text; chờ theo file PDF xem trước thay vì giây cố định; biến PowerShell không phân biệt hoa thường ($H đè $h).
- Lỗi phát hiện khi quay: tick HM không tick ô file (nút báo "0 file"), tick ô file không tick HM con.
  Lõi (Claude): `OnTreeCheck` tick ô file nhiều HM = tick mọi HM; `SyncFileCheck(fi)` ô file bật khi còn ≥ 1 HM (gọi cả ở `ReadN`).
  UI (Antigravity BRIEF_7 → XONG_7): ô "một phần" (CheckBoxRenderer MixedNormal), nhãn "1/2 HM", tooltip, bấm ô một phần = chọn đủ.
  Claude sửa 3 lỗi biên dịch (ContentAlignment trùng tên khi thêm using VisualStyles; CheckBoxRenderer.IsSupported không có ở .NET 4 → Application.RenderWithVisualStyles).
  Kiểm bằng chuột thật: bỏ HM2 → "1/2 HM" + vạch; bấm ô → đủ; bỏ ô file → 14 file; Bỏ hết → 0 file; Chọn hết → 15 file · 22 PDF.
