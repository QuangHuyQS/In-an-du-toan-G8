# HỢP ĐỒNG GIAO DIỆN ↔ LÕI – phần mềm In ấn dự toán G8 (bản 7.1)

`MainForm` là `partial class`, chia file:

| File | Ai sửa | Nội dung |
|---|---|---|
| `InAnG8.cs` | **Claude** (lõi) | Program, Engine (Excel COM, chống Excel ngầm), PdfDoc, FileItem, TongItem, TNode, Vn (luật TB), MainForm: **khai báo control dùng chung**, xử lý dữ liệu, quét link, xem trước, xuất, cài đặt |
| `Tab1Ui.cs` | **Antigravity** | `MkBtn`, `MkHdr`, `BuildUi`, `ToggleLog`, `BuildTab1`, `ShowOptions` |
| `Tab2Ui.cs` | **Antigravity** | `BuildTab2`, cây có cột (`RebuildGrid`, `DgvT_CellPainting`, `DgvT_CellMouseClick`, `HmBadge`, `StatusText`), `ShowItem`, `UpdateTBNote`, `UpdatePrevButton`, `UpdateTongButton`, `ShowReport` |
| `LayoutUi.cs` | **Antigravity** | Tách vùng / thu gọn / co giãn / lưu bố cục (`InitZones`, `ApplyLayout`, `TogglePreviewFloat`, `Responsive`, `ResetLayout`, `LayoutLines`, `RestoreLayout`, `RestoreLayout2`, `MkBar`…) |
| `PreviewUi.cs` | **Antigravity** | Lớp `PreviewPane` (khung xem trước PDF) |

Quy tắc:
1. **C# 5, .NET Framework 4, WinForms** (biên dịch bằng `csc.exe` có sẵn): KHÔNG `$"..."`, `?.`, `nameof`, `=>` cho thân hàm/thuộc tính, `out var`, tuple, `async/await`; không NuGet / thư viện ngoài / WebView2.
2. Chú thích trong code: ASCII không dấu. Chữ hiển thị: tiếng Việt có dấu.
3. **Không đổi tên / kiểu / xóa** các control và hàm ở mục A, B dưới đây (lõi dùng). Được thêm control, hàm, field mới **trong file của mình** (field mới khai báo trong file Ui).
4. Không gọi Excel / COM / `eng` trực tiếp; không đọc / ghi file ngoài `caidat.txt` (qua `LayoutLines` / `layCfg`).
5. Mọi thay đổi xong: Claude biên dịch (`app\build_app.ps1`) + chạy thử thật. Ghi rõ trong XONG những hàm đã sửa.

## A. Control lõi dùng – UI phải TẠO (gán vào field có sẵn ở InAnG8.cs)
Chung: `tabs`, `prg`, `lblStatus`, `pnlLog`, `txtLog`, `btnLogToggle`, `prev1`, `prev2` (PreviewPane).
Tab 1: `tree` (TreeView có CheckBoxes; sự kiện → `OnTreeSelect()`, `OnTreeCheck`), `pnl1`, `pnlN`, `pnlNone`, `lblOptTitle`, `clb1` (ItemCheck → `Read1`), `gridN` (cột: on, tt, nm, s1..s6; CellValueChanged → `ReadN`; ColumnHeaderMouseClick → `ToggleColumn`), `lstOrderN`, `chkTongCot`, `chkTenHM`, `chkOpenFolder`, `btnPreview`, `btnExport`, `btnStop`, `txtOut`, `optTach`, `optGhep`, `optPrint`.
Tab 2: `txtSum`, `cboSumSheet`, `btnTongDry`, `chkTongSheet[1..6]`, `chkTongCot2`, `lblScanSummary`, `txtSumOut`, `chkOpen2`, `lblSelSummary`, `btnTongRun`, `btnReport`, `txtSearch`, `radAll`, `radErr`, `dgvT` (cột Name: chk, h, val, src, hm, sh, st), `lblGridWait`, `cboItemPdf`, `pnlPdfSel`, `pnlItem`, `lblItemNone`, `lblItemTitle`, `grpInfo`, `lblItemInfo`, `grpSheets`, `chkItem[1..6]` (CheckedChanged → `ReadItemSel(k)`), `lblTBRule`, `pnlTBWarn`, `lblTBWarn`, `btnTongPrev`, `btnApplyAll`, `btnResetItem`.
Bố cục: `lay1`, `lay2` (TabLayout: Page, SpL, SpR) – BuildTab1/2 phải gán.

## B. Hàm lõi UI được gọi
Tab 1: `PickFiles()`, `PickFolder(start)`, `AddPaths(paths)`, `RemoveSelected()`, `RequestPreview(fi, true)`, `Move1(±1)`, `Preset1(ids)`, `ApplyAll1()`, `MoveN(±1)`, `ApplyAllN()`, `MarkAllStale()`, `StartExport()`, `UpdateExportButton()`.
Tab 2: `PickSum()`, `TongScan()`, `TongExport()`, `TongPreview(it, true)`, `OpenItemPdf(it, idx)`, `TongApplyAll()`, `TongDefaults(it)`, `MarkItemStale(it)`, `TongCheckAll(bool)`, `ToggleCheck(node)`, `SetExpandAll(bool)`, `OnTongSelect()`, `ReadItemSel(k)`, `SrcOf(it)`, `ItemIsTB(it, fi[, out kw])`, `HmText(it)`, `ItemTitle(it)`, `CheckState(node)`, `Money(s)`, `MoneyVal(s)`, `LeafVisible(it, q)`.
Chung: `SetStatus(s)`, `AddLog(s)`, `Ui(action)`, `ToInt(s)`, `Vn.IsTB(tên)`, `Vn.NormUp(s)`.
Dữ liệu đọc được: `files` (List<FileItem>), `curFile`, `curHm`, `tongItems` (List<TongItem>), `tongRoots` (List<TNode>), `curItem`, `curNode`, `tongSum`, `lastReport`, `lastOutDir`, `loadingUi` (đặt true khi nạp control để chặn sự kiện), `exporting`, `tongBusy`, `eng != null`.

## C. Hàm UI mà lõi gọi (giữ nguyên chữ ký)
`BuildUi()`, `ToggleLog(bool)`, `ShowOptions(FileItem)`, `MkBtn(text, back, fore, w)`, `MkHdr(text)`, `RebuildGrid()`, `ShowItem(TongItem)`, `UpdateTBNote(TongItem)`, `UpdatePrevButton()`, `UpdateTongButton()`, `ShowReport()`, `InitZones()`, `RestoreLayout()`, `RestoreLayout2()`, `ResetLayout()`, `LayoutLines()` (IEnumerable<string> "lay_*=..."), `layCfg` (Dictionary nạp từ caidat.txt), `ActiveLay`, `TogglePreviewFloat()`, `ToggleCollapse(t, zone, manual)`.
PreviewPane: `TitleText`, `Banner(text|null)`, `Clear()`, `Loading(text)`, `Open(pdfPath, title)`, `NextPage()`, `PrevPage()`.

## D. Tên trên bản in (thêm 12/11/2026 – BRIEF_5)
Lõi đã làm xong phần dữ liệu + gửi tên cho Excel trước mỗi lần xem trước / xuất (cả tab 1 lẫn tab file tổng). UI chỉ đọc / ghi các field dưới đây rồi gọi `NamesChanged`.

Dữ liệu (đọc / ghi):
- `GlobalDuAn`, `GlobalCongTrinh` (string, MainForm): tên chung áp cho MỌI file + tab file tổng. `""` = giữ tên trong từng file. Lõi tự lưu vào caidat.txt.
- `FileItem.DuAnGoc`, `FileItem.CongTrinhGoc` (chỉ đọc): tên đang ghi trong file (đọc lúc nhận diện, có thể `""`).
- `FileItem.DuAnRieng`, `FileItem.CongTrinhRieng`: ghi đè riêng file này (`""` = theo tên chung).
- `FileItem.HmTen` (List<string>, số phần tử = số HM; file 1 HM có 1 phần tử): tên HM mới khi in, `""` = giữ tên gốc. Tên gốc: `fi.HmGoc(h)` (h tính từ 0). Tên mới dùng cho bản in VÀ tên PDF (khi bật "Tên PDF kèm tên hạng mục").

Hàm lõi UI được gọi:
- `NamesChanged(fi)`: gọi SAU khi sửa tên. `fi = null` → đổi tên chung (đánh dấu mọi file + mọi mục file tổng cần xem trước lại, lưu cài đặt); `fi` → chỉ file đó.
- `EffDuAn(fi)`, `EffCongTrinh(fi)`: tên sẽ GỬI cho lõi (`""` = giữ tên file). `ShownDuAn(fi)`, `ShownCongTrinh(fi)`: tên sẽ IN RA (đã tính ghi đè / tên gốc) – để hiển thị.
- `NameConflicts()` → `List<NameGroup>` (`What` 1 = dự án, 2 = công trình; `Value` = tên gốc; `Files` = các file có tên đó). Chỉ trả về loại có ≥ 2 tên gốc khác nhau giữa các file đã nhận diện.

Lõi gọi UI:
- `NamesUiRefresh` (field `Action`, UI gán trong BuildUi): lõi gọi sau mỗi lần nhận diện xong 1 file và sau khi bỏ file khỏi danh sách → UI vẽ lại vùng tên / bảng tên lệch.

Không đổi: luật TB, nhận diện khối HM vẫn theo tên GỐC (đổi tên không làm đổi việc in sheet TB).

## E. Thứ tự + hướng in riêng từng mục file tổng (thêm 12/11/2026 – BRIEF_6)
Lõi đã xong (Excel + C#), đã kiểm PDF thật. UI chỉ đọc / ghi 2 field của `TongItem` rồi gọi `ItemLayoutChanged(it)`.

Dữ liệu (`TongItem`):
- `Ord` (`List<int>`, đủ 6 mã sheet 1..6 theo thứ tự in; `null` = mặc định `TongItem.Order` = TB, THKP, GTH, ĐGCT, CT, GT). Thứ tự áp cho các sheet đang tick (`Sel[k]`); sheet bỏ tick vẫn nằm trong `Ord` (giữ chỗ).
- `Ori` (`int[7]`, chỉ số 1..6): 0 = mặc định, 1 = dọc, 2 = ngang.
- Đọc: `EffOrder()` (thứ tự đang dùng), `EffOri(k)` (1 dọc / 2 ngang sẽ in), `TongItem.DefOri(k)` (hướng mặc định: THKP, GTH, TB ngang; ĐGCT, CT, GT dọc), `HasCustomLayout()` (có thứ tự / hướng riêng).
- `ResetLayout()`: về thứ tự + hướng mặc định (lõi đã gọi trong `TongDefaults`, tức nút "Về mặc định ban đầu").

Hàm lõi UI gọi: `ItemLayoutChanged(it)` sau khi sửa `Ord` / `Ori` (đánh dấu xem trước cũ). `TongApplyAll()` đã chép cả `Ord` + `Ori` sang các mục khác.
Áp dụng: file nhiều HM (mỗi mục = 1 HM, thứ tự / hướng riêng từng HM) và file 1 HM. Mục loại "chỉ in sheet TB" / "in nguyên sheet" (`CanSel = false`) không áp.
