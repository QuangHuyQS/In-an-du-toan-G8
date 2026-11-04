Attribute VB_Name = "BuildInAnForm"
Option Explicit

' ============================================================================
'  BuildInAnForm - CHAY 1 LAN de dung lai form frmInDuToan + class clsInAnBtn.
'  Can: modInAnMain da co trong project, bat "Trust access to the VBA project".
'  Goi: BuildInAnForm_Run "<duong dan code form .txt>", "<duong dan clsInAnBtn.cls>"
'  (script BUILD_V6.ps1 tu goi; build xong module nay bi xoa khoi file).
' ============================================================================

Private mComp As Object

Public Sub BuildInAnForm_Run(codePath As String, clsPath As String)
    Dim vbp As Object: Set vbp = ThisWorkbook.VBProject
    RemoveComp vbp, "frmInDuToan"
    RemoveComp vbp, "clsInAnBtn"

    Set mComp = vbp.VBComponents.Add(3)
    DoEvents
    mComp.Name = "frmInDuToan"
    mComp.Properties("Caption") = "Tool in an du toan G8 / ESCON"   ' thanh tieu de chi hien ANSI
    mComp.Properties("Width") = 730
    mComp.Properties("Height") = 600
    mComp.Properties("BackColor") = ClrBg()
    mComp.Properties("StartUpPosition") = 1

    BuildControls

    mComp.CodeModule.AddFromFile codePath
    ' AddFromFile giu dong "Option Explicit" trung neu VBE tu chen -> xoa dong trung
    DedupOptionExplicit mComp.CodeModule

    vbp.VBComponents.Import clsPath
End Sub

' Nut mo form + huong dan tren sheet BANG DIEU KHIEN
Public Sub BuildInAnForm_Sheet()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(1)
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.Name = "btnMoForm" Or shp.Name = "btnXuatTong" Then shp.Delete: Exit For
    Next shp
    For Each shp In ws.Shapes
        If shp.Name = "btnMoForm" Or shp.Name = "btnXuatTong" Then shp.Delete: Exit For
    Next shp

    Dim r As Range: Set r = ws.Range("B5:D7")
    r.UnMerge
    r.ClearContents
    r.Merge
    Set shp = ws.Shapes.AddShape(5, r.Left + 6, r.Top + 4, 360, r.Height - 8)   ' rounded rectangle
    With shp
        .Name = "btnMoForm"
        .OnAction = "MoBangDieuKhien"
        .Fill.ForeColor.RGB = ClrOrange()
        .Line.Visible = False
        With .TextFrame2
            .VerticalAnchor = 3                        ' msoAnchorMiddle
            .TextRange.ParagraphFormat.Alignment = 2   ' msoAlignCenter
            .TextRange.Text = "▶   MỞ FORM IN ẤN DỰ TOÁN"
            .TextRange.Font.Size = 14
            .TextRange.Font.Bold = True
            .TextRange.Font.Name = "Segoe UI"
            .TextRange.Font.Fill.ForeColor.RGB = ClrOnOrange()
        End With
    End With

    Dim shp2 As Shape
    Set shp2 = ws.Shapes.AddShape(5, shp.Left + shp.Width + 12, shp.Top, 330, shp.Height)
    With shp2
        .Name = "btnXuatTong"
        .OnAction = "XuatTheoFileTong"
        .Fill.ForeColor.RGB = ClrGreen()
        .Line.Visible = False
        With .TextFrame2
            .VerticalAnchor = 3
            .TextRange.ParagraphFormat.Alignment = 2
            .TextRange.Text = "▶   XUẤT PDF THEO FILE TỔNG HỢP"
            .TextRange.Font.Size = 14
            .TextRange.Font.Bold = True
            .TextRange.Font.Name = "Segoe UI"
            .TextRange.Font.Fill.ForeColor.RGB = ClrOnOrange()
        End With
    End With

    ws.Range("J13:L20").Clear                          ' so lieu + khung thua dan nham vao file khung

    ws.Range("B9").Value = "  1.  CÁCH DÙNG"
    ws.Range("C10").Value = "Bấm nút cam phía trên (hoặc Alt+F8 ▸ MoBangDieuKhien) để mở form."
    SetRow ws, 11, "Một file", "Chọn file dự toán đang mở (hoặc Mở file...). Form tự nhận 1 hay nhiều hạng mục."
    SetRow ws, 12, "Cả thư mục", "Xử lý hàng loạt, mỗi file tự nhận diện. PDF lưu cạnh từng file gốc."
    SetRow ws, 13, "Tick sheet + số TT", "Số TT quyết định thứ tự trang trong PDF. Bấm vào tên sheet để tick nhanh."
    ws.Range("C14").Value = "Lưu PDF vào: để trống = cạnh file gốc. Form nhớ mọi lựa chọn cho lần sau."
    ws.Range("C15").Value = "Kết quả từng lần chạy hiện ở ô nhật ký cuối form (dòng mới nhất ở trên cùng)."
    ws.Range("C16").Value = "File tổng này không bao giờ tự in chính nó."

    ws.Range("B18").Value = "  2.  TÊN FILE PDF"
    ws.Range("C19").Value = "1 hạng mục — một file:  <tên excel>_FULL.pdf   |   cả thư mục:  <tên excel>.pdf"
    ws.Range("C20").Value = "Nhiều hạng mục:  <tên excel>_HM1.pdf … _HMn.pdf  (hoặc -<tên hạng mục>.pdf)   |   ghép:  <tên excel>.pdf"
    ws.Range("C21").Value = "Ghép: Excel xuất mọi hạng mục trong một lần — không cần phần mềm ghép PDF. Form mở mà vẫn làm việc với Excel được."

    ws.Range("B23").Value = "  3.  CÁC NÚT TRÊN FORM"
    SetRow ws, 24, "Căn chỉnh", "Chỉ căn trang (khổ, hướng, vùng in, ẩn cột/dòng) — không in. File để chưa lưu."
    SetRow ws, 25, "Xem trước", "Print Preview các sheet đã chọn (file 1 hạng mục). Đóng preview là form hiện lại."
    SetRow ws, 26, "In máy in", "In ra máy in mặc định. File nhiều hạng mục chỉ xuất PDF."
    SetRow ws, 27, "XUẤT PDF", "Xuất PDF theo các sheet đã tick."
    SetRow ws, 28, "Đóng / Dừng", "Khi đang chạy cả thư mục, nút Đóng thành Dừng: dừng sau file hiện tại."
    SetRow ws, 29, "Xuất theo file tổng", "Nút xanh trên sheet (hoặc Alt+F8 ▸ XuatTheoFileTong): chọn file tổng hợp (vd 0.KSDT DD D.TOAN T.HOP); tool đi theo link tới từng file dự toán + hạng mục, xuất mỗi hạng mục 1 PDF <NN>_<STT>_<tên>.pdf theo đúng thứ tự file tổng. Thứ tự sheet: TH chi phí TB, THKP, Giá tổng hợp, Đơn giá CT, Giá tháng. Mặc định lưu vào thư mục PDF cạnh file tổng, kèm _THU_TU_XUAT.txt."

    ws.Range("B31").Value = "  4.  LỖI THƯỜNG GẶP"
    SetRow ws, 33, "Báo thiếu form frmInDuToan", "Chạy lại V6\BUILD_V6.ps1 để dựng lại form."
    SetRow ws, 34, "Chạy cả thư mục có file lỗi", "Danh sách lỗi lưu LOI_IN_<ngày giờ>.txt trong thư mục đã chọn."
    SetRow ws, 35, "Không nhận ra sheet dự toán", "File không theo mẫu G8 / ESCON (tên sheet / tiêu đề khác)."
    ws.Range("B37").Value = "  Mã nguồn: thư mục V6\src (form = frmInDuToan_code.txt, engine = inanDTG8_V6.bas). Dựng lại: V6\BUILD_V6.ps1."
    ' chu dai -> xuong dong, gian chieu cao dong cho du
    ws.Range("C10:D37").WrapText = True
    ws.Range("C10:C37").Font.Name = "Segoe UI"
    Dim rr As Long
    For rr = 10 To 37
        If Not ws.Cells(rr, 3).MergeCells And Not ws.Cells(rr, 2).MergeCells Then ws.Rows(rr).AutoFit
    Next rr
    ws.Range("D29").Value = "Nút xanh: chọn file tổng hợp → sang sheet IN THEO FILE TONG để chọn sheet cần in (file 1 hạng mục / nhiều hạng mục), thay tên dự án / công trình, rồi bấm BẮT ĐẦU XUẤT. Tool đi theo link nhiều cấp, PDF đặt trong thư mục con giống thư mục file nguồn, kèm _THU_TU_XUAT.txt."
    ws.Rows(29).AutoFit
    BuildTongCfgSheet
    ws.Activate
    ws.Range("A1").Select
End Sub

' ============================================================================
'  Sheet cau hinh "IN THEO FILE TONG" (V6.5)
'  Doc boi modInAnMain.XuatTong_BatDau - giu nguyen dia chi o (TONG_CFG_*).
' ============================================================================
Private Sub BuildTongCfgSheet()
    Const SH As String = "IN THEO FILE TONG"
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ws As Worksheet, shp As Shape
    On Error Resume Next
    Set ws = wb.Worksheets(SH)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = SH
    End If
    For Each shp In ws.Shapes: shp.Delete: Next shp
    ws.Cells.Clear
    ws.Cells.Validation.Delete
    ws.Cells.Font.Name = "Segoe UI": ws.Cells.Font.Size = 10
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 34
    ws.Columns("C").ColumnWidth = 22
    ws.Columns("D").ColumnWidth = 24
    ws.Columns("E").ColumnWidth = 62
    ws.Columns("F").ColumnWidth = 2
    ws.Columns("G").ColumnWidth = 30

    With ws.Range("B2")
        .Value = "XUẤT PDF THEO FILE TỔNG HỢP — CẤU HÌNH"
        .Font.Size = 14: .Font.Bold = True: .Font.Color = ClrOrange()
    End With
    ws.Range("B3").Value = "Chọn file tổng → chọn sheet cần in + tên dự án → bấm BẮT ĐẦU XUẤT. Các ô trắng được nhớ cho lần sau."
    ws.Range("B3").Font.Color = RGB(100, 116, 139)

    Dim lbls As Variant, i As Long
    lbls = Array("File tổng hợp", "Thư mục lưu PDF  (trống = thư mục PDF cạnh file tổng)", _
                 "Chạy từ mục số  (1 = từ đầu)", "Chạy đến mục số  (0 = đến hết)", _
                 "Tên dự án in ra  (trống = giữ như file nguồn)", "Tên công trình in ra  (trống = giữ như file nguồn)", _
                 "In cột Tổng cộng ở Giá tổng hợp", "Chỉ dò link, KHÔNG xuất PDF (xem trước)")
    For i = 0 To UBound(lbls)
        ws.Cells(4 + i, 2).Value = lbls(i)
        ws.Range(ws.Cells(4 + i, 3), ws.Cells(4 + i, 5)).Merge
        ws.Cells(4 + i, 3).HorizontalAlignment = xlLeft
        With ws.Range(ws.Cells(4 + i, 3), ws.Cells(4 + i, 5))
            .Interior.Color = RGB(255, 255, 255)
            .Borders.Color = RGB(203, 213, 225)
            .Locked = False
        End With
    Next i
    ws.Range("B4:B11").Font.Bold = True
    ws.Range("C6").Value = 1
    ws.Range("C7").Value = 0
    ws.Range("C10").Value = "Không"
    ws.Range("C11").Value = "Không"
    With ws.Range("C10:C11").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Có,Không"
    End With
    ws.Range("C8:C9").WrapText = True
    ws.Rows("8:9").RowHeight = 30

    ' ---- bang chon sheet ----
    ws.Range("B13").Value = "SHEET"
    ws.Range("C13").Value = "File 1 hạng mục" & vbLf & "thứ tự in"
    ws.Range("D13").Value = "File 2, 3… hạng mục" & vbLf & "thứ tự in"
    ws.Range("E13").Value = "Ghi chú"
    With ws.Range("B13:E13")
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(51, 65, 85)
        .WrapText = True: .VerticalAlignment = xlCenter
    End With
    ws.Range("C13:D13").HorizontalAlignment = xlCenter
    ws.Rows(13).RowHeight = 32

    Dim shNames As Variant, ord1 As Variant, ordN As Variant, notes As Variant
    shNames = Array("TH chi phí TB", "THKP hạng mục", "Giá tổng hợp", "Đơn giá chi tiết", "Công trình", _
                 "Giá tháng", "Tổng hợp VT", "TH chi phí XD")
    ord1 = Array(1, 2, 3, 4, "", 5, "", "")
    ordN = Array(1, 2, 3, 4, "", 5, "—", "—")
    notes = Array("Chỉ in khi tổng thiết bị khác 0. Nhiều hạng mục: đi kèm hạng mục đầu.", _
               "Có dòng Tên dự án. Nhiều hạng mục: đúng khối của từng hạng mục.", _
               "Nhiều hạng mục: tách đúng khối của từng hạng mục.", _
               "Nhiều hạng mục: tách đúng khối của từng hạng mục.", _
               "Nhiều hạng mục: tách theo dòng HM / THM ở cột E.", _
               "Bảng chung: in kèm MỌI hạng mục.", _
               "Chỉ file 1 hạng mục.", "Chỉ file 1 hạng mục.")
    For i = 0 To 7
        ws.Cells(14 + i, 2).Value = shNames(i)
        ws.Cells(14 + i, 3).Value = ord1(i)
        ws.Cells(14 + i, 4).Value = ordN(i)
        ws.Cells(14 + i, 5).Value = notes(i)
    Next i
    With ws.Range("B14:E21")
        .Borders.Color = RGB(203, 213, 225)
        .VerticalAlignment = xlCenter
    End With
    ws.Range("C14:D21").HorizontalAlignment = xlCenter
    ws.Range("C14:D21").Interior.Color = RGB(255, 255, 255)
    ws.Range("C14:C21,D14:D19").Locked = False
    ws.Range("D20:D21").Interior.Color = RGB(226, 232, 240)
    ws.Range("D20:D21").Font.Color = RGB(148, 163, 184)
    ws.Range("E14:E21").Font.Color = RGB(100, 116, 139)
    ws.Range("E14:E21").WrapText = True
    With ws.Range("C14:C21,D14:D19").Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="1", Formula2:="20"
        .ErrorMessage = "Nhập số thứ tự in (1, 2, 3…) hoặc để trống = không in."
    End With
    ws.Range("B23").Value = "Ô thứ tự để TRỐNG = không in sheet đó. Số nhỏ in trước. Sheet nào file nguồn không có thì tự bỏ qua."
    ws.Range("B24").Value = "File 1 hạng mục: cả file 1 PDF.   File 2, 3… hạng mục: mỗi hạng mục 1 PDF, các sheet trên được tách đúng phần của hạng mục đó."
    ws.Range("B23:B24").Font.Color = RGB(100, 116, 139)

    ' ---- nut ----
    Dim r As Range
    Set r = ws.Range("G4:G5")
    Set shp = ws.Shapes.AddShape(5, r.Left, r.Top + 2, r.Width, r.Height - 4)
    With shp
        .Name = "btnTongChonFile": .OnAction = "XuatTong_ChonFile"
        .Fill.ForeColor.RGB = RGB(71, 85, 105): .Line.Visible = False
        With .TextFrame2
            .VerticalAnchor = 3: .TextRange.ParagraphFormat.Alignment = 2
            .TextRange.Text = "Chọn file tổng hợp…"
            .TextRange.Font.Size = 11: .TextRange.Font.Bold = True: .TextRange.Font.Name = "Segoe UI"
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        End With
    End With
    Set r = ws.Range("G13:G16")
    Set shp = ws.Shapes.AddShape(5, r.Left, r.Top + 2, r.Width, r.Height - 4)
    With shp
        .Name = "btnTongBatDau": .OnAction = "XuatTong_BatDau"
        .Fill.ForeColor.RGB = ClrGreen(): .Line.Visible = False
        With .TextFrame2
            .VerticalAnchor = 3: .TextRange.ParagraphFormat.Alignment = 2
            .TextRange.Text = "▶  BẮT ĐẦU XUẤT"
            .TextRange.Font.Size = 14: .TextRange.Font.Bold = True: .TextRange.Font.Name = "Segoe UI"
            .TextRange.Font.Fill.ForeColor.RGB = ClrOnOrange()
        End With
    End With
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ws.Range("C4").Select
End Sub

Private Sub SetRow(ws As Worksheet, r As Long, c As String, d As String)
    ws.Cells(r, 3).Value = c
    ws.Cells(r, 4).Value = d
End Sub

Private Sub RemoveComp(vbp As Object, nm As String)
    Dim c As Object
    On Error Resume Next
    Set c = vbp.VBComponents(nm)
    On Error GoTo 0
    If Not c Is Nothing Then vbp.VBComponents.Remove c: DoEvents
End Sub

Private Sub DedupOptionExplicit(cm As Object)
    Dim i As Long, seen As Boolean
    For i = cm.CountOfDeclarationLines To 1 Step -1
        If Trim$(cm.Lines(i, 1)) = "Option Explicit" Then
            If seen Then cm.DeleteLines i, 1 Else seen = True
        End If
    Next i
End Sub

' ============================================================================
'  BO CUC  (don vi point; vung trong rong 720)
' ============================================================================
Private Sub BuildControls()
    Dim i As Integer, y As Double

    ' ---- dai tieu de ----
    Fill "lblHdr", 0, 0, 720, 30, ClrOrange()
    Lbl "lblHdrTitle", 14, 6, 500, 18, "IN ẤN DỰ TOÁN   G8 / ESCON", 12, True, ClrOnOrange()
    Lbl "lblHdrVer", 520, 9, 186, 14, "V6.1  •  1 hoặc nhiều hạng mục", 8.5, False, ClrOnOrange(), 3

    ' ---- nguon ----
    Opt "optFile", 12, 40, 90, 18, "Một file", "src"
    Cbo "cboFile", 104, 39, 420, 19
    Btn "btnOpenFile", 530, 39, 84, 20, "Mở file...", ClrBtnGray(), 9
    Btn "btnRefresh", 620, 39, 88, 20, "Làm mới", ClrBtnGray(), 9
    Opt "optFolder", 12, 64, 90, 18, "Cả thư mục", "src"
    Txt "txtFolder", 104, 63, 420, 19, True
    Btn "btnFolder", 530, 63, 84, 20, "Chọn...", ClrBtnGray(), 9
    Chk "chkSub", 620, 64, 96, 18, "Gồm thư mục con"
    Lbl "lblDetect", 12, 89, 696, 16, "", 9.5, True, ClrTextDim()
    Fill "lnSrc", 12, 110, 696, 1, ClrLine()

    ' ---- panel 1 hang muc ----
    Fill "pnl1", 12, 116, 300, 260, ClrPanel()
    Lbl "lblP1", 20, 121, 200, 16, "1 HẠNG MỤC  —  8 sheet", 10, True, ClrOrange()
    Btn "btnAll", 20, 141, 60, 18, "Tất cả", ClrBtnGray(), 8.5
    Btn "btnNone", 84, 141, 60, 18, "Bỏ chọn", ClrBtnGray(), 8.5
    Btn "btn5", 148, 141, 88, 18, "5 sheet chính", ClrBtnGray(), 8.5
    Lbl "lblC1In", 22, 164, 20, 12, "In", 8, True, ClrTextDim()
    Lbl "lblC1TT", 42, 164, 26, 12, "TT", 8, True, ClrTextDim(), 2
    Lbl "lblC1Sheet", 76, 164, 200, 12, "SHEET", 8, True, ClrTextDim()
    y = 178
    For i = 1 To 8
        If i Mod 2 = 0 Then Fill "row1_" & i, 16, y - 1, 292, 20, ClrRowAlt()
        Chk "chk1_" & i, 24, y + 1, 16, 16, ""
        TT "txt1_" & i, 42, y, 26, 17
        Lbl "lbl1_" & i, 76, y + 2, 228, 15, "", 9, False, ClrText()
        y = y + 20
    Next i

    ' ---- panel nhieu hang muc: luoi hang muc x sheet ----
    Dim k As Integer, x0 As Double: x0 = 520          ' cot sheet dau tien
    Fill "pnl2", 324, 116, 384, 260, ClrPanel()
    Lbl "lblP2", 332, 121, 370, 16, "NHIỀU HẠNG MỤC  —  tick sheet cho từng hạng mục", 10, True, ClrGreen()
    Lbl "lblC2HM", 332, 142, 180, 12, "HẠNG MỤC", 8, True, ClrTextDim()
    For k = 1 To 6
        Lbl "lblC2_" & k, x0 + (k - 1) * 29 - 2, 142, 30, 12, Choose(k, "THKP", "GTH", "ĐGCT", "CT", "GT", "TB"), 7.5, True, ClrTextDim(), 2
        Me_Tip "lblC2_" & k, Choose(k, "THKP hạng mục", "Giá tổng hợp", "Đơn giá chi tiết", "Công trình", "Giá tháng (chung)", "TH chi phí TB (chung)")
    Next k
    Lbl "lblTTN", 332, 160, 180, 14, "Số thứ tự trong PDF", 8.5, False, ClrTextDim()
    For k = 1 To 6
        TT "txtN_" & k, x0 + (k - 1) * 29, 158, 26, 17
    Next k
    Fill "rowM", 328, 179, 376, 20, ClrRowAlt()
    Lbl "lblMaster", 332, 182, 186, 15, "Tất cả hạng mục", 9, True, ClrText()
    For k = 1 To 6
        Chk "chkM_" & k, x0 + (k - 1) * 29 + 6, 181, 16, 16, ""
    Next k

    Dim fr As Object
    Set fr = Add("Forms.Frame.1", "fraHM", 328, 203, 376, 122)
    fr.Caption = "": fr.BackColor = ClrPanel(): fr.BorderStyle = 0: fr.SpecialEffect = 0
    fr.ScrollBars = 2: fr.KeepScrollBarsVisible = 0          ' cuon doc khi nhieu hang muc
    fr.ScrollHeight = 122
    Dim r As Integer, c As Object
    For r = 1 To 20
        y = (r - 1) * 20
        If r Mod 2 = 0 Then
            Set c = fr.Controls.Add("Forms.Label.1", "rowH_" & r, True)
            c.Left = 0: c.Top = y: c.Width = 376: c.Height = 20
            c.Caption = "": c.BackStyle = 1: c.BackColor = ClrRowAlt(): c.BorderStyle = 0
        End If
        Set c = fr.Controls.Add("Forms.Label.1", "lblH_" & r, True)
        c.Left = 4: c.Top = y + 3: c.Width = 184: c.Height = 15
        c.Caption = "": c.BackStyle = 0: c.WordWrap = False
        c.Font.Name = "Segoe UI": c.Font.Size = 8.5: c.ForeColor = ClrText()
        For k = 1 To 6
            Set c = fr.Controls.Add("Forms.CheckBox.1", "chkH_" & r & "_" & k, True)
            c.Left = x0 - 328 + (k - 1) * 29 + 6: c.Top = y + 2: c.Width = 16: c.Height = 16
            c.Caption = "": c.BackStyle = 0
        Next k
    Next r
    Set c = fr.Controls.Add("Forms.Label.1", "lblFolderNote", True)
    c.Left = 4: c.Top = 6: c.Width = 360: c.Height = 60
    c.Caption = U("Chế độ cả thư mục: dòng 'Tất cả hạng mục' là mẫu cho mọi file." & vbLf & _
                  "Giá tháng in kèm hạng mục CUỐI, TH chi phí TB in kèm hạng mục ĐẦU.")
    c.BackStyle = 0: c.WordWrap = True
    c.Font.Name = "Segoe UI": c.Font.Size = 8.5: c.ForeColor = ClrTextDim()

    Lbl "lblXuat", 332, 332, 80, 14, "Xuất PDF:", 8.5, False, ClrTextDim()
    Opt "optTach", 400, 330, 150, 17, "Mỗi hạng mục 1 file", "xk"
    Opt "optGhep", 556, 330, 140, 17, "Ghép 1 file", "xk"
    Chk "chkTenHM", 400, 350, 300, 17, "Đặt tên PDF theo tên hạng mục"

    ' ---- tuy chon chung ----
    Chk "chkTong", 12, 383, 340, 18, "In cột Tổng cộng (cột L) ở sheet Giá tổng hợp"
    Chk "chkOpen", 360, 383, 220, 18, "Mở thư mục PDF khi xong"
    Lbl "lblOut", 12, 407, 80, 14, "Lưu PDF vào:", 9, False, ClrTextDim()
    Txt "txtOut", 92, 405, 432, 19, False
    Btn "btnOut", 530, 405, 84, 20, "Chọn...", ClrBtnGray(), 9
    Btn "btnOutClr", 620, 405, 88, 20, "Cạnh file gốc", ClrBtnGray(), 9
    Fill "lnOpt", 12, 432, 696, 1, ClrLine()

    ' ---- nut thao tac ----
    Btn "btnAlign", 12, 439, 110, 32, "Căn chỉnh", ClrBtnGray(), 10
    Btn "btnPreview", 128, 439, 110, 32, "Xem trước", ClrBtnGray(), 10
    Btn "btnPrint", 244, 439, 110, 32, "In máy in", ClrBtnGray(), 10
    Btn "btnPDF", 360, 439, 242, 32, "XUẤT PDF", ClrOrange(), 11, ClrOnOrange()
    Btn "btnClose", 608, 439, 100, 32, "Đóng", ClrBtnCancel(), 10

    ' ---- nhat ky ----
    Set c = Add("Forms.TextBox.1", "txtLog", 12, 478, 696, 82)
    c.MultiLine = True: c.WordWrap = True: c.ScrollBars = 2   ' fmScrollBarsVertical
    c.Locked = True
    c.Font.Name = "Segoe UI": c.Font.Size = 8.5
    c.SpecialEffect = 0: c.BorderStyle = 1: c.BorderColor = ClrLine()
    c.BackColor = ClrInputBg(): c.ForeColor = ClrTextDim()
End Sub

' ============================================================================
'  TIEN ICH TAO CONTROL
' ============================================================================
Private Function Add(progId As String, nm As String, L As Double, T As Double, W As Double, H As Double) As Object
    Dim c As Object
    Set c = mComp.Designer.Controls.Add(progId, nm, True)
    c.Left = L: c.Top = T: c.Width = W: c.Height = H
    Set Add = c
End Function

Private Sub Lbl(nm As String, L As Double, T As Double, W As Double, H As Double, cap As String, _
        Optional sz As Double = 9, Optional bold As Boolean = False, Optional fg As Long = -1, _
        Optional align As Long = 1)
    Dim c As Object: Set c = Add("Forms.Label.1", nm, L, T, W, H)
    c.Caption = U(cap)
    c.Font.Name = "Segoe UI": c.Font.Size = sz: c.Font.Bold = bold
    c.ForeColor = IIf(fg = -1, ClrText(), fg)
    c.TextAlign = align
    c.BackStyle = 0
    c.WordWrap = False
End Sub

Private Sub Me_Tip(nm As String, tip As String)
    mComp.Designer.Controls(nm).ControlTipText = U(tip)
End Sub

Private Sub Fill(nm As String, L As Double, T As Double, W As Double, H As Double, bg As Long)
    Dim c As Object: Set c = Add("Forms.Label.1", nm, L, T, W, H)
    c.Caption = "": c.BackStyle = 1: c.BackColor = bg: c.BorderStyle = 0
End Sub

Private Sub Btn(nm As String, L As Double, T As Double, W As Double, H As Double, cap As String, _
        bg As Long, Optional sz As Double = 9.5, Optional fg As Long = -1)
    ' Nut = 2 Label: nen <nm> (doi mau hover) + chu <nm>_t can giua theo chieu doc
    ' (Label khong tu can giua doc). Ca 2 cung bat Click qua clsInAnBtn.
    Dim c As Object: Set c = Add("Forms.Label.1", nm, L, T, W, H)
    c.Caption = "": c.BackStyle = 1: c.BackColor = bg: c.BorderStyle = 0
    c.Tag = "1"
    Dim th As Double: th = sz * 1.36
    Set c = Add("Forms.Label.1", nm & "_t", L, T + (H - th) / 2, W, th + 1)
    c.Caption = U(cap)
    c.Font.Name = "Segoe UI": c.Font.Size = sz: c.Font.Bold = True
    c.ForeColor = IIf(fg = -1, ClrText(), fg)
    c.TextAlign = 2
    c.BackStyle = 0: c.WordWrap = False
End Sub

Private Sub Chk(nm As String, L As Double, T As Double, W As Double, H As Double, cap As String)
    Dim c As Object: Set c = Add("Forms.CheckBox.1", nm, L, T, W, H)
    c.Caption = U(cap)
    c.BackStyle = 0
    c.ForeColor = ClrText()
    c.Font.Name = "Segoe UI": c.Font.Size = 9
End Sub

Private Sub Opt(nm As String, L As Double, T As Double, W As Double, H As Double, cap As String, grp As String)
    Dim c As Object: Set c = Add("Forms.OptionButton.1", nm, L, T, W, H)
    c.Caption = U(cap)
    c.GroupName = grp
    c.BackStyle = 0
    c.ForeColor = ClrText()
    c.Font.Name = "Segoe UI": c.Font.Size = 9
End Sub

Private Sub Txt(nm As String, L As Double, T As Double, W As Double, H As Double, locked As Boolean)
    Dim c As Object: Set c = Add("Forms.TextBox.1", nm, L, T, W, H)
    c.Font.Name = "Segoe UI": c.Font.Size = 9
    c.SpecialEffect = 0: c.BorderStyle = 1: c.BorderColor = ClrLine()
    c.BackColor = ClrInputBg(): c.ForeColor = ClrText()
    c.Locked = locked
End Sub

Private Sub TT(nm As String, L As Double, T As Double, W As Double, H As Double)
    Dim c As Object: Set c = Add("Forms.TextBox.1", nm, L, T, W, H)
    c.Font.Name = "Segoe UI": c.Font.Size = 9: c.Font.Bold = True
    c.TextAlign = 2: c.MaxLength = 2
    c.SpecialEffect = 0: c.BorderStyle = 1: c.BorderColor = ClrLine()
    c.BackColor = ClrInputBg(): c.ForeColor = RGB(244, 224, 190)
End Sub

Private Sub Cbo(nm As String, L As Double, T As Double, W As Double, H As Double)
    Dim c As Object: Set c = Add("Forms.ComboBox.1", nm, L, T, W, H)
    c.Style = 2                        ' fmStyleDropDownList
    c.Font.Name = "Segoe UI": c.Font.Size = 9
    c.SpecialEffect = 0: c.BorderStyle = 1: c.BorderColor = ClrLine()
    c.BackColor = ClrInputBg(): c.ForeColor = ClrText()
    c.ListRows = 12
End Sub
