Attribute VB_Name = "modInAnMain"
Option Explicit

' ============================================================================
'  modInAnMain  -  diem vao cua TOOL IN AN DU TOAN G8 / ESCON (V6)
'
'  MoBangDieuKhien : mo form frmInDuToan (gan nut tren sheet vao macro nay)
'  MoToolInAn      : ten cu, giu de nut/phim tat cu van chay
'
'  Form lam viec qua cac ham API_... trong module inanDTG8_V6.
'  Chu tieng Viet co dau tren form duoc giai ma bang U("...\uXXXX...")
'  vi ma nguon VBA chi luu duoc ky tu ANSI.
' ============================================================================

Public gFrm As Object               ' form dang mo (modeless); form tu dat Nothing khi dong
Public gTuBat As Boolean            ' form mo tu MO_FORM_IN_AN.bat (file tool an)
Private Const APP_KEY As String = "ToolInAnG8"

' MsgBox cua VBA chi hien chu ANSI -> tieng Viet co dau thanh "?". Goi thang MessageBoxW (Unicode).
#If VBA7 Then
Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hWnd As LongPtr, ByVal lpText As LongPtr, _
        ByVal lpCaption As LongPtr, ByVal uType As Long) As Long
#Else
Private Declare Function MessageBoxW Lib "user32" (ByVal hWnd As Long, ByVal lpText As Long, _
        ByVal lpCaption As Long, ByVal uType As Long) As Long
#End If

' Thay MsgBox: hien dung tieng Viet. Tra ve vbOK / vbYes / vbNo / vbCancel nhu MsgBox.
Public Function MsgBoxU(ByVal prompt As String, Optional ByVal buttons As Long = vbOKOnly, _
        Optional ByVal title As String = "") As VbMsgBoxResult
    If Len(title) = 0 Then title = "In ấn dự toán G8"
    MsgBoxU = MessageBoxW(Application.hWnd, StrPtr(prompt), StrPtr(title), buttons Or &H10000)   ' MB_SETFOREGROUND
End Function

' Form MODELESS: mo xong van lam viec voi cac cua so Excel khac.
' Bam lai nut khi form dang mo -> dua form len tren, khong mo them.
Public Sub MoBangDieuKhien()
    If Not gFrm Is Nothing Then
        On Error Resume Next
        gFrm.Show vbModeless
        If Err.Number = 0 Then Exit Sub
        Err.Clear
        On Error GoTo 0
        Set gFrm = Nothing
    End If
    On Error Resume Next
    Set gFrm = VBA.UserForms.Add("frmInDuToan")   ' late bind: module van bien dich khi chua co form
    On Error GoTo 0
    If gFrm Is Nothing Then
        MsgBoxU "Thiếu form frmInDuToan trong file này. Chạy lại BUILD_V6.ps1.", vbExclamation
        Exit Sub
    End If
    gFrm.Show vbModeless
End Sub

Public Sub MoToolInAn()
    MoBangDieuKhien
End Sub

' Goi tu MO_FORM_IN_AN.bat: an cua so file tool, chi hien form in.
' Dong form -> tu dong file tool (va thoat Excel neu khong con file nao khac).
Public Sub MoFormTuBat()
    gTuBat = True
    On Error Resume Next
    ThisWorkbook.Windows(1).Visible = False
    On Error GoTo 0
    MoBangDieuKhien
End Sub

Public Sub DongToolTuBat()
    If Not gFrm Is Nothing Then Exit Sub
    gTuBat = False
    Dim wb As Workbook, nVis As Integer, w As Window
    For Each wb In Application.Workbooks
        If Not wb Is ThisWorkbook Then
            For Each w In wb.Windows
                If w.Visible Then nVis = nVis + 1
            Next w
        End If
    Next wb
    If nVis = 0 Then
        ThisWorkbook.Saved = True          ' file tool mo chi doc, cai dat luu trong registry
        Application.Quit
    Else
        ThisWorkbook.Close False
    End If
End Sub

Public Sub MoToolInAn_FileTong()
    MoBangDieuKhien
End Sub

' Giai ma chuoi co \uXXXX thanh Unicode
Public Function U(ByVal s As String) As String
    Dim p As Long
    p = InStr(1, s, "\u")
    Do While p > 0
        s = Left$(s, p - 1) & ChrW$(CLng("&H" & Mid$(s, p + 2, 4))) & Mid$(s, p + 6)
        p = InStr(p + 1, s, "\u")
    Loop
    U = s
End Function

' ---- Luu / doc cau hinh form (HKCU\Software\VB and VBA Program Settings) ----
Public Function CfgGet(key As String, def As String) As String
    On Error Resume Next
    CfgGet = GetSetting(APP_KEY, "Form", key, def)
    If Err.Number <> 0 Then CfgGet = def
    On Error GoTo 0
End Function

Public Sub CfgSet(key As String, v As String)
    On Error Resume Next
    SaveSetting APP_KEY, "Form", key, v
    On Error GoTo 0
End Sub

' ---- Bang mau giao dien ----
Public Function ClrBg() As Long:            ClrBg = RGB(26, 33, 45):           End Function
Public Function ClrPanel() As Long:         ClrPanel = RGB(33, 42, 56):        End Function
Public Function ClrRowAlt() As Long:        ClrRowAlt = RGB(38, 48, 64):       End Function
Public Function ClrLine() As Long:          ClrLine = RGB(58, 72, 90):         End Function
Public Function ClrText() As Long:          ClrText = RGB(226, 232, 240):      End Function
Public Function ClrTextDim() As Long:       ClrTextDim = RGB(148, 163, 184):   End Function
Public Function ClrTextOff() As Long:       ClrTextOff = RGB(92, 106, 124):    End Function
Public Function ClrInputBg() As Long:       ClrInputBg = RGB(20, 26, 36):      End Function
Public Function ClrOrange() As Long:        ClrOrange = RGB(226, 140, 43):     End Function
Public Function ClrOrangeHover() As Long:   ClrOrangeHover = RGB(242, 160, 68): End Function
Public Function ClrOnOrange() As Long:      ClrOnOrange = RGB(38, 26, 10):     End Function
Public Function ClrGreen() As Long:         ClrGreen = RGB(74, 200, 130):      End Function
Public Function ClrRed() As Long:           ClrRed = RGB(240, 110, 110):       End Function
Public Function ClrBtnGray() As Long:       ClrBtnGray = RGB(46, 58, 74):      End Function
Public Function ClrBtnGrayHover() As Long:  ClrBtnGrayHover = RGB(62, 78, 98): End Function
Public Function ClrBtnOff() As Long:        ClrBtnOff = RGB(36, 44, 56):       End Function
Public Function ClrBtnCancel() As Long:     ClrBtnCancel = RGB(70, 44, 44):    End Function
Public Function ClrBtnCancelHover() As Long: ClrBtnCancelHover = RGB(92, 58, 58): End Function

' ============================================================================
'  XuatTheoFileTong : xuat PDF theo FILE TONG HOP (gan vao nut xanh tren sheet)
'  Chon file tong -> di theo link toi tung file du toan + hang muc -> moi hang
'  muc 1 PDF dat ten <NN>_<STT>_<ten>.pdf theo dung thu tu file tong.
'  Engine: inanDTG8_V6.API_RunTong.
' ============================================================================
' Nut xanh tren sheet chinh: chon file tong -> sang sheet cau hinh de chon sheet / ten du an
Public Sub XuatTheoFileTong()
    Dim ws As Worksheet: Set ws = TongCfgSheet()
    If ws Is Nothing Then Exit Sub
    If Not XuatTong_PickFile(ws) Then Exit Sub
    ws.Activate
    ws.Range("C8").Select
    MsgBoxU "Đã chọn file tổng hợp:" & vbCrLf & ws.Range("C4").Value & vbCrLf & vbCrLf & _
           "Chọn sheet cần in (cột thứ tự cho file 1 hạng mục / nhiều hạng mục)," & vbCrLf & _
           "điền Tên dự án nếu muốn thay, rồi bấm  ▶ BẮT ĐẦU XUẤT.", vbInformation, "Xuất PDF theo file tổng hợp"
End Sub

' Goi tu form (nut "In theo file tong"): hien cua so file tool (neu dang an do MO_FORM_IN_AN.bat), sang sheet cau hinh
Public Sub XuatTong_TuForm()
    gTuBat = False                      ' dang dung sheet cau hinh -> dong form khong tu dong file tool
    On Error Resume Next
    ThisWorkbook.Windows(1).Visible = True
    Application.Visible = True
    ThisWorkbook.Activate
    On Error GoTo 0
    XuatTheoFileTong
End Sub

' Nut "Chon file tong hop..." tren sheet cau hinh
Public Sub XuatTong_ChonFile()
    Dim ws As Worksheet: Set ws = TongCfgSheet()
    If ws Is Nothing Then Exit Sub
    XuatTong_PickFile ws
End Sub

Private Function TongCfgSheet() As Worksheet
    On Error Resume Next
    Set TongCfgSheet = ThisWorkbook.Worksheets("IN THEO FILE TONG")
    On Error GoTo 0
    If TongCfgSheet Is Nothing Then
        MsgBoxU "Thiếu sheet 'IN THEO FILE TONG'. Chạy lại BUILD_V6.ps1 để dựng lại tool.", vbExclamation
    End If
End Function

Private Function XuatTong_PickFile(ws As Worksheet) As Boolean
    Dim fd As Object, sumPath As String
    sumPath = Trim$(CStr(ws.Range("C4").Value))
    If Len(sumPath) = 0 Then sumPath = CfgGet("sumfile", "")
    Set fd = Application.FileDialog(3)          ' msoFileDialogFilePicker
    fd.Title = "Chọn FILE TỔNG HỢP (chứa công thức link tới các file dự toán)"
    fd.AllowMultiSelect = False
    fd.Filters.Clear
    fd.Filters.Add "Excel", "*.xls;*.xlsx;*.xlsm"
    If Len(sumPath) > 0 Then fd.InitialFileName = sumPath
    If fd.Show <> -1 Then Exit Function
    sumPath = fd.SelectedItems(1)
    CfgSet "sumfile", sumPath
    ws.Range("C4").Value = sumPath
    XuatTong_PickFile = True
End Function

' Doc cot thu tu (C = 1 hang muc, D = nhieu hang muc) dong 16..23 -> "8,1,2,3,5"
'   dong : 16 TB, 17 THKP, 18 Gia TH, 19 DGCT, 20 Cong trinh, 21 Gia thang, 22 Tong hop VT, 23 TH chi phi XD
Private Function TongSelList(ws As Worksheet, col As Long, isMulti As Boolean) As String
    Dim ids As Variant, ord(1 To 8) As Double, idx(1 To 8) As Integer, n As Integer
    Dim i As Integer, j As Integer, v As Variant, tD As Double, tI As Integer, s As String
    If isMulti Then
        ids = Array(6, 1, 2, 3, 4, 5, 0, 0)
    Else
        ids = Array(8, 1, 2, 3, 4, 5, 6, 7)
    End If
    For i = 0 To 7
        v = ws.Cells(16 + i, col).Value
        If ids(i) > 0 And IsNumeric(v) And Not IsEmpty(v) And CStr(v) <> "" Then
            If CDbl(v) > 0 Then n = n + 1: ord(n) = CDbl(v): idx(n) = CInt(ids(i))
        End If
    Next i
    For i = 2 To n                                   ' sap xep theo so thu tu
        For j = i To 2 Step -1
            If ord(j) < ord(j - 1) Then
                tD = ord(j): ord(j) = ord(j - 1): ord(j - 1) = tD
                tI = idx(j): idx(j) = idx(j - 1): idx(j - 1) = tI
            End If
        Next j
    Next i
    For i = 1 To n: s = s & IIf(i > 1, ",", "") & idx(i): Next i
    If n = 0 Then s = "0"                            ' khong chon gi -> danh dau (khong dung mac dinh)
    TongSelList = s
End Function

' Nut "BAT DAU XUAT" tren sheet cau hinh
Public Sub XuatTong_BatDau()
    Dim ws As Worksheet: Set ws = TongCfgSheet()
    If ws Is Nothing Then Exit Sub
    Dim sumPath As String, outF As String, n As Long, rep As String
    sumPath = Trim$(CStr(ws.Range("C4").Value))
    If Len(sumPath) = 0 Or Not CreateObject("Scripting.FileSystemObject").FileExists(sumPath) Then
        If Not XuatTong_PickFile(ws) Then Exit Sub
        sumPath = Trim$(CStr(ws.Range("C4").Value))
    End If
    outF = Trim$(CStr(ws.Range("C5").Value))
    If Len(outF) = 0 Then outF = Left$(sumPath, InStrRev(sumPath, "\")) & "PDF"

    Dim fromN As Long, toN As Long, dry As Boolean, inclTC As Boolean
    fromN = CLng(Val(CStr(ws.Range("C6").Value))): If fromN < 1 Then fromN = 1
    toN = CLng(Val(CStr(ws.Range("C7").Value))): If toN < 0 Then toN = 0
    inclTC = (LCase$(Left$(Trim$(CStr(ws.Range("C10").Value)), 1)) = "c")
    dry = (LCase$(Left$(Trim$(CStr(ws.Range("C11").Value)), 1)) = "c")

    Dim s1 As String, sN As String
    s1 = TongSelList(ws, 3, False)
    sN = TongSelList(ws, 4, True)
    If s1 = "0" And sN = "0" Then
        MsgBoxU "Chưa chọn sheet nào để in (cột thứ tự đang trống).", vbExclamation, "Xuất PDF theo file tổng hợp"
        Exit Sub
    End If

    API_SetTongSheets s1, sN
    API_SetTongTen CStr(ws.Range("C8").Value), CStr(ws.Range("C9").Value)
    API_SetTongDryRun dry
    API_SetTongNoDGCT False
    API_SetTongFilter ""
    API_SetTongSheetName CStr(ws.Range("C14").Value)
    API_SetTongRowSel RowSelString()
    API_SetTongUpper (LCase$(Left$(Trim$(CStr(ws.Range("C12").Value)), 1)) <> "k")
    API_SetGiaThangAnMaSo (LCase$(Left$(Trim$(CStr(ws.Range("C13").Value)), 1)) <> "k")

    Application.StatusBar = "Đang xuất PDF theo file tổng hợp ..."
    n = API_RunTong(sumPath, outF, inclTC, rep, fromN, toN)
    Application.StatusBar = False
    API_SetTongDryRun False
    If dry Then BuildRowSelSheet ws
    If dry Then
        MsgBoxU "Chạy khô xong (không xuất PDF). Danh sách + nguồn từng mục:" & vbCrLf & _
               outF & "\_THU_TU_XUAT.txt" & vbCrLf & vbCrLf & _
               "Muốn chọn sheet in riêng cho từng dòng: mở sheet CHON SHEET THEO MUC" & vbCrLf & _
               "(gõ x vào sheet cần in, ô C2 = Có), rồi chạy với 'Chỉ dò link' = Không.", vbInformation, "Xuất PDF theo file tổng hợp"
        On Error Resume Next
        CreateObject("Shell.Application").Open CVar(outF & "\_THU_TU_XUAT.txt")
        On Error GoTo 0
        Exit Sub
    End If
    If n > 0 Then
        MsgBoxU "Đã xuất " & n & " PDF vào:" & vbCrLf & outF & vbCrLf & vbCrLf & _
               "Thứ tự + nguồn từng mục: _THU_TU_XUAT.txt" & _
               IIf(InStr(rep, "KHONG XUAT") > 0, vbCrLf & "CÓ mục không xuất được - mở file đó, tìm 'KHONG XUAT'.", ""), _
               vbInformation, "Xuất PDF theo file tổng hợp"
        On Error Resume Next
        CreateObject("Shell.Application").Open CVar(outF)
        On Error GoTo 0
    Else
        MsgBoxU "Không xuất được PDF nào." & vbCrLf & Left$(rep, 800), vbExclamation, "Xuất PDF theo file tổng hợp"
    End If
End Sub

' ============================================================================
'  Bang chon sheet in rieng theo tung dong cua file tong (sheet "CHON SHEET THEO MUC")
'  Sinh sau moi lan CHAY KHO. Cot G..L = ma sheet 1..6 (THKP, Gia TH, DGCT, Cong trinh, Gia thang, TH chi phi TB)
' ============================================================================
Private Function RowSelName() As String: RowSelName = "CHON SHEET THEO MUC": End Function

Private Sub BuildRowSelSheet(cfg As Worksheet)
    Dim v As Variant: v = API_TongItems()
    If IsEmpty(v) Then Exit Sub
    Dim ws As Worksheet, i As Long, k As Long, r As Long, n As Long, ord As Variant, def(1 To 6) As Boolean
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(RowSelName())
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=cfg)
        ws.Name = RowSelName()
    End If
    Dim keepApply As String: keepApply = "Không"
    If Len(CStr(ws.Range("C2").Value)) > 0 Then keepApply = CStr(ws.Range("C2").Value)
    ws.Cells.Clear
    ws.Cells.Font.Name = "Arial": ws.Cells.Font.Size = 10
    ' mac dinh = cau hinh cot D cua sheet IN THEO FILE TONG: dong 16 TB(6), 17 THKP(1), 18 GTH(2), 19 DGCT(3), 20 CT(4), 21 GT(5)
    ord = Array(6, 1, 2, 3, 4, 5)
    For i = 0 To 5
        If IsNumeric(cfg.Cells(16 + i, 4).Value) And Not IsEmpty(cfg.Cells(16 + i, 4).Value) Then
            If CDbl(cfg.Cells(16 + i, 4).Value) > 0 Then def(ord(i)) = True
        End If
    Next i
    ws.Range("B1").Value = "CHỌN SHEET IN RIÊNG THEO TỪNG DÒNG CỦA BẢNG TỔNG HỢP"
    ws.Range("B1").Font.Size = 13: ws.Range("B1").Font.Bold = True
    ws.Range("B2").Value = "Áp dụng bảng này khi xuất PDF (Có / Không):"
    ws.Range("B2").Font.Bold = True
    ws.Range("C2").Value = keepApply
    With ws.Range("C2").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Có,Không"
    End With
    ws.Range("C2").Interior.Color = RGB(255, 242, 204)
    ws.Range("B3").Value = "Gõ x vào ô sheet cần in cho dòng đó (xóa x = không in). Dòng xám = in nguyên sheet / không có file dự toán nên không chọn được. Bảng được dựng lại mỗi lần Chạy khô."
    Dim hdr As Variant
    hdr = Array("TT", "Nhãn", "Hạng mục trong file tổng", "Giá trị", "File nguồn", "Vị trí", "THKP", "Giá tổng hợp", "Đơn giá chi tiết", "Công trình", "Giá tháng", "TH chi phí TB", "")
    For k = 0 To 11: ws.Cells(5, 1 + k).Value = hdr(k): Next k
    With ws.Range("A5:L5")
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255): .Interior.Color = RGB(51, 65, 85)
        .WrapText = True: .VerticalAlignment = xlCenter: .HorizontalAlignment = xlCenter
    End With
    n = UBound(v, 1)
    For i = 1 To n
        r = 5 + i
        ws.Cells(r, 1).Value = CLng(Val(CStr(v(i, 1))))
        ws.Cells(r, 2).Value = "'" & CStr(v(i, 2))
        ws.Cells(r, 3).Value = CStr(v(i, 3))
        ws.Cells(r, 4).Value = "'" & CStr(v(i, 4))
        ws.Cells(r, 5).Value = CStr(v(i, 5))
        ws.Cells(r, 6).Value = CStr(v(i, 6))
        ws.Cells(r, 13).Value = CStr(v(i, 7))
        If CStr(v(i, 7)) = "1" Then
            For k = 1 To 6
                If def(k) Then ws.Cells(r, 6 + k).Value = "x"
            Next k
            ws.Range(ws.Cells(r, 7), ws.Cells(r, 12)).Interior.Color = RGB(255, 255, 255)
            ws.Range(ws.Cells(r, 7), ws.Cells(r, 12)).Locked = False
        Else
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 12)).Interior.Color = RGB(226, 232, 240)
        End If
    Next i
    ws.Columns("M").Hidden = True
    With ws.Range(ws.Cells(5, 1), ws.Cells(5 + n, 12))
        .Borders.Color = RGB(203, 213, 225)
        .VerticalAlignment = xlCenter
    End With
    ws.Range(ws.Cells(6, 7), ws.Cells(5 + n, 12)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(6, 4), ws.Cells(5 + n, 4)).HorizontalAlignment = xlRight
    ws.Columns("A").ColumnWidth = 5: ws.Columns("B").ColumnWidth = 14: ws.Columns("C").ColumnWidth = 46
    ws.Columns("D").ColumnWidth = 16: ws.Columns("E").ColumnWidth = 34: ws.Columns("F").ColumnWidth = 16
    ws.Columns("G:L").ColumnWidth = 11
    ws.Rows(5).RowHeight = 32
End Sub

' "seq=ids;seq=ids" tu sheet CHON SHEET THEO MUC khi C2 = Co ; rong neu khong ap dung
Private Function RowSelString() As String
    Dim ws As Worksheet, r As Long, last As Long, k As Long, ids As String, s As String
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(RowSelName())
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    If LCase$(Left$(Trim$(CStr(ws.Range("C2").Value)), 1)) <> "c" Then Exit Function
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 6 To last
        If CStr(ws.Cells(r, 13).Value) = "1" And IsNumeric(ws.Cells(r, 1).Value) Then
            ids = ""
            For k = 1 To 6
                If LCase$(Trim$(CStr(ws.Cells(r, 6 + k).Value))) = "x" Then ids = ids & IIf(Len(ids) > 0, ",", "") & k
            Next k
            s = s & IIf(Len(s) > 0, ";", "") & CLng(ws.Cells(r, 1).Value) & "=" & ids
        End If
    Next r
    RowSelString = s
End Function
