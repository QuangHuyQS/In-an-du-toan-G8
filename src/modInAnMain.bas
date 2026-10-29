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
Private Const APP_KEY As String = "ToolInAnG8"

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
        MsgBox "Thieu form frmInDuToan trong file nay. Chay lai BUILD_V6.ps1.", vbExclamation
        Exit Sub
    End If
    gFrm.Show vbModeless
End Sub

Public Sub MoToolInAn()
    MoBangDieuKhien
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
Public Sub XuatTheoFileTong()
    Dim fd As Object, sumPath As String, outF As String, defOut As String
    Dim n As Long, rep As String, ans As VbMsgBoxResult
    sumPath = CfgGet("sumfile", "")
    Set fd = Application.FileDialog(3)          ' msoFileDialogFilePicker
    fd.Title = "Chọn FILE TỔNG HỢP (chứa công thức link tới các file dự toán)"
    fd.AllowMultiSelect = False
    fd.Filters.Clear
    fd.Filters.Add "Excel", "*.xls;*.xlsx;*.xlsm"
    If Len(sumPath) > 0 Then fd.InitialFileName = sumPath
    If fd.Show <> -1 Then Exit Sub
    sumPath = fd.SelectedItems(1)
    CfgSet "sumfile", sumPath

    defOut = Left$(sumPath, InStrRev(sumPath, "\")) & "PDF"
    ans = MsgBox("Lưu PDF vào:" & vbCrLf & defOut & vbCrLf & vbCrLf & _
                 "Yes = lưu vào đây   |   No = chọn thư mục khác   |   Cancel = huỷ", _
                 vbYesNoCancel + vbQuestion, "Xuất PDF theo file tổng hợp")
    If ans = vbCancel Then Exit Sub
    If ans = vbYes Then
        outF = defOut
    Else
        Set fd = Application.FileDialog(4)      ' msoFileDialogFolderPicker
        fd.Title = "Chọn thư mục lưu PDF"
        If fd.Show <> -1 Then Exit Sub
        outF = fd.SelectedItems(1)
    End If

    Dim fromN As Long
    fromN = CLng(Val(InputBox("Chạy từ mục số (số NN đầu tên file PDF).  1 = chạy từ đầu." & vbCrLf & _
                              "Dùng để chạy tiếp khi lần trước bị ngắt giữa chừng.", _
                              "Xuất PDF theo file tổng hợp", "1")))
    If fromN < 1 Then fromN = 1
    Application.StatusBar = "Đang xuất PDF theo file tổng hợp ..."
    n = API_RunTong(sumPath, outF, False, rep, fromN)       ' False = KHONG in cot Tong cong o Gia tong hop
    Application.StatusBar = False
    If n > 0 Then
        MsgBox "Đã xuất " & n & " PDF vào:" & vbCrLf & outF & vbCrLf & vbCrLf & _
               "Thứ tự + nguồn từng mục: _THU_TU_XUAT.txt" & _
               IIf(InStr(rep, "KHONG XUAT") > 0, vbCrLf & "CÓ mục không xuất được - mở file đó, tìm 'KHONG XUAT'.", ""), _
               vbInformation, "Xuất PDF theo file tổng hợp"
        On Error Resume Next
        CreateObject("Shell.Application").Open CVar(outF)
        On Error GoTo 0
    Else
        MsgBox "Không xuất được PDF nào." & vbCrLf & Left$(rep, 800), vbExclamation, "Xuất PDF theo file tổng hợp"
    End If
End Sub