Attribute VB_Name = "inanDTG8_V6"
Option Explicit

' ================================================================
' TOOL IN AN DU TOAN G8/ESCON - V6 (ENGINE, KHONG CO GIAO DIEN)
' ================================================================
' Module nay CHI chua logic in. Giao dien = frmInDuToan (mo bang
' modInAnMain.MoBangDieuKhien). Form nap cau hinh vao cac bien Public
' duoi day roi goi cac ham API_... o cuoi module:
'
'   API_Detect       - nhan dien file 1 hay 2 hang muc + sheet nao co
'   API_Run1HM       - can chinh / PDF / in / xem truoc file 1 hang muc
'   API_RunNHM       - xuat PDF file nhieu hang muc (tach tung HM / ghep 1)
'   API_ProcessFile  - xu ly 1 file tren dia (in ca thu muc)
'   API_CollectFiles - quet file Excel trong thu muc
'
' Ten PDF:
'   1 hang muc, 1 file : <ten excel>_FULL.pdf
'   1 hang muc, thu muc: <ten excel>.pdf
'   nhieu hang muc     : <ten excel>_HM1.pdf .. _HMn.pdf (hoac -<ten HM>.pdf; ghep: <ten excel>.pdf)
'   Thu muc luu = gOutFolder1 / gHM_OutFolder (rong = canh file goc).
'
' GHI CHU KIEN TRUC:
'   - Nhanh 1HM dung he index 8 sheet cua V3 (FindSheetByMarkerOrName...).
'   - Nhanh 2HM dung he index 6 sheet cua engine 2HM (Find2Sheet...).
'   Hai he doc lap, khong xung dot ten ham (engine 2HM da co hau to "2").
' ================================================================

' ================================================================
' KHAI BAO BIEN MODULE (gop tu V3 + engine 2HM)
' ================================================================
' --- Tu V3 (nhanh 1 hang muc, 8 sheet) ---
Public Const SHEET_COUNT As Integer = 8
Public gSelSheet(1 To SHEET_COUNT) As Boolean
Public gAction As String         ' "ALIGN" | "PDF" | "PRINT" | "PREVIEW" | "BATCH"
Public gRenameDuAn As String     ' Ten du an moi (rong = khong doi)
Public gRenameHangMuc As String  ' Ten hang muc moi (rong = khong doi)
Public gFormOK As Boolean        ' True neu bam OK, False neu Cancel
Public gOrder(1 To SHEET_COUNT) As Integer  ' Thu tu in cua tung idx (1..8), mac dinh = idx
Public gBatchIsPDF As Boolean    ' Khi batch: True=xuat PDF, False=in may in
Public gIsHubMode As Boolean     ' True = chay tu FILE TONG (luon in file khac)
Public gLastErrDesc As String    ' Mo ta loi gan nhat khi xu ly 1 file (batch)
Public gGiaTongHopInTongCong As Boolean  ' True = in cot Tong cong (L) o Gia tong hop

' --- Nhanh NHIEU hang muc (2..MAX_HM khoi "HANG MUC:"), he 6 sheet ---
Public Const HM_SHEET_COUNT As Integer = 6   ' 6 loai sheet co the tham gia
Public Const MAX_HM As Integer = 20          ' so hang muc toi da / 1 file
Public gAssign2(1 To HM_SHEET_COUNT) As Integer   ' noi bo: 3 = sheet dang in theo khoi
Public gHMCount As Integer                   ' so hang muc cua file (0/1 = file 1 hang muc)
Public gHMNames(1 To MAX_HM) As String       ' ten hang muc (doc tu file)
Public gHMKinds(1 To MAX_HM) As String       ' "TB" | "XD" | ""
Public gSelN(1 To HM_SHEET_COUNT, 1 To MAX_HM) As Boolean  ' in sheet idx cho hang muc hm?
Public gOrderN(1 To HM_SHEET_COUNT) As Integer   ' so TT cua sheet trong PDF moi hang muc
Public gTplSel(1 To HM_SHEET_COUNT) As Boolean   ' mau chon sheet khi in ca thu muc
Public gHM_TenTheoHM As Boolean              ' TACH: dat ten PDF theo ten hang muc
Public gHM1Name As String, gHM2Name As String    ' = gHMNames(1), (2) - giu cho tuong thich
Public gLastErr2 As String
Public gGTH_InTongCong2 As Boolean

' ---- Cau hinh xuat 2 hang muc (moi) ----
' gHM_XuatKieu: "TACH" = moi hang muc 1 file ; "GHEP" = 1 file gop
Public gHM_XuatKieu As String
' gHM_OutFolder: thu muc luu thanh pham. Rong = cung thu muc file Excel goc.
Public gHM_OutFolder As String
' Ap dung cho ca nhanh 1 hang muc: thu muc output rieng (rong = canh file goc)
Public gOutFolder1 As String

' Ten PDF ep buoc cho tung hang muc (duong dan day du). Rong = dat ten theo quy tac thuong.
' Dung boi API_RunTong (xuat theo file tong hop).
Public gForcePdf(1 To 20) As String
Public gTongNoDGCT As Boolean           ' True = KHONG xuat sheet Don gia chi tiet
Public gTongFilter As String           ' chi xuat cac muc co duong dan file chi tiet chua chuoi nay (rong = tat ca)
Public gTongDryRun As Boolean           ' True = chi do lien ket, khong xuat PDF
Public gTongReport As String  ' bao cao lan chay API_RunTong gan nhat (doc qua Application.Run khi goi tu ngoai)

' --- Xuat theo file tong hop (xem cuoi module) ---
Private Type TongJob
    Seq As Long          ' thu tu dong la trong file tong (1..n)
    SumRow As Long
    Label As String      ' vd "A-1.1", "B.1-2.3"
    Nm As String         ' ten hang muc ghi trong file tong
    DetPath As String    ' file du toan chi tiet
    ThkpRow As Long      ' dong tren sheet THKP (0 = chua biet)
    Hm As Integer        ' hang muc (1..gHMCount) ; 0 = file 1 hang muc
    OutPdf As String
    DupOf As Long        ' > 0: chi so (trong mTJ) cua muc da xuat chung PDF
    Done As Boolean
    Note As String       ' ly do bo qua / canh bao
    Tag As String        ' "NN" hoac "NN.kk" (muc con khi 1 dong file tong tach ra nhieu file)
    Kind As Integer      ' TJ_THKP / TJ_TB / TJ_RAW
    SheetNm As String    ' sheet dich (TJ_TB, TJ_RAW)
    HmName As String     ' ten hang muc doc o dong duoc tro toi ('TH chi phi XD') khi khong co dong THKP
End Type

Private Const TJ_THKP As Integer = 0   ' file du toan: THKP -> Gia TH -> DGCT -> Gia thang
Private Const TJ_TB As Integer = 1     ' file du toan, lien ket vao 'TH chi phi TB': chi in sheet do
Private Const TJ_RAW As Integer = 2    ' sheet khong chuan (bao gia, bang tong hop...): in nguyen sheet

Private mTJ() As TongJob
Private mTN As Long
Private mTRx As Object
Private mIRx As Object
Private mVis As Object
Private mCurTop As Long
Private mCurLabel As String
Private mCurSumRow As Long
Private mWalkBudget As Long
Private mMidWb() As Workbook
Private mMidOwn() As Boolean
Private mMidKey() As String
Private mMidN As Long


' De quy thu thap file Excel (bo file tam ~$, bo file dang mo)
' ----------------------------------------------------------------
' DUNG FileSystemObject THAY VI Dir(): ham Dir() cua VBA di qua API
' ANSI cu, co the LOI THAM LANG (bi On Error Resume Next nuot mat)
' khi duong dan thu muc/file chua ky tu tieng Viet co dau kieu to
' hop (vd "HAM" viet hoa co dau ^ + huyen) ma khong khop voi
' codepage ANSI hien hanh cua may. FileSystemObject dung API Unicode
' on dinh hon nhieu, tranh duoc loi "khong tim thay file" gia.
Private Sub CollectExcelFiles(folderPath As String, inclSub As Boolean, _
        ByRef arr() As String, ByRef cnt As Long)
    Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error GoTo 0
    If fso Is Nothing Then Exit Sub
    CollectExcelFilesFSO fso, folderPath, inclSub, arr, cnt
End Sub

Private Sub CollectExcelFilesFSO(fso As Object, folderPath As String, inclSub As Boolean, _
        ByRef arr() As String, ByRef cnt As Long)
    Dim fld As Object
    On Error Resume Next
    Set fld = Nothing
    Set fld = fso.GetFolder(folderPath)
    On Error GoTo 0
    If fld Is Nothing Then Exit Sub

    Dim fl As Object, nm As String
    On Error Resume Next
    For Each fl In fld.Files
        nm = fl.Name
        If Left$(nm, 2) <> "~$" And IsExcelFileName(nm) Then
            cnt = cnt + 1
            If cnt > UBound(arr) Then ReDim Preserve arr(1 To cnt + 500)
            arr(cnt) = fl.Path
        End If
    Next fl
    On Error GoTo 0

    If inclSub Then
        Dim sf As Object
        On Error Resume Next
        For Each sf In fld.SubFolders
            CollectExcelFilesFSO fso, sf.Path & Application.PathSeparator, True, arr, cnt
        Next sf
        On Error GoTo 0
    End If
End Sub

' Kiem tra ten file co dung duoi Excel (.xls / .xlsx / .xlsm) - khong phan biet hoa thuong
Private Function IsExcelFileName(nm As String) As Boolean
    Dim low As String: low = LCase$(nm)
    IsExcelFileName = (Right$(low, 4) = ".xls") Or (Right$(low, 5) = ".xlsx") Or (Right$(low, 5) = ".xlsm")
End Function

' Ghi log loi ra file .txt cung thu muc (khi co loi) + tra ve duong dan
Private Function WriteBatchErrorLog(folderPath As String, logText As String) As String
    On Error GoTo Fail
    Dim p As String
    p = folderPath & "LOI_IN_" & Format(Now, "yyyymmdd_hhmmss") & ".txt"
    Dim fnum As Integer: fnum = FreeFile
    Open p For Output As #fnum
    Print #fnum, logText
    Close #fnum
    WriteBatchErrorLog = p
    Exit Function
Fail:
    WriteBatchErrorLog = ""
End Function

' Cau hinh gan sheet mac dinh thong minh cho file 2HM (dung trong batch)
'   Sheet tach khoi (THKP, Gia tong hop, Don gia CT) -> ca hai (3)
'   Cong trinh -> ca hai (3) neu co 2 khoi HM/THM, nguoc lai chi HM1
'   Gia thang (chung) -> HM2 ; TH chi phi TB (chung) -> HM1
' Chi bat sheet thuc su ton tai.

' ================================================================
' ============  FORM 2 COT + HO TRO (nhanh 2HM)  ===============
' ================================================================

' ---- Cac ham ho tro form 2 cot ----

Private Function IsSplitSheet2(idx As Integer) As Boolean
    ' Sheet co the tach 2 khoi: THKP(1), Gia tong hop(2), Don gia CT(3), Cong trinh(4)
    IsSplitSheet2 = (idx >= 1 And idx <= 4)
End Function

' Mac dinh tick cot HM1

' Mac dinh tick cot HM2

' ================================================================
' ============  ENGINE V3 (1 HANG MUC - 8 SHEET)  ================
' ================================================================

' act: ALIGN | PDF | PRINT | PREVIEW ; pdfPath chi dung khi act = PDF
Private Function AlignAndExportOneWorkbook(wbk As Workbook, act As String, pdfPath As String) As Boolean
    On Error GoTo Fail
    Dim ws(1 To SHEET_COUNT) As Worksheet
    Dim i As Integer, foundAny As Boolean
    For i = 1 To SHEET_COUNT
        Set ws(i) = FindSheetByMarkerOrName(wbk, i)
        If Not ws(i) Is Nothing Then foundAny = True
    Next i
    If Not foundAny Then AlignAndExportOneWorkbook = False: Exit Function

    ' QUAN TRONG: mot so phien ban Excel bao loi 1004 khi Select/Activate sheet
    ' luc ScreenUpdating=False va workbook vua moi mo. Bat tam ScreenUpdating
    ' + activate workbook de tranh loi. Se tra lai trang thai o cuoi ham.
    Dim savedSU As Boolean: savedSU = Application.ScreenUpdating
    Application.ScreenUpdating = True
    On Error Resume Next
    wbk.Activate
    wbk.Windows(1).Activate
    On Error GoTo Fail

    ' Chuan bi + PageSetup (chi cac sheet DUOC CHON va co trong file nay)
    For i = 1 To SHEET_COUNT
        If gSelSheet(i) And Not ws(i) Is Nothing Then
            On Error Resume Next
            PrepareSheetForPrint i, ws(i)
            HideAllComments ws(i)
            On Error GoTo Fail
        End If
    Next i
    For i = 1 To SHEET_COUNT
        If gSelSheet(i) And Not ws(i) Is Nothing Then
            On Error Resume Next
            ApplySettingsToWorksheet i, ws(i)
            ws(i).ResetAllPageBreaks
            On Error GoTo Fail
        End If
    Next i
    For i = 1 To SHEET_COUNT
        If gSelSheet(i) And Not ws(i) Is Nothing Then
            On Error Resume Next
            ws(i).Activate
            Application.Calculate
            Dim dv As Integer
            For dv = 1 To 8: DoEvents: Next dv
            EnsureBottomBordersInPrintArea ws(i), ws(i).PageSetup.PrintArea
            On Error GoTo Fail
        End If
    Next i

    ' Ep orientation (bo group truoc)
    Dim firstIdx As Integer: firstIdx = 0
    For i = 1 To SHEET_COUNT
        If gSelSheet(i) And Not ws(i) Is Nothing Then firstIdx = i: Exit For
    Next i
    If firstIdx = 0 Then AlignAndExportOneWorkbook = False: Exit Function
    On Error Resume Next
    ws(firstIdx).Select
    On Error GoTo Fail
    Application.PrintCommunication = True
    For i = 1 To SHEET_COUNT
        If gSelSheet(i) And Not ws(i) Is Nothing Then
            On Error Resume Next
            ws(i).Activate
            On Error GoTo Fail
            ForceOrientation ws(i), i
        End If
    Next i
    DoEvents

    ' Sap xep tab theo thu tu, roi group theo thu tu
    ' Bao dam wbk active truoc khi Move/Select (tranh loi 1004)
    On Error Resume Next
    wbk.Activate
    On Error GoTo Fail
    MoveTabsInOrderWbk wbk, ws
    On Error Resume Next
    wbk.Activate
    On Error GoTo Fail
    SelectGroupSelectedWbk wbk, ws

    ' Xuat PDF / In / Xem truoc
    Select Case act
    Case "PDF"
        KillFile2 pdfPath
        On Error Resume Next
        wbk.ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=pdfPath, Quality:=xlQualityStandard, _
            IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
        If Err.Number <> 0 Then
            gLastErrDesc = "Loi xuat PDF " & Err.Number & ": " & Err.Description
            Err.Clear
            On Error GoTo Fail
            AlignAndExportOneWorkbook = False
            Application.ScreenUpdating = savedSU
            Exit Function
        End If
        On Error GoTo Fail
    Case "PRINT"
        On Error Resume Next
        wbk.Windows(1).SelectedSheets.PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
        On Error GoTo Fail
    Case "PREVIEW"
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.DisplayAlerts = True
        wbk.Windows(1).SelectedSheets.PrintPreview
    Case Else
        ' ALIGN: chi can chinh -> bo group, dung o sheet dau
        On Error Resume Next
        ws(firstIdx).Select
        On Error GoTo Fail
    End Select

    AlignAndExportOneWorkbook = True
    Application.ScreenUpdating = savedSU
    Exit Function
Fail:
    gLastErrDesc = "Loi " & Err.Number & ": " & Err.Description
    Application.ScreenUpdating = savedSU
    AlignAndExportOneWorkbook = False
End Function

Private Sub MoveTabsInOrderWbk(wbk As Workbook, ws() As Worksheet)
    Dim cnt As Integer
    Dim ol As Variant
    ol = OrderedIdxList(cnt)
    If cnt = 0 Then Exit Sub
    Dim pos As Integer
    On Error Resume Next
    For pos = 1 To cnt
        If Not ws(ol(pos)) Is Nothing Then ws(ol(pos)).Move Before:=wbk.Sheets(pos)
    Next pos
    On Error GoTo 0
End Sub

Private Sub SelectGroupSelectedWbk(wbk As Workbook, ws() As Worksheet)
    Dim cnt As Integer
    Dim ol As Variant
    ol = OrderedIdxList(cnt)
    If cnt = 0 Then Exit Sub
    On Error Resume Next
    wbk.Activate
    On Error GoTo 0

    ' Tao mang TEN sheet theo thu tu, roi Select 1 lan (on dinh hon Select tung cai)
    Dim names() As String, nNm As Integer
    ReDim names(1 To cnt)
    Dim p As Integer
    For p = 1 To cnt
        If Not ws(ol(p)) Is Nothing Then
            nNm = nNm + 1
            names(nNm) = ws(ol(p)).Name
        End If
    Next p
    If nNm = 0 Then Exit Sub
    ReDim Preserve names(1 To nNm)

    On Error Resume Next
    wbk.Sheets(names).Select
    ' Neu cach mang that bai, fallback Select tung sheet
    If Err.Number <> 0 Then
        Err.Clear
        Dim first As Boolean: first = True
        For p = 1 To cnt
            If Not ws(ol(p)) Is Nothing Then
                If first Then
                    ws(ol(p)).Select
                    first = False
                Else
                    ws(ol(p)).Select False
                End If
            End If
        Next p
    End If
    On Error GoTo 0
End Sub

Private Function GetFileName(path As String) As String
    Dim pos As Long
    pos = InStrRev(path, Application.PathSeparator)
    If pos > 0 Then
        GetFileName = Mid(path, pos + 1)
    Else
        GetFileName = path
    End If
End Function

Private Function CountSelected() As Integer
    Dim i As Integer, c As Integer
    For i = 1 To SHEET_COUNT
        If gSelSheet(i) Then c = c + 1
    Next i
    CountSelected = c
End Function

Private Function FirstSelected() As Integer
    Dim i As Integer
    For i = 1 To SHEET_COUNT
        If gSelSheet(i) Then FirstSelected = i: Exit Function
    Next i
    FirstSelected = 1
End Function

Private Function OrderedIdxList(ByRef cnt As Integer) As Variant
    Dim arr(1 To SHEET_COUNT) As Integer
    Dim n As Integer, i As Integer
    n = 0
    For i = 1 To SHEET_COUNT
        If gSelSheet(i) Then
            n = n + 1
            arr(n) = i
        End If
    Next i
    ' insertion sort theo (gOrder, idx)
    Dim j As Integer, keyIdx As Integer
    For i = 2 To n
        keyIdx = arr(i)
        j = i - 1
        Do While j >= 1
            If (gOrder(arr(j)) > gOrder(keyIdx)) Or _
               (gOrder(arr(j)) = gOrder(keyIdx) And arr(j) > keyIdx) Then
                arr(j + 1) = arr(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        arr(j + 1) = keyIdx
    Next i
    cnt = n
    Dim res() As Integer
    ReDim res(1 To IIf(n = 0, 1, n))
    For i = 1 To n: res(i) = arr(i): Next i
    OrderedIdxList = res
End Function

Private Sub SelectGroupSelected(ws() As Worksheet)
    Dim cnt As Integer
    Dim ol As Variant
    ol = OrderedIdxList(cnt)
    Dim p As Integer, first As Boolean
    first = True
    For p = 1 To cnt
        If Not ws(ol(p)) Is Nothing Then
            If first Then
                ws(ol(p)).Select
                first = False
            Else
                ws(ol(p)).Select False
            End If
        End If
    Next p
End Sub

Private Sub ReapplyOrientationSelected(ws() As Worksheet)
    On Error Resume Next
    ws(FirstSelected()).Select   ' bo group de PageSetup ghi duoc
    On Error GoTo 0
    Application.PrintCommunication = True
    Dim i As Integer
    For i = 1 To SHEET_COUNT
        If gSelSheet(i) And Not ws(i) Is Nothing Then
            On Error Resume Next
            ws(i).Activate
            On Error GoTo 0
            ForceOrientation ws(i), i
        End If
    Next i
    DoEvents
End Sub

Private Sub ReadCurrentProjectNames(ws() As Worksheet, ByRef duAn As String, ByRef hangMuc As String)
    duAn = "": hangMuc = ""
    Dim i As Integer, r As Long
    Dim cellText As String, norm As String
    ' Uu tien lay tu THKP (idx1) hoac sheet dau tien co du lieu
    For i = 1 To SHEET_COUNT
        If Not ws(i) Is Nothing Then
            For r = 1 To 6
                cellText = SafeCellText(ws(i), r, 1)
                If Len(cellText) > 0 Then
                    norm = NormalizeVietnamese(cellText)
                    If duAn = "" And IsDuAnLabel(norm) Then
                        duAn = ExtractAfterColon(cellText)
                    ElseIf hangMuc = "" And IsHangMucLabel(norm) Then
                        hangMuc = ExtractAfterColon(cellText)
                    End If
                End If
            Next r
        End If
        If duAn <> "" And hangMuc <> "" Then Exit For
    Next i
End Sub

Private Sub ApplyRename(ws As Worksheet, idx As Integer, newDuAn As String, newHangMuc As String)
    Dim r As Long
    Dim cellText As String, norm As String
    Dim doneDuAn As Boolean, doneHangMuc As Boolean
    ' Quet dong 1-6, cot A. Chi doi DONG DAU TIEN moi loai (tranh doi trung).
    For r = 1 To 6
        cellText = SafeCellText(ws, r, 1)
        If Len(cellText) = 0 Then GoTo NextR
        norm = NormalizeVietnamese(cellText)

        If (Not doneDuAn) And IsDuAnLabel(norm) Then
            ws.Cells(r, 1).Value = RelabelKeepOrReplace(cellText, LabelDuAn(), newDuAn)
            doneDuAn = True
        ElseIf (Not doneHangMuc) And IsHangMucLabel(norm) Then
            ws.Cells(r, 1).Value = RelabelKeepOrReplace(cellText, LabelHangMuc(), newHangMuc)
            doneHangMuc = True
        End If
NextR:
    Next r
End Sub

Private Function IsDuAnLabel(norm As String) As Boolean
    IsDuAnLabel = (Left(norm, 5) = "du an") _
        Or (Left(norm, 8) = "ten du an") _
        Or (Left(norm, 14) = "ten cong trinh") _
        Or (Left(norm, 10) = "cong trinh" And InStr(norm, ":") > 0)
End Function

Private Function IsHangMucLabel(norm As String) As Boolean
    IsHangMucLabel = (Left(norm, 8) = "hang muc") _
        Or (Left(norm, 12) = "ten hang muc")
End Function

Private Function RelabelKeepOrReplace(original As String, newPrefix As String, newName As String) As String
    Dim content As String
    If Trim(newName) <> "" Then
        content = Trim(newName)
    Else
        content = ExtractAfterColon(original)  ' giu noi dung goc
    End If
    RelabelKeepOrReplace = newPrefix & ": " & content
End Function

Private Function ExtractAfterColon(s As String) As String
    Dim cur As String
    cur = Trim(s)
    Dim guard As Integer
    Do
        guard = guard + 1
        If guard > 10 Then Exit Do   ' chan vong lap vo han
        Dim pos As Long
        pos = InStr(cur, ":")
        If pos = 0 Then Exit Do
        Dim prefix As String, prefixNorm As String
        prefix = Trim(Left(cur, pos - 1))
        prefixNorm = NormalizeVietnamese(prefix)
        ' Chi boc neu phan TRUOC dau ":" DUNG LA mot nhan (khong phai ten that)
        If prefixNorm = "cong trinh" Or prefixNorm = "du an" _
           Or prefixNorm = "ten cong trinh" Or prefixNorm = "ten du an" _
           Or prefixNorm = "hang muc" Or prefixNorm = "ten hang muc" Then
            cur = Trim(Mid(cur, pos + 1))
        Else
            Exit Do   ' phan truoc ":" la noi dung that -> dung
        End If
    Loop
    ExtractAfterColon = cur
End Function

Private Function RebuildLabel(original As String, defaultPrefix As String, newName As String) As String
    Dim pos As Long
    pos = InStr(original, ":")
    If pos > 0 Then
        ' Giu nguyen tien to + dau ":" + 1 khoang trang
        RebuildLabel = RTrim(Left(original, pos)) & " " & newName
    Else
        RebuildLabel = defaultPrefix & ": " & newName
    End If
End Function

Private Function LabelDuAn() As String
    LabelDuAn = "D" & ChrW(7920) & " " & ChrW(193) & "N"        ' DU AN
End Function

Private Function LabelHangMuc() As String
    LabelHangMuc = "H" & ChrW(7840) & "NG M" & ChrW(7908) & "C" ' HANG MUC
End Function

Private Function SafeCellText(ws As Worksheet, r As Long, c As Long) As String
    Dim v As Variant
    SafeCellText = ""
    On Error Resume Next
    v = ws.Cells(r, c).Value
    On Error GoTo 0
    If Not IsEmpty(v) Then
        If Not IsError(v) Then SafeCellText = Trim(CStr(v))
    End If
End Function

Private Function GetSheetLabel(idx As Integer) As String
    Select Case idx
    Case 1: GetSheetLabel = "THKP hang muc"
    Case 2: GetSheetLabel = "Gia tong hop"
    Case 3: GetSheetLabel = "Don gia chi tiet"
    Case 4: GetSheetLabel = "Cong trinh"
    Case 5: GetSheetLabel = "Gia thang"
    Case 6: GetSheetLabel = "Tong hop VT"
    Case 7: GetSheetLabel = "TH chi phi XD"
    Case 8: GetSheetLabel = "TH chi phi TB"
    Case Else: GetSheetLabel = "Sheet " & idx
    End Select
End Function

Private Function GetOrientationLabel(idx As Integer) As String
    Select Case idx
    Case 1, 2, 7, 8: GetOrientationLabel = "Ngang"
    Case Else: GetOrientationLabel = "Doc"
    End Select
End Function

Private Sub ForceOrientation(ws As Worksheet, idx As Integer)
    Dim ori As XlPageOrientation
    Select Case idx
        Case 1, 2, 7, 8: ori = xlLandscape
        Case Else: ori = xlPortrait
    End Select
    Application.PrintCommunication = True
    Dim attempt As Integer
    For attempt = 1 To 3
        On Error Resume Next
        ws.PageSetup.PaperSize = xlPaperA4
        ws.PageSetup.Orientation = ori
        On Error GoTo 0
        DoEvents
        On Error Resume Next
        If ws.PageSetup.Orientation = ori Then On Error GoTo 0: Exit For
        On Error GoTo 0
    Next attempt
End Sub

' ================================================================
' FORCE PRINT TITLE ROWS - CHONG LOI EXCEL "NUOT LENH"
' ----------------------------------------------------------------
' Trieu chung: khi bang in tran sang trang 2, tieu de lap lai (Print
' Titles) bi thieu mat 1 dong (vi du header 2 dong "$4:$5" chi con
' lai 1 dong tren trang 2), lam lech cot / lech du lieu.
'
' Nguyen nhan: .PrintTitleRows duoc gan BEN TRONG khoi lenh
' "Application.PrintCommunication = False" (giong het loi Orientation
' va Header/Footer da tung gap va da fix rieng cho 2 muc do - xem
' ghi chu "V3.1 FIX FOOTER" o duoi). Khi PrintCommunication = False,
' Excel gom nhieu lenh PageSetup lai de tang toc, va doi khi "nuot"
' mat 1 phan gia tri cua PrintTitleRows (dac biet voi vung 2 dong)
' ma khong bao loi. Ket qua: gia tri DOC LAI ngay sau do co the dung,
' nhung khi thuc su in/xuat PDF (dac biet o trang thu 2 tro di) Excel
' lai dung ban ghi PageSetup noi bo bi thieu -> mat 1 dong tieu de.
'
' Cach fix: gan LAI .PrintTitleRows o NGOAI khoi PrintCommunication=False
' (sau khi da bat lai True), co doc lai de xac nhan, thu toi da 3 lan -
' dung mo hinh giong ForceOrientation da chung minh hieu qua.
' ================================================================
Private Sub ForcePrintTitleRows(ws As Worksheet, titleRowsStr As String)
    Dim attempt As Integer
    Application.PrintCommunication = True
    For attempt = 1 To 3
        On Error Resume Next
        ws.PageSetup.PrintTitleRows = titleRowsStr
        On Error GoTo 0
        DoEvents
        On Error Resume Next
        If ws.PageSetup.PrintTitleRows = titleRowsStr Then On Error GoTo 0: Exit For
        On Error GoTo 0
    Next attempt
End Sub

' ================================================================
' REFRESH MERGE TIEU DE 2 DONG - SUA LOI "MAT 1 DONG" O TRANG 2+
' ----------------------------------------------------------------
' Trieu chung THUC TE (khac voi lo PrintTitleRows da fix truoc do):
' header 2 dong (vd $4:$5 cua TH chi phi TB) co CAC COT MERGE KIEU
' HON HOP:
'   - Da so cot MERGE XUYEN CA 2 DONG lam 1 o (STT, Ten vat tu, Don
'     vi, So luong, Don gia, Thanh tien, Thanh tien sau thue, Ghi chu...)
'   - Rieng nhom "Thue VAT" (hay "Don gia"/"Kich thuoc" o sheet khac)
'     chi MERGE NGANG o dong tren (vd G4:H4 = "Thue VAT"), con dong
'     duoi lai TACH thanh cac o rieng le (G5="%", H5="Thanh tien").
'
' Kieu merge "hon hop" nay la truong hop Excel HAY LOI khi LAP LAI
' tieu de (Print Titles) tren trang thu 2 tro di: Excel "quen" mat
' dong tach (dong duoi cua nhom Thue VAT/Don gia/Kich thuoc), khien
' header o trang 2 chi con hien 1 dong va CAC O SAU DO BI LECH CHUY
' so voi du lieu that ben duoi - dung nhu anh chup nguoi dung gui.
'
' Cach fix: XOA TOAN BO MERGE trong vung header 2 dong, roi MERGE
' LAI CHINH XAC tu dau (giu nguyen y do: cot thuong -> merge xuyen 2
' dong; cot thuoc nhom split -> chi merge ngang o dong tren). Thao
' tac nay buoc Excel dung 1 bang ghi merge HOAN TOAN MOI, sach cac
' du lieu noi bo "stale" tich luy qua nhieu lan luu/sua file .xls,
' giup Excel LAP LAI tieu de dung tren moi trang.
'
' flatSplitPairs: mang 1 chieu, tung CAP (colStart, colEnd) noi tiep
' nhau danh dau cac nhom CHI merge ngang o dong tren (r1), KHONG
' merge xuyen xuong r2. Truyen Array() (rong) neu khong co nhom nao.
' Vi du TH chi phi TB (G,H la cot 7,8): Array(7, 8)
' Vi du Gia tong hop (F:H va I:K la Don gia/Thanh tien): Array(6, 8, 9, 11)
' ================================================================
Private Sub RefreshTwoRowHeaderMerge(ws As Worksheet, r1 As Long, r2 As Long, _
        colFrom As Long, colTo As Long, flatSplitPairs As Variant)
    On Error Resume Next
    Dim c As Long, p As Long
    Dim inSplit As Boolean
    Dim hasSplit As Boolean
    hasSplit = False
    If IsArray(flatSplitPairs) Then
        If UBound(flatSplitPairs) >= LBound(flatSplitPairs) Then hasSplit = True
    End If

    ' 1) Xoa merge toan bo vung header 2 dong (r1:r2, colFrom:colTo)
    ws.Range(ws.Cells(r1, colFrom), ws.Cells(r2, colTo)).UnMerge

    ' 2) Cot KHONG thuoc nhom split -> merge xuyen ca 2 dong (r1:r2)
    For c = colFrom To colTo
        inSplit = False
        If hasSplit Then
            For p = LBound(flatSplitPairs) To UBound(flatSplitPairs) Step 2
                If c >= flatSplitPairs(p) And c <= flatSplitPairs(p + 1) Then
                    inSplit = True
                    Exit For
                End If
            Next p
        End If
        If Not inSplit Then
            ws.Range(ws.Cells(r1, c), ws.Cells(r2, c)).Merge
        End If
    Next c

    ' 3) Tung nhom split -> merge NGANG rieng o dong tren (r1), giu
    '    dong duoi (r2) TACH RIENG tung o (khong merge)
    If hasSplit Then
        For p = LBound(flatSplitPairs) To UBound(flatSplitPairs) Step 2
            If flatSplitPairs(p + 1) > flatSplitPairs(p) Then
                ws.Range(ws.Cells(r1, flatSplitPairs(p)), ws.Cells(r1, flatSplitPairs(p + 1))).Merge
            End If
        Next p
    End If

    On Error GoTo 0
End Sub

Private Sub PrepareSheetForPrint(idx As Integer, ws As Worksheet)
    Dim lastRow7 As Long, lastRow8 As Long
    On Error Resume Next
    ws.Cells.ClearComments
    On Error GoTo 0

    Select Case idx
    Case 1 ' THKP hang muc: an dong 2,5,25 + an cot thua E,F,G (He so)
        On Error Resume Next
        ws.Rows("2:2").Hidden = True
        ws.Rows("5:5").Hidden = True
        ws.Rows("25:25").Hidden = True
        ' An cac cot thua giua "Cach tinh" (D) va "Gia tri" (H): E, F, G
        ws.Columns("E:G").Hidden = True
        On Error GoTo 0

    Case 2 ' Gia tong hop
        On Error Resume Next
        ws.Columns("L:L").ColumnWidth = 16
        ' Tuy chon: an cot Tong cong (L) neu nguoi dung khong tick
        If gGiaTongHopInTongCong Then
            ws.Columns("L:L").Hidden = False
        Else
            ws.Columns("L:L").Hidden = True
        End If
        On Error GoTo 0
        ' Header $4:$5: "Don gia" (F:H) va "Thanh tien" (I:K) chi merge
        ' ngang o dong 4, dong 5 tach Vat lieu/Nhan cong/May T.C rieng.
        RefreshTwoRowHeaderMerge ws, 4, 5, 1, 12, Array(6, 8, 9, 11)

    Case 3 ' Don gia chi tiet
        ' Khong thay doi

    Case 4 ' Cong trinh: WHITELIST theo V2 - hien A, E, F:L, O:P
        ' (an het cot khac gom B,C,D va M,N). An A..AM roi mo lai cot can hien.
        On Error Resume Next
        ws.Columns("A:AM").Hidden = True
        ws.Range("A:A").EntireColumn.Hidden = False
        ws.Range("E:E").EntireColumn.Hidden = False
        ws.Range("F:L").EntireColumn.Hidden = False
        ws.Range("O:P").EntireColumn.Hidden = False
        On Error GoTo 0
        ' Header $4:$5: "Kich thuoc" (I:K) chi merge ngang o dong 4,
        ' dong 5 tach Dai/Rong/Cao rieng.
        RefreshTwoRowHeaderMerge ws, 4, 5, 1, 16, Array(9, 11)

    Case 5 ' Gia thang: an E (Gia goc) va G:J (He so/Chinh/Gia sau VAT/Ma chuan) theo V2
        On Error Resume Next
        ws.Columns("E:E").Hidden = True
        ws.Columns("G:J").Hidden = True
        On Error GoTo 0

    Case 6 ' Tong hop VT: an cac cot phu (F,J,Q,R), chinh do rong O,P
        ' Theo macro: an F, J, Q:R; giu O,P; PrintArea den cot S
        On Error Resume Next
        ws.Columns("F:F").Hidden = True
        ws.Columns("J:J").Hidden = True
        ws.Columns("Q:R").Hidden = True
        ws.Columns("P:P").ColumnWidth = 18.33
        ws.Columns("O:O").ColumnWidth = 15.67
        ws.Columns("B:B").AutoFit
        On Error GoTo 0

    Case 7 ' TH chi phi XD: an dong 8 (dong trong/rac duoi bang) + moi dong trong hoan toan khac
        On Error Resume Next
        ws.Rows("8:8").Hidden = True
        On Error GoTo 0
        ' An them BAT KY dong nao trong khoang du lieu (A..G, tu dong 5)
        ' ma KHONG CO chu/so hien thi o bat ky cot nao (dong trong hoan toan),
        ' de bang gon lai, khong con dong rac lo lung khi in.
        lastRow7 = GetLastRowWithBorder(ws, 4, "G")
        HideFullyBlankRows ws, 5, lastRow7, 1, 7

    Case 8 ' TH chi phi TB: an moi dong TRONG HOAN TOAN (khong co du lieu o bat ky cot nao)
        ' Header 2 dong la $4:$5 -> du lieu bat dau tu dong 6.
        ' Header $4:$5: "Thue VAT" (G:H) chi merge ngang o dong 4,
        ' dong 5 tach %/Thanh tien rieng -> refresh truoc de tranh loi
        ' mat dong khi lap tieu de sang trang 2 (dung anh nguoi dung gui).
        RefreshTwoRowHeaderMerge ws, 4, 5, 1, 10, Array(7, 8)
        lastRow8 = GetLastRowWithBorder(ws, 4, "J")
        HideFullyBlankRows ws, 6, lastRow8, 1, 10
    End Select
End Sub

Private Sub ApplySettingsToWorksheet(idx As Integer, ws As Worksheet)
    Dim lastRow As Long
    Dim titleRowsStr As String   ' luu gia tri PrintTitleRows, se gan LAI o ngoai (xem ForcePrintTitleRows)

    Application.PrintCommunication = False
    With ws.PageSetup
        .PaperSize          = xlPaperA4
        .Zoom               = False
        .FitToPagesWide     = 1
        .FitToPagesTall     = False
        .LeftMargin         = Application.CentimetersToPoints(2.5)
        .RightMargin        = Application.CentimetersToPoints(1)
        .TopMargin          = Application.CentimetersToPoints(2.5)
        .BottomMargin       = Application.CentimetersToPoints(1.5)
        .HeaderMargin       = Application.CentimetersToPoints(1)
        .FooterMargin       = Application.CentimetersToPoints(1)
        .BlackAndWhite      = True
        .PrintQuality       = 600
        .CenterHorizontally = True
        .CenterVertically   = False
        .PrintGridlines     = False
        .PrintHeadings      = False
        .Draft              = False
        .Order              = xlDownThenOver
        .FirstPageNumber    = xlAutomatic
        .PrintComments      = xlPrintNoComments
        .PrintErrors        = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter    = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter       = True
        .AlignMarginsHeaderFooter       = True

        Select Case idx
        Case 1 ' THKP hang muc - Ngang
            lastRow = GetLastRowWithBorder(ws, 6, "H")
            .PrintArea         = "$A$1:$H$" & lastRow
            titleRowsStr       = "$6:$6"
            .Orientation       = xlLandscape

        Case 2 ' Gia tong hop - Ngang
            lastRow = GetLastDataRow(ws, 5, "L")
            ' In toi cot L (co Tong cong) hoac K (khong in Tong cong)
            If gGiaTongHopInTongCong Then
                .PrintArea     = "$A$1:$L$" & lastRow
            Else
                .PrintArea     = "$A$1:$K$" & lastRow
            End If
            titleRowsStr       = "$4:$5"
            .Orientation       = xlLandscape

        Case 3 ' Don gia chi tiet - Doc
            lastRow = GetLastDataRow(ws, 4, "H")
            .PrintArea         = "$A$1:$H$" & lastRow
            titleRowsStr       = "$4:$4"
            .Orientation       = xlPortrait

        Case 4 ' Cong trinh - Doc (chot theo V1 cu)
            lastRow = GetLastDataRow(ws, 5, "P")
            .PrintArea         = "$A$1:$P$" & lastRow
            titleRowsStr       = "$4:$5"
            .Orientation       = xlPortrait

        Case 5 ' Gia thang - Doc
            lastRow = GetLastDataRow(ws, 4, "K")
            .PrintArea         = "$A$1:$K$" & lastRow
            titleRowsStr       = "$4:$4"
            .Orientation       = xlPortrait

        Case 6 ' Tong hop VT - Doc (in den cot S; cot phu da an)
            lastRow = GetLastDataRow(ws, 5, "S")
            .PrintArea         = "$A$1:$S$" & lastRow
            titleRowsStr       = "$4:$4"
            .Orientation       = xlPortrait

        Case 7 ' TH chi phi XD - Ngang (bang nho)
            lastRow = GetLastRowWithBorder(ws, 4, "G")
            .PrintArea         = "$A$1:$G$" & lastRow
            titleRowsStr       = "$4:$4"
            .Orientation       = xlLandscape

        Case 8 ' TH chi phi TB - Ngang (bang nho)
            lastRow = GetLastRowWithBorder(ws, 4, "J")
            .PrintArea         = "$A$1:$J$" & lastRow
            titleRowsStr       = "$4:$5"
            .Orientation       = xlLandscape
        End Select
    End With
    Application.PrintCommunication = True

    ' V3.1 FIX FOOTER: dat header/footer SAU khi bat PrintCommunication=True,
    ' vi neu dat trong khoi PrintCommunication=False, Excel se NUOT lenh
    ' (giong loi orientation). Phai xoa o day thi footer moi thuc su mat.
    On Error Resume Next
    With ws.PageSetup
        ' Tat che do trang chan/le + trang dau khac truoc
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        ' Xoa 6 vung header/footer chinh
        .CenterHeader = "": .LeftHeader = "": .RightHeader = ""
        .CenterFooter = "": .LeftFooter = "": .RightFooter = ""
        ' Xoa EvenPage (phong khi file da bat OddAndEven truoc do)
        .EvenPage.LeftHeader.Text = "": .EvenPage.CenterHeader.Text = "": .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = "": .EvenPage.CenterFooter.Text = "": .EvenPage.RightFooter.Text = ""
        ' Xoa FirstPage (phong khi file da bat DifferentFirstPage truoc do)
        .FirstPage.LeftHeader.Text = "": .FirstPage.CenterHeader.Text = "": .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = "": .FirstPage.CenterFooter.Text = "": .FirstPage.RightFooter.Text = ""
    End With
    On Error GoTo 0

    ForceOrientation ws, idx

    ' V4 FIX TIEU DE MAT 1 DONG O TRANG 2: gan lai PrintTitleRows o ngoai
    ' khoi PrintCommunication=False, tranh bi Excel "nuot lenh" (xem chu
    ' thich tai khai bao ForcePrintTitleRows o tren).
    ForcePrintTitleRows ws, titleRowsStr

    ' V3.1 FIX #####: AutoFit cac cot SO tren MOI sheet de khong bi ###.
    AutoFitNumericColumns ws, idx
End Sub

Private Sub AutoFitNumericColumns(ws As Worksheet, idx As Integer)
    On Error Resume Next
    Select Case idx
    Case 1 ' THKP: cot gia tri la H (E,F,G da an)
        ws.Columns("H:H").AutoFit
        EnsureMinWidth ws, "H", 14
    Case 2 ' Gia tong hop: cot don gia + thanh tien (F..L)
        ws.Columns("F:L").AutoFit
        EnsureMinWidth ws, "I", 12
        EnsureMinWidth ws, "J", 12
        EnsureMinWidth ws, "K", 12
        EnsureMinWidth ws, "L", 14
    Case 3 ' Don gia chi tiet: cot dinh muc/don gia/thanh tien (E..H)
        ws.Columns("E:H").AutoFit
    Case 4 ' Cong trinh: AutoFit cac cot HIEN (bo qua M,N da an)
        ws.Columns("E:L").AutoFit
        ws.Columns("O:P").AutoFit
        ' Cot tong cong (P) hay bi ### -> dat min rong hon
        EnsureMinWidth ws, "P", 15
        EnsureMinWidth ws, "O", 13
    Case 5 ' Gia thang: chi AutoFit rieng cot F va K (dang HIEN).
        ' FIX: truoc day dung "F:K" gop ca vung G:J (da bi PrepareSheetForPrint
        ' Case 5 an truoc do) -> AutoFit tren dai cot co LAN cot an lam Excel
        ' tu dong MO LAI (Hidden=False) cac cot G:J, khien sheet Gia thang khi
        ' in "1 hang muc" hien du ca He so/Chinh/Gia sau VAT/Ma chuan (dung
        ' loi nguoi dung bao cao, doi chieu anh chup gui kem). Dong bo voi
        ' AutoFitNum2 (nhanh 2 hang muc, dang dung dung "F:F" + "K:K" rieng
        ' le, khong bi loi nay).
        ws.Columns("F:F").AutoFit
        ws.Columns("K:K").AutoFit
    Case 6 ' Tong hop VT: cot so (in den S)
        ws.Columns("G:S").AutoFit
    Case 7 ' TH chi phi XD: cot gia tri
        ws.Columns("C:G").AutoFit
    Case 8 ' TH chi phi TB: cot gia tri
        ws.Columns("C:J").AutoFit
    End Select
    On Error GoTo 0
End Sub

Private Sub EnsureMinWidth(ws As Worksheet, col As String, minW As Double)
    On Error Resume Next
    If ws.Columns(col & ":" & col).ColumnWidth < minW Then
        ws.Columns(col & ":" & col).ColumnWidth = minW
    End If
    On Error GoTo 0
End Sub

Private Function FindSheetByMarkerOrName(wb As Workbook, idx As Integer) As Worksheet
    Dim s As Worksheet, bestSheet As Worksheet
    Dim bestScore As Long, sc As Long
    bestScore = 0
    Set bestSheet = Nothing
    For Each s In wb.Worksheets
        If s.Visible = xlSheetVisible Or idx = 5 Then sc = ScoreSheetForIdx(s, idx) Else sc = 0   ' sheet an = khong dung (tru Gia thang)
        If sc > bestScore Then
            bestScore = sc
            Set bestSheet = s
        End If
    Next s
    If bestScore >= 50 Then
        Set FindSheetByMarkerOrName = bestSheet
        If idx = 5 Then
            If bestSheet.Visible <> xlSheetVisible Then
                On Error Resume Next
                bestSheet.Visible = xlSheetVisible
                On Error GoTo 0
            End If
        End If
    Else
        Set FindSheetByMarkerOrName = Nothing
    End If
End Function

Private Function ScoreSheetForIdx(ws As Worksheet, idx As Integer) As Long
    Dim sc As Long: sc = 0
    Dim normName As String
    normName = NormalizeVietnamese(ws.Name)

    ' 1. Marker tieu de (+100)
    Dim mkArr() As String
    mkArr = GetMarkerTexts(idx)
    If SheetHeaderContainsMarker(ws, mkArr) Then sc = sc + 100

    ' 2. Ten tab
    Dim tabArr() As String
    tabArr = GetPossibleTabNames(idx)
    Dim i As Long
    Dim hasExact As Boolean, hasContains As Boolean
    For i = LBound(tabArr) To UBound(tabArr)
        If normName = tabArr(i) Then hasExact = True: Exit For
    Next i
    If hasExact Then sc = sc + 50
    For i = LBound(tabArr) To UBound(tabArr)
        If Len(tabArr(i)) > 0 Then
            If InStr(normName, tabArr(i)) > 0 Then hasContains = True: Exit For
        End If
    Next i
    If hasContains Then sc = sc + 10

    ' 3. Phat suffix sheet phu (-30)
    Dim padded As String, sufArr() As String
    padded = " " & normName & " "
    sufArr = Split(" tl | qt | thau | tham | tham dinh | goi thau ", "|")
    For i = LBound(sufArr) To UBound(sufArr)
        If InStr(padded, sufArr(i)) > 0 Then sc = sc - 30: Exit For
    Next i

    If sc < 0 Then sc = 0
    ScoreSheetForIdx = sc
End Function

Private Function GetMarkerTexts(idx As Integer) As String()
    Dim m() As String
    Select Case idx
    Case 1
        m = Split("bang tong hop du toan chi phi xay dung|" & _
                  "tong hop du toan chi phi xay dung", "|")
    Case 2
        m = Split("gia xay dung tong hop hang muc cong trinh|" & _
                  "gia xay dung tong hop hang muc", "|")
    Case 3
        m = Split("don gia xay dung chi tiet hang muc cong trinh|" & _
                  "don gia xay dung chi tiet hang muc|" & _
                  "don gia xay dung chi tiet", "|")
    Case 4
        m = Split("bang du toan hang muc cong trinh|" & _
                  "du toan hang muc cong trinh|du toan hang muc", "|")
    Case 5
        m = Split("bang gia thong bao cong trinh|" & _
                  "gia thong bao cong trinh|bang gia thong bao|gia thong bao", "|")
    Case 6 ' Tong hop VT
        m = Split("bang tong hop vat tu va chenh lech gia|" & _
                  "tong hop vat tu va chenh lech|" & _
                  "tong hop vat tu va chenh lech gia", "|")
    Case 7 ' TH chi phi XD
        m = Split("bang tong hop chi phi xay dung cong trinh|" & _
                  "tong hop chi phi xay dung cong trinh", "|")
    Case 8 ' TH chi phi TB
        m = Split("bang tong hop chi phi thiet bi cong trinh|" & _
                  "tong hop chi phi thiet bi cong trinh|" & _
                  "tong hop chi phi thiet bi", "|")
    End Select
    GetMarkerTexts = m
End Function

Private Function GetPossibleTabNames(idx As Integer) As String()
    Dim p() As String
    Select Case idx
    Case 1
        p = Split("thkp hang muc|thkp|thkphangmuc", "|")
    Case 2
        p = Split("gia tong hop|giatonghop|gth|bang gia tong hop", "|")
    Case 3
        p = Split("don gia chi tiet|dgct|dongiachitiet|don gia xay dung chi tiet", "|")
    Case 4
        p = Split("cong trinh|congtrinh|ct", "|")
    Case 5
        p = Split("gia thang|giathang|gt|bang gia thang", "|")
    Case 6
        p = Split("tong hop vt|tonghopvt|th vt|thvt", "|")
    Case 7
        p = Split("th chi phi xd|thcpxd|th cp xd|tong hop chi phi xd", "|")
    Case 8
        p = Split("th chi phi tb|thcptb|th cp tb|tong hop chi phi tb", "|")
    End Select
    GetPossibleTabNames = p
End Function

Private Function NormalizeVietnamese(txt As String) As String
    Dim s As String
    s = txt
    
    ' Doi tab/newline thanh space
    s = Replace(s, Chr(9), " ")
    s = Replace(s, Chr(10), " ")
    s = Replace(s, Chr(13), " ")
    
    ' Bo dau tieng Viet (dung ChrW de tranh mojibake khi import .bas)
    s = Replace(s, ChrW(224), "a")
    s = Replace(s, ChrW(225), "a")
    s = Replace(s, ChrW(7843), "a")
    s = Replace(s, ChrW(227), "a")
    s = Replace(s, ChrW(7841), "a")
    s = Replace(s, ChrW(226), "a")
    s = Replace(s, ChrW(7847), "a")
    s = Replace(s, ChrW(7845), "a")
    s = Replace(s, ChrW(7849), "a")
    s = Replace(s, ChrW(7851), "a")
    s = Replace(s, ChrW(7853), "a")
    s = Replace(s, ChrW(259), "a")
    s = Replace(s, ChrW(7857), "a")
    s = Replace(s, ChrW(7855), "a")
    s = Replace(s, ChrW(7859), "a")
    s = Replace(s, ChrW(7861), "a")
    s = Replace(s, ChrW(7863), "a")
    s = Replace(s, ChrW(192), "a")
    s = Replace(s, ChrW(193), "a")
    s = Replace(s, ChrW(7842), "a")
    s = Replace(s, ChrW(195), "a")
    s = Replace(s, ChrW(7840), "a")
    s = Replace(s, ChrW(194), "a")
    s = Replace(s, ChrW(7846), "a")
    s = Replace(s, ChrW(7844), "a")
    s = Replace(s, ChrW(7848), "a")
    s = Replace(s, ChrW(7850), "a")
    s = Replace(s, ChrW(7852), "a")
    s = Replace(s, ChrW(258), "a")
    s = Replace(s, ChrW(7856), "a")
    s = Replace(s, ChrW(7854), "a")
    s = Replace(s, ChrW(7858), "a")
    s = Replace(s, ChrW(7860), "a")
    s = Replace(s, ChrW(7862), "a")
    s = Replace(s, ChrW(232), "e")
    s = Replace(s, ChrW(233), "e")
    s = Replace(s, ChrW(7867), "e")
    s = Replace(s, ChrW(7869), "e")
    s = Replace(s, ChrW(7865), "e")
    s = Replace(s, ChrW(234), "e")
    s = Replace(s, ChrW(7873), "e")
    s = Replace(s, ChrW(7871), "e")
    s = Replace(s, ChrW(7875), "e")
    s = Replace(s, ChrW(7877), "e")
    s = Replace(s, ChrW(7879), "e")
    s = Replace(s, ChrW(200), "e")
    s = Replace(s, ChrW(201), "e")
    s = Replace(s, ChrW(7866), "e")
    s = Replace(s, ChrW(7868), "e")
    s = Replace(s, ChrW(7864), "e")
    s = Replace(s, ChrW(202), "e")
    s = Replace(s, ChrW(7872), "e")
    s = Replace(s, ChrW(7870), "e")
    s = Replace(s, ChrW(7874), "e")
    s = Replace(s, ChrW(7876), "e")
    s = Replace(s, ChrW(7878), "e")
    s = Replace(s, ChrW(236), "i")
    s = Replace(s, ChrW(237), "i")
    s = Replace(s, ChrW(7881), "i")
    s = Replace(s, ChrW(297), "i")
    s = Replace(s, ChrW(7883), "i")
    s = Replace(s, ChrW(204), "i")
    s = Replace(s, ChrW(205), "i")
    s = Replace(s, ChrW(7880), "i")
    s = Replace(s, ChrW(296), "i")
    s = Replace(s, ChrW(7882), "i")
    s = Replace(s, ChrW(242), "o")
    s = Replace(s, ChrW(243), "o")
    s = Replace(s, ChrW(7887), "o")
    s = Replace(s, ChrW(245), "o")
    s = Replace(s, ChrW(7885), "o")
    s = Replace(s, ChrW(244), "o")
    s = Replace(s, ChrW(7891), "o")
    s = Replace(s, ChrW(7889), "o")
    s = Replace(s, ChrW(7893), "o")
    s = Replace(s, ChrW(7895), "o")
    s = Replace(s, ChrW(7897), "o")
    s = Replace(s, ChrW(417), "o")
    s = Replace(s, ChrW(7901), "o")
    s = Replace(s, ChrW(7899), "o")
    s = Replace(s, ChrW(7903), "o")
    s = Replace(s, ChrW(7905), "o")
    s = Replace(s, ChrW(7907), "o")
    s = Replace(s, ChrW(210), "o")
    s = Replace(s, ChrW(211), "o")
    s = Replace(s, ChrW(7886), "o")
    s = Replace(s, ChrW(213), "o")
    s = Replace(s, ChrW(7884), "o")
    s = Replace(s, ChrW(212), "o")
    s = Replace(s, ChrW(7890), "o")
    s = Replace(s, ChrW(7888), "o")
    s = Replace(s, ChrW(7892), "o")
    s = Replace(s, ChrW(7894), "o")
    s = Replace(s, ChrW(7896), "o")
    s = Replace(s, ChrW(416), "o")
    s = Replace(s, ChrW(7900), "o")
    s = Replace(s, ChrW(7898), "o")
    s = Replace(s, ChrW(7902), "o")
    s = Replace(s, ChrW(7904), "o")
    s = Replace(s, ChrW(7906), "o")
    s = Replace(s, ChrW(249), "u")
    s = Replace(s, ChrW(250), "u")
    s = Replace(s, ChrW(7911), "u")
    s = Replace(s, ChrW(361), "u")
    s = Replace(s, ChrW(7909), "u")
    s = Replace(s, ChrW(432), "u")
    s = Replace(s, ChrW(7915), "u")
    s = Replace(s, ChrW(7913), "u")
    s = Replace(s, ChrW(7917), "u")
    s = Replace(s, ChrW(7919), "u")
    s = Replace(s, ChrW(7921), "u")
    s = Replace(s, ChrW(217), "u")
    s = Replace(s, ChrW(218), "u")
    s = Replace(s, ChrW(7910), "u")
    s = Replace(s, ChrW(360), "u")
    s = Replace(s, ChrW(7908), "u")
    s = Replace(s, ChrW(431), "u")
    s = Replace(s, ChrW(7914), "u")
    s = Replace(s, ChrW(7912), "u")
    s = Replace(s, ChrW(7916), "u")
    s = Replace(s, ChrW(7918), "u")
    s = Replace(s, ChrW(7920), "u")
    s = Replace(s, ChrW(7923), "y")
    s = Replace(s, ChrW(253), "y")
    s = Replace(s, ChrW(7927), "y")
    s = Replace(s, ChrW(7929), "y")
    s = Replace(s, ChrW(7925), "y")
    s = Replace(s, ChrW(7922), "y")
    s = Replace(s, ChrW(221), "y")
    s = Replace(s, ChrW(7926), "y")
    s = Replace(s, ChrW(7928), "y")
    s = Replace(s, ChrW(7924), "y")
    s = Replace(s, ChrW(273), "d")
    s = Replace(s, ChrW(272), "d")
    
    ' Lowercase ASCII
    s = LCase(s)
    
    ' Bo dau cau pho bien
    s = Replace(s, ":", " ")
    s = Replace(s, ";", " ")
    s = Replace(s, ",", " ")
    s = Replace(s, ".", " ")
    s = Replace(s, "!", " ")
    s = Replace(s, "?", " ")
    s = Replace(s, "-", " ")
    s = Replace(s, "_", " ")
    s = Replace(s, "(", " ")
    s = Replace(s, ")", " ")
    s = Replace(s, "[", " ")
    s = Replace(s, "]", " ")
    s = Replace(s, "{", " ")
    s = Replace(s, "}", " ")
    s = Replace(s, "/", " ")
    s = Replace(s, "\", " ")
    s = Replace(s, Chr(34), " ")
    s = Replace(s, "*", " ")
    s = Replace(s, "#", " ")
    
    ' Bo space thua
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    
    NormalizeVietnamese = Trim(s)
End Function

' ================================================================
' AN DONG TRONG HOAN TOAN (khong co chu/so hien thi o bat ky cot nao)
' ----------------------------------------------------------------
' Dung cho cac bang tong hop nho (TH chi phi XD/TB...) co the con lai
' vai dong du phong/rac giua du lieu va dong TONG CONG (nhu trong file
' mau: cac dong sau item cuoi, truoc "TONG CONG", khong co STT/Ten vat
' tu/gia tri nao). AN (Hidden = True) thay vi XOA de khong lam sai
' cong thuc SUM/dia chi tham chieu cua cac sheet khac.
'
' Dung .Text (chu KHONG dung .Value) de kiem tra "trong" theo dung
' NHU MAT NGUOI THAY khi in: mot o co cong thuc tra ve 0 nhung dinh
' dang so an so 0 (vi du dinh dang "#,##0;;;") se co .Text = "" nen
' van duoc coi la trong, giong nhu trong anh chup nguoi dung gui.
' ================================================================
Private Sub HideFullyBlankRows(ws As Worksheet, firstDataRow As Long, lastRow As Long, _
        firstCol As Long, lastCol As Long)
    Dim r As Long, c As Long, blank As Boolean
    On Error Resume Next
    For r = firstDataRow To lastRow
        blank = True
        For c = firstCol To lastCol
            If Len(Trim$(ws.Cells(r, c).Text)) > 0 Then
                blank = False
                Exit For
            End If
        Next c
        If blank Then ws.Rows(r).Hidden = True
    Next r
    On Error GoTo 0
End Sub

Private Function GetLastDataRow(ws As Worksheet, startRow As Long, endCol As String) As Long
    Dim lastRow As Long
    Dim tempRow As Long
    Dim colCheck As Variant
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    For Each colCheck In Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P")
        If colCheck <= endCol Then
            On Error Resume Next
            tempRow = ws.Cells(ws.Rows.Count, colCheck).End(xlUp).Row
            On Error GoTo 0
            If tempRow > lastRow Then lastRow = tempRow
        End If
    Next colCheck
    
    If lastRow < startRow Then lastRow = startRow
    
    GetLastDataRow = lastRow
End Function

Private Function GetLastRowWithBorder(ws As Worksheet, startRow As Long, endCol As String) As Long
    Dim maxScan As Long
    maxScan = GetLastDataRow(ws, startRow, endCol)
    
    ' Quet tu duoi len, tim dong dau tien co border
    Dim r As Long
    Dim c As Long
    Dim endColNum As Long
    endColNum = ws.Range(endCol & "1").Column
    
    For r = maxScan To startRow Step -1
        For c = 1 To endColNum
            Dim cell As Range
            Set cell = ws.Cells(r, c)
            
            ' Kiem tra 4 border cua cell
            If cell.Borders(xlEdgeTop).LineStyle <> xlNone Then
                GetLastRowWithBorder = r
                Exit Function
            End If
            If cell.Borders(xlEdgeBottom).LineStyle <> xlNone Then
                GetLastRowWithBorder = r
                Exit Function
            End If
            If cell.Borders(xlEdgeLeft).LineStyle <> xlNone Then
                GetLastRowWithBorder = r
                Exit Function
            End If
            If cell.Borders(xlEdgeRight).LineStyle <> xlNone Then
                GetLastRowWithBorder = r
                Exit Function
            End If
        Next c
    Next r
    
    ' Khong tim thay border -> fallback
    GetLastRowWithBorder = maxScan
End Function

Private Sub EnsureBottomBordersInPrintArea(ws As Worksheet, printRangeAddr As String)
    On Error Resume Next

    If printRangeAddr = "" Then Exit Sub
    Dim rng As Range
    Set rng = ws.Range(printRangeAddr)
    If rng Is Nothing Then Exit Sub

    Dim firstRow As Long, lastRow As Long, firstCol As Long, lastCol As Long
    firstRow = rng.Row
    lastRow  = rng.Row + rng.Rows.Count - 1
    firstCol = rng.Column
    lastCol  = rng.Column + rng.Columns.Count - 1

    Dim r As Long, c As Long
    Dim bdr As Border

    ' --- Nang cap moi border bottom hairline -> thin trong vung in ---
    For r = firstRow To lastRow
        If ws.Rows(r).Hidden Then GoTo NextR
        For c = firstCol To lastCol
            Set bdr = ws.Cells(r, c).Borders(xlEdgeBottom)
            ' Chi dong vao cell DA CO border bottom (bo qua o khong co duong)
            If bdr.LineStyle <> xlNone Then
                ' Neu dang la hairline thi nang len thin; thin/medium giu nguyen
                If bdr.Weight = xlHairline Then
                    bdr.LineStyle = xlContinuous
                    bdr.Weight    = xlThin
                End If
            End If
        Next c
NextR:
    Next r

    ' --- Dam bao DONG CUOI cua PrintArea co border bottom (phong khi thieu) ---
    For c = firstCol To lastCol
        With ws.Cells(lastRow, c).Borders(xlEdgeBottom)
            If .LineStyle = xlNone Then
                .LineStyle = xlContinuous
                .Weight    = xlThin
                .ColorIndex = xlAutomatic
            End If
        End With
    Next c

    On Error GoTo 0
End Sub

Private Sub HideAllComments(ws As Worksheet)
    On Error Resume Next
    Dim cm As Comment
    For Each cm In ws.Comments
        cm.Visible = False
    Next cm
    On Error GoTo 0
End Sub

Private Function SheetHeaderContainsMarker(ws As Worksheet, markerTexts() As String) As Boolean
    SheetHeaderContainsMarker = False
    
    ' Kiem tra mang marker co hop le
    If (Not IsArrayInitialized(markerTexts)) Then Exit Function
    
    Dim mLo As Long, mHi As Long
    mLo = LBound(markerTexts)
    mHi = UBound(markerTexts)
    If mHi < mLo Then Exit Function
    
    Dim r As Long, c As Long
    Dim cellVal As Variant
    Dim cellText As String
    Dim mIdx As Long
    Dim mk As String
    
    For r = 1 To 5
        For c = 1 To 6
            cellVal = Empty
            On Error Resume Next
            cellVal = ws.Cells(r, c).Value
            On Error GoTo 0
            
            If Not IsEmpty(cellVal) Then
                If Not IsError(cellVal) Then
                    cellText = NormalizeVietnamese(CStr(cellVal))
                    If Len(cellText) >= 10 Then
                        For mIdx = mLo To mHi
                            mk = Trim(markerTexts(mIdx))
                            If Len(mk) > 0 Then
                                If InStr(cellText, mk) > 0 Then
                                    SheetHeaderContainsMarker = True
                                    Exit Function
                                End If
                            End If
                        Next mIdx
                    End If
                End If
            End If
        Next c
    Next r
End Function

Private Function IsArrayInitialized(arr() As String) As Boolean
    On Error Resume Next
    Dim lb As Long, ub As Long
    lb = LBound(arr)
    If Err.Number <> 0 Then
        IsArrayInitialized = False
        Err.Clear
        Exit Function
    End If
    ub = UBound(arr)
    If Err.Number <> 0 Then
        IsArrayInitialized = False
        Err.Clear
        Exit Function
    End If
    IsArrayInitialized = (ub >= lb)
    On Error GoTo 0
End Function

' ================================================================
' ============  ENGINE 2HM (2 HANG MUC - TACH KHOI)  ============
' ================================================================

' File co >= 2 khoi "HANG MUC:" ? Dien gHMCount / gHMNames / gHMKinds.
' (Ten giu nguyen cho tuong thich; thuc te nhan dien N hang muc.)
Private Function DetectTwoHangMuc(wb As Workbook) As Boolean
    DetectTwoHangMuc = False
    gHMCount = 0: gHM1Name = "": gHM2Name = ""
    Dim k As Integer
    For k = 1 To MAX_HM: gHMNames(k) = "": gHMKinds(k) = "": Next k

    ' Lay sheet co NHIEU khoi nhat trong Don gia CT / Gia tong hop / THKP / Cong trinh
    Dim tryIdx As Variant, ws As Worksheet
    Dim rows() As Long, cnt As Integer, names() As String
    Dim bestCnt As Integer, bestNames() As String
    For Each tryIdx In Array(3, 2, 1, 4)
        Set ws = Find2Sheet(wb, CInt(tryIdx))
        If Not ws Is Nothing Then
            FindHangMucHeaderRows ws, GetMarkerCol2(CInt(tryIdx)), rows, names, cnt
            If cnt > bestCnt Then bestCnt = cnt: bestNames = names
        End If
    Next tryIdx
    If bestCnt < 2 Then Exit Function

    If bestCnt > MAX_HM Then bestCnt = MAX_HM
    gHMCount = bestCnt
    For k = 1 To bestCnt
        gHMNames(k) = bestNames(k)
        gHMKinds(k) = KindFromName(bestNames(k))
    Next k
    gHM1Name = gHMNames(1): gHM2Name = gHMNames(2)
    DetectTwoHangMuc = True
End Function

' ================================================================
' NHIEU HANG MUC (N >= 2) - 1 file co nhieu khoi "HANG MUC:"
' ----------------------------------------------------------------
' gHMCount / gHMNames do DetectTwoHangMuc dien.
' gSelN(idx, hm) = in sheet idx cho hang muc hm ; gOrderN(idx) = so TT.
' Sheet tach khoi (THKP, Gia tong hop, Don gia CT, Cong trinh): moi
' hang muc in DUNG khoi cua no (khop theo TEN hang muc, roi loai TB/XD,
' roi thu tu). Sheet chung (Gia thang, TH chi phi TB): in nguyen sheet.
'   TACH: moi hang muc 1 PDF (<ten excel>_HMk.pdf hoac -<ten HM>.pdf)
'   GHEP: tat ca hang muc 1 PDF <ten excel>.pdf (1 lan ExportAsFixedFormat;
'         sheet dung lai cho hang muc sau = ban sao tam).
' ================================================================
Private Function ProcessWorkbookNHM(wb As Workbook, ByRef outFiles() As String, ByRef nOut As Integer) As Boolean
    On Error GoTo Fail
    ProcessWorkbookNHM = False
    nOut = 0
    ReDim outFiles(1 To MAX_HM)

    Dim savedSU As Boolean: savedSU = Application.ScreenUpdating
    Application.ScreenUpdating = True
    On Error Resume Next
    wb.Activate
    wb.Windows(1).Activate
    On Error GoTo Fail

    Dim ws(1 To HM_SHEET_COUNT) As Worksheet
    Dim k As Integer
    For k = 1 To HM_SHEET_COUNT
        Set ws(k) = Find2Sheet(wb, k)
    Next k

    Dim folder As String, baseName As String
    If Len(gHM_OutFolder) > 0 Then
        folder = gHM_OutFolder
        If Right$(folder, 1) <> Application.PathSeparator Then folder = folder & Application.PathSeparator
    Else
        folder = wb.Path & Application.PathSeparator
    End If
    Dim dotP As Long: dotP = InStrRev(wb.Name, ".")
    baseName = IIf(dotP > 1, Left(wb.Name, dotP - 1), wb.Name)

    Dim ghep As Boolean: ghep = (UCase$(gHM_XuatKieu) = "GHEP")
    Dim tmpNames() As String, nTmp As Integer
    ReDim tmpNames(1 To HM_SHEET_COUNT * MAX_HM)
    Dim used(1 To HM_SHEET_COUNT) As Boolean
    Dim allNames() As String, nAll As Integer
    ReDim allNames(1 To HM_SHEET_COUNT * MAX_HM)
    Dim warn As String

    ' Sheet nao that su chia khoi (>= 2 khoi)?
    Dim isSplit(1 To HM_SHEET_COUNT) As Boolean
    For k = 1 To 4
        If Not ws(k) Is Nothing Then
            ws(k).Cells.EntireRow.Hidden = False
            isSplit(k) = (BlockCount(ws(k), k) >= 2)
        End If
    Next k

    Dim hm As Integer, i As Integer, idx As Integer
    Dim lst() As Integer, nL As Integer
    Dim expNames() As String, nExp As Integer
    Dim target As Worksheet, physBlk As Integer

    For hm = 1 To gHMCount
        SortedSheetsForHm hm, ws, lst, nL
        If nL = 0 Then GoTo NextHm

        ReDim expNames(1 To HM_SHEET_COUNT)
        nExp = 0
        For i = 1 To nL
            idx = lst(i)
            ' sheet goc da dung trong cung 1 lan xuat (GHEP) -> ban sao tam
            If ghep And used(idx) Then
                Set target = CopySheetTemp(wb, ws(idx))
                nTmp = nTmp + 1: tmpNames(nTmp) = target.Name
            Else
                Set target = ws(idx)
            End If
            used(idx) = True
            target.Cells.EntireRow.Hidden = False

            If isSplit(idx) Then
                physBlk = PhysBlockForHm(target, idx, hm)
                If physBlk = 0 Then
                    warn = warn & " [" & Sheet2Name(idx) & " khong co khoi HM" & hm & "]"
                    GoTo NextSheet
                End If
                gAssign2(idx) = 3
            Else
                physBlk = 1
                gAssign2(idx) = 1
            End If

            PrepareCols2 idx, target
            HideComments2 target
            If isSplit(idx) Then HideOtherBlock target, idx, physBlk
            ApplyPageSetup2 idx, target, physBlk
            target.ResetAllPageBreaks

            On Error Resume Next
            target.Activate
            Application.Calculate
            Dim dv As Integer
            For dv = 1 To 6: DoEvents: Next dv
            EnsureBorders2 target, target.PageSetup.PrintArea
            HideColsFinal2 idx, target
            Application.PrintCommunication = True
            ForceOri2 target, idx
            On Error GoTo Fail

            nExp = nExp + 1
            expNames(nExp) = target.Name
NextSheet:
        Next i
        If nExp = 0 Then GoTo NextHm

        If ghep Then
            For i = 1 To nExp: nAll = nAll + 1: allNames(nAll) = expNames(i): Next i
        Else
            Dim outP As String
            outP = folder & baseName & HmSuffix(hm) & ".pdf"
            If Len(gForcePdf(hm)) > 0 Then outP = gForcePdf(hm)
            If Not ExportSheetGroupToPdf2(wb, expNames, nExp, outP) Then GoTo Fail
            nOut = nOut + 1: outFiles(nOut) = outP
            ' tra sheet goc ve nguyen trang truoc khi dung cho hang muc sau
            For i = 1 To nL
                If Not ws(lst(i)) Is Nothing Then ws(lst(i)).Cells.EntireRow.Hidden = False
            Next i
        End If
NextHm:
    Next hm

    If ghep And nAll > 0 Then
        Dim outM As String: outM = folder & baseName & ".pdf"
        If Not ExportSheetGroupToPdf2(wb, allNames, nAll, outM) Then GoTo Fail
        nOut = 1: outFiles(1) = outM
    End If

    CleanupTmpSheets wb, tmpNames, nTmp
    UnhideAllBlocks ws
    On Error Resume Next
    For k = 1 To HM_SHEET_COUNT
        If Not ws(k) Is Nothing Then ws(k).Select: Exit For
    Next k
    On Error GoTo 0
    Application.ScreenUpdating = savedSU

    If nOut = 0 Then
        gLastErr2 = "Khong co sheet nao de in." & warn
        Exit Function
    End If
    If Len(warn) > 0 Then gLastErr2 = "Bo qua:" & warn
    ProcessWorkbookNHM = True
    Exit Function
Fail:
    If Len(gLastErr2) = 0 Then gLastErr2 = "Loi " & Err.Number & ": " & Err.Description
    On Error Resume Next
    CleanupTmpSheets wb, tmpNames, nTmp
    UnhideAllBlocks ws
    Application.ScreenUpdating = savedSU
    On Error GoTo 0
    ProcessWorkbookNHM = False
End Function

' Danh sach idx sheet duoc chon cho hang muc hm, sap theo gOrderN (on dinh theo idx)
Private Sub SortedSheetsForHm(hm As Integer, ws() As Worksheet, ByRef lst() As Integer, ByRef n As Integer)
    Dim k As Integer, a As Integer, b As Integer, t As Integer
    ReDim lst(1 To HM_SHEET_COUNT)
    n = 0
    For k = 1 To HM_SHEET_COUNT
        If gSelN(k, hm) And Not ws(k) Is Nothing Then n = n + 1: lst(n) = k
    Next k
    For a = 2 To n
        t = lst(a): b = a - 1
        Do While b >= 1
            If OrdN(lst(b)) > OrdN(t) Then lst(b + 1) = lst(b): b = b - 1 Else Exit Do
        Loop
        lst(b + 1) = t
    Next a
End Sub

Private Function OrdN(idx As Integer) As Integer
    OrdN = gOrderN(idx)
    If OrdN < 1 Then OrdN = 99
End Function

Private Function HmSuffix(hm As Integer) As String
    If gHM_TenTheoHM Then
        HmSuffix = "-" & CleanName2(gHMNames(hm), "HM" & hm)
    Else
        HmSuffix = "_HM" & hm
    End If
End Function

Private Function BlockCount(ws As Worksheet, idx As Integer) As Integer
    Dim t() As Long, b() As Long, kd() As String, h() As Long, n As Integer
    ScanBlocks ws, idx, t, b, kd, h, n
    BlockCount = n
End Function

' Khoi vat ly (1..nBlk) cua hang muc hm tren sheet ws; 0 = sheet khong co khoi nay.
'   1) khop TEN hang muc o dong tieu de khoi
'   2) khop loai TB / XD (file 2 hang muc kieu cu)
'   3) so khoi bang so hang muc -> theo thu tu
Private Function PhysBlockForHm(ws As Worksheet, idx As Integer, hm As Integer) As Integer
    Dim bkTop() As Long, bkBottom() As Long, bkKind() As String, bkHmRow() As Long, nBlk As Integer
    ScanBlocks ws, idx, bkTop, bkBottom, bkKind, bkHmRow, nBlk
    If nBlk = 0 Then PhysBlockForHm = 0: Exit Function

    Dim want As String: want = NormHmName(gHMNames(hm))
    Dim i As Integer, nm As String, nHit As Integer, iHit As Integer
    If Len(want) > 0 Then
        For i = 1 To nBlk                       ' bang tuyet doi
            If BlockHmName(ws, bkHmRow(i)) = want Then PhysBlockForHm = i: Exit Function
        Next i
        For i = 1 To nBlk                       ' chua ten (dong co them chu khac) - chi nhan neu duy nhat
            nm = BlockHmName(ws, bkHmRow(i))
            If Len(nm) > 0 Then
                If InStr(nm, want) > 0 Or InStr(want, nm) > 0 Then nHit = nHit + 1: iHit = i
            End If
        Next i
        If nHit = 1 Then PhysBlockForHm = iHit: Exit Function
    End If

    Dim wantKind As String: wantKind = gHMKinds(hm)
    If Len(wantKind) > 0 Then
        Dim hit As Integer, cnt As Integer
        For i = 1 To nBlk
            If bkKind(i) = wantKind Then cnt = cnt + 1: hit = i
        Next i
        If cnt = 1 Then PhysBlockForHm = hit: Exit Function
    End If

    If nBlk = gHMCount And hm <= nBlk Then PhysBlockForHm = hm Else PhysBlockForHm = 0
End Function

' Ten hang muc doc o dong tieu de khoi (phan sau "HANG MUC:"), da chuan hoa
Private Function BlockHmName(ws As Worksheet, r As Long) As String
    Dim c As Long, t As String, up As String, p As Long
    For c = 1 To 8
        t = SafeText2(ws, r, c)
        If Len(t) > 0 Then
            up = NormUp2(t)
            p = InStr(up, "HANG MUC")
            If p > 0 Then
                BlockHmName = NormHmName(t)
                Exit Function
            End If
        End If
    Next c
    ' Cong trinh: dong "HM" o cot E, ten o cot F tro di
    For c = 6 To 8
        t = SafeText2(ws, r, c)
        If Len(Trim$(t)) > 3 Then BlockHmName = NormHmName(t): Exit Function
    Next c
    ' Gia tong hop / Don gia CT: ten HM tran o cot A duoi dong tieu de
    t = SafeText2(ws, r, 1)
    If Len(Trim$(t)) > 3 Then BlockHmName = NormHmName(t)
End Function

' Chuan hoa ten hang muc de so sanh: bo dau, viet hoa, bo tien to "...HANG MUC:", gop khoang trang
Private Function NormHmName(s As String) As String
    Dim up As String: up = NormUp2(s)
    Dim p As Long: p = InStr(up, "HANG MUC")
    If p > 0 Then
        up = Mid$(up, p + 8)
        Do While Len(up) > 0 And (Left$(up, 1) = " " Or Left$(up, 1) = ":")
            up = Mid$(up, 2)
        Loop
    End If
    up = Replace(up, "-", " ")
    up = Replace(up, ",", " ")
    Do While InStr(up, "  ") > 0
        up = Replace(up, "  ", " ")
    Loop
    NormHmName = Trim$(up)
End Function

' Mac dinh / mau chon sheet cho moi hang muc cua file (dung khi in ca thu muc):
'   sheet tach khoi theo gTplSel ; Gia thang -> hang muc CUOI ; TH chi phi TB -> hang muc DAU
Public Sub ApplyTemplateSel()
    Dim hm As Integer, k As Integer
    For hm = 1 To MAX_HM
        For k = 1 To HM_SHEET_COUNT
            gSelN(k, hm) = False
        Next k
    Next hm
    For hm = 1 To gHMCount
        For k = 1 To HM_SHEET_COUNT
            Select Case k
                Case 5: gSelN(k, hm) = gTplSel(k) And (hm = gHMCount)
                Case 6: gSelN(k, hm) = gTplSel(k) And (hm = 1)
                Case Else: gSelN(k, hm) = gTplSel(k)
            End Select
        Next k
    Next hm
End Sub

Private Function KindFromName(nm As String) As String
    Dim up As String: up = NormUp2(nm)
    If InStr(up, "LAP DAT THIET BI") > 0 Then KindFromName = "TB": Exit Function
    If InStr(up, "THIET BI") > 0 Then KindFromName = "TB": Exit Function
    If InStr(up, "XAY DUNG") > 0 Then KindFromName = "XD": Exit Function
    KindFromName = ""
End Function

Private Sub FindHangMucHeaderRows(ws As Worksheet, markerCol As Long, _
        ByRef rows() As Long, ByRef names() As String, ByRef cnt As Integer)
    cnt = 0
    ReDim rows(1 To 10)
    ReDim names(1 To 10)

    ' End(xlUp) bo qua dong AN -> neu khoi duoi dang bi an (lan in truoc)
    ' se khong thay "HANG MUC:" thu 2. Lay them dong cuoi cua UsedRange.
    Dim lastR As Long, lastU As Long
    lastR = ws.Cells(ws.Rows.Count, markerCol).End(xlUp).Row
    On Error Resume Next
    lastU = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    On Error GoTo 0
    If lastU > lastR Then lastR = lastU
    If lastR < 1 Then lastR = 1

    Dim r As Long, txt As String, up As String
    For r = 1 To lastR
        txt = SafeText2(ws, r, markerCol)
        up = NormUp2(txt)
        ' chap nhan "HANG MUC:" hoac "HANG MUC :" (co dau tieng Viet da bo)
        If (InStr(up, "HANG MUC:") > 0 Or InStr(up, "HANG MUC :") > 0) _
           And Left$(up, 4) <> "TONG" And Left$(up, 1) <> "+" Then
            cnt = cnt + 1
            If cnt > UBound(rows) Then
                ReDim Preserve rows(1 To cnt + 5)
                ReDim Preserve names(1 To cnt + 5)
            End If
            rows(cnt) = r
            names(cnt) = Trim(Mid(txt, InStr(txt, ":") + 1))
        End If
    Next r
End Sub

Private Function GetMarkerCol2(idx As Integer) As Long
    Select Case idx
        Case 4: GetMarkerCol2 = 5     ' Cong trinh -> cot E
        Case Else: GetMarkerCol2 = 1  ' cot A
    End Select
End Function

Private Function Sheet2Name(idx As Integer) As String
    Select Case idx
        Case 1: Sheet2Name = "THKP hang muc"
        Case 2: Sheet2Name = "Gia tong hop"
        Case 3: Sheet2Name = "Don gia chi tiet"
        Case 4: Sheet2Name = "Cong trinh"
        Case 5: Sheet2Name = "Gia thang (chung)"
        Case 6: Sheet2Name = "TH chi phi TB (chung)"
    End Select
End Function

' ================================================================
' XUAT 1 NHOM SHEET (cua 1 hang muc) RA 1 FILE PDF RIENG
' ----------------------------------------------------------------
' Sap xep lai vi tri tab vat ly theo dung thu tu trong expNames()
' (ExportAsFixedFormat tren nhom sheet duoc chon luon xuat theo VI TRI
' TAB trai-sang-phai, KHONG theo thu tu trong mang/luot chon), roi
' chon nhom va xuat PDF. Dung chung cho ca hang muc 1 va hang muc 2.
' ================================================================
Private Function ExportSheetGroupToPdf2(wb As Workbook, expNames() As String, nExp As Integer, pdfPath As String) As Boolean
    On Error GoTo Fail
    ExportSheetGroupToPdf2 = False

    On Error Resume Next
    Dim mv As Integer
    For mv = 2 To nExp
        wb.Sheets(expNames(mv)).Move After:=wb.Sheets(expNames(mv - 1))
    Next mv
    On Error GoTo Fail

    On Error Resume Next
    wb.Activate
    Dim vArr() As Variant
    ReDim vArr(0 To nExp - 1)
    Dim j As Integer
    For j = 1 To nExp
        vArr(j - 1) = expNames(j)
    Next j

    Err.Clear
    wb.Sheets(vArr).Select
    If Err.Number <> 0 Then
        ' fallback: select tung cai
        Err.Clear
        wb.Sheets(expNames(1)).Select
        For j = 2 To nExp
            wb.Sheets(expNames(j)).Select Replace:=False
        Next j
    End If
    On Error GoTo Fail

    KillFile2 pdfPath
    On Error Resume Next
    wb.ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=pdfPath, Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
    Dim expErr As Long: expErr = Err.Number
    Dim expDesc As String: expDesc = Err.Description
    On Error GoTo Fail

    ' Bo group ngay: sheet con group thi lan dat PageSetup sau (hang muc ke tiep)
    ' bi Excel ap sai (vd THKP ra kho doc).
    On Error Resume Next
    wb.Sheets(expNames(1)).Select
    On Error GoTo Fail

    If expErr <> 0 Or Not FileExists2(pdfPath) Then
        gLastErr2 = "Loi xuat PDF (" & expErr & "): " & expDesc
        Exit Function
    End If

    ExportSheetGroupToPdf2 = True
    Exit Function
Fail:
    If Len(gLastErr2) = 0 Then gLastErr2 = "Loi xuat PDF " & Err.Number & ": " & Err.Description
    ExportSheetGroupToPdf2 = False
End Function

' Tao ban sao tam cua 1 sheet, dat ngay sau sheet goc, tra ve sheet moi
Private Function CopySheetTemp(wb As Workbook, src As Worksheet) As Worksheet
    src.Copy After:=src
    ' Ngay sau Copy, ban sao la sheet dang active -> lay chinh xac
    Dim cp As Worksheet
    On Error Resume Next
    Set cp = wb.ActiveSheet
    On Error GoTo 0
    If cp Is Nothing Then Set cp = wb.Sheets(src.Index + 1)
    ' Dat ten tam duy nhat de biet ma xoa (tranh trung ten)
    Dim baseNm As String, tryNm As String, n As Integer
    baseNm = Left$("TMP2HM_" & src.Name, 27)
    tryNm = baseNm
    n = 0
    Do While SheetNameExists(wb, tryNm)
        n = n + 1
        tryNm = Left$(baseNm, 27) & "_" & n
    Loop
    On Error Resume Next
    cp.Name = tryNm
    On Error GoTo 0
    Set CopySheetTemp = cp
End Function

' Kiem tra ten sheet da ton tai chua
Private Function SheetNameExists(wb As Workbook, nm As String) As Boolean
    Dim s As Worksheet
    On Error Resume Next
    For Each s In wb.Worksheets
        If StrComp(s.Name, nm, vbTextCompare) = 0 Then SheetNameExists = True: Exit Function
    Next s
    On Error GoTo 0
    SheetNameExists = False
End Function

' Xoa cac sheet ban sao tam
Private Sub CleanupTmpSheets(wb As Workbook, tmpNames() As String, nTmp As Integer)
    If nTmp < 1 Then Exit Sub
    Dim savedAlerts As Boolean: savedAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error Resume Next
    Dim i As Integer
    For i = 1 To nTmp
        If Len(tmpNames(i)) > 0 Then
            wb.Sheets(tmpNames(i)).Delete
        End If
    Next i
    On Error GoTo 0
    Application.DisplayAlerts = savedAlerts
End Sub

Private Sub ScanBlocks(ws As Worksheet, idx As Integer, _
        ByRef bkTop() As Long, ByRef bkBottom() As Long, _
        ByRef bkKind() As String, ByRef bkHmRow() As Long, ByRef nBlk As Integer)

    nBlk = 0
    ReDim bkTop(1 To 10): ReDim bkBottom(1 To 10)
    ReDim bkKind(1 To 10): ReDim bkHmRow(1 To 10)

    Dim lastR As Long: lastR = SheetLastRow2(ws, idx)

    If idx = 4 Then
        ' ---- CONG TRINH: chia theo HM / THM o cot E ----
        Dim hmRows() As Long, thmRows() As Long, nh As Integer, nt As Integer
        ReDim hmRows(1 To 20): ReDim thmRows(1 To 20)
        nh = 0: nt = 0
        Dim r As Long, e As String
        For r = 1 To lastR
            e = UCase$(Trim(SafeText2(ws, r, 5)))   ' cot E
            If e = "HM" Then
                nh = nh + 1: hmRows(nh) = r
            ElseIf e = "THM" Then
                nt = nt + 1: thmRows(nt) = r
            End If
        Next r

        ' Tim dong tieu de tong the (row dau)
        Dim firstTitle As Long: firstTitle = FindOverallTitleRow(ws, idx, 1, lastR)
        If firstTitle = 0 Then firstTitle = 1

        Dim b As Integer
        For b = 1 To nh
            nBlk = nBlk + 1
            If nBlk > UBound(bkTop) Then GrowBlockArrays bkTop, bkBottom, bkKind, bkHmRow
            ' top: khoi 1 gom tieu de tong the; khoi sau bat dau tai HM
            If b = 1 Then
                bkTop(nBlk) = firstTitle
            Else
                bkTop(nBlk) = hmRows(b)
            End If
            ' bottom: THM tuong ung (hoac truoc HM ke tiep, hoac lastR)
            If b <= nt Then
                bkBottom(nBlk) = thmRows(b)
            ElseIf b < nh Then
                bkBottom(nBlk) = hmRows(b + 1) - 1
            Else
                bkBottom(nBlk) = lastR
            End If
            bkHmRow(nBlk) = hmRows(b)
            bkKind(nBlk) = ClassifyBlockKind(ws, bkTop(nBlk), bkBottom(nBlk))
        Next b
        Exit Sub
    End If

    ' ---- THKP / GIA TONG HOP / DON GIA CT ----
    ' Tim tat ca dong "HANG MUC:" (cot A)
    Dim hmR() As Long, nHM As Integer
    ReDim hmR(1 To 20): nHM = 0
    Dim rr As Long, up As String, i As Integer
    For rr = 1 To lastR
        up = NormUp2(SafeText2(ws, rr, 1))
        If IsHangMucLine(up) Then
            nHM = nHM + 1: hmR(nHM) = rr
        End If
    Next rr
    ' Gia tong hop / Don gia CT: khong co dong "HANG MUC:" -> khoi bat dau tai dong tieu de
    ' lap lai ("GIA XAY DUNG TONG HOP..." / "DON GIA XAY DUNG CHI TIET..."); ten HM nam duoi tieu de.
    If (idx = 2 Or idx = 3) And nHM < 2 Then
        Dim tk() As String: tk = GetTitleMarkers(idx)
        Dim tt As Integer, q As Long, nmRow As Long, ua As String, hit2 As Boolean
        For rr = 1 To lastR
            up = NormUp2(SafeText2(ws, rr, 1))
            hit2 = False
            For tt = LBound(tk) To UBound(tk)
                If Len(tk(tt)) > 0 Then If InStr(up, tk(tt)) > 0 Then hit2 = True
            Next tt
            If hit2 Then
                nmRow = rr
                For q = rr + 1 To rr + 5
                    ua = NormUp2(SafeText2(ws, q, 1))
                    If Len(ua) > 2 And Left$(ua, 10) <> "CONG TRINH" And ua <> "STT" Then nmRow = q: Exit For
                Next q
                nBlk = nBlk + 1
                If nBlk > UBound(bkTop) Then GrowBlockArrays bkTop, bkBottom, bkKind, bkHmRow
                bkTop(nBlk) = rr: bkHmRow(nBlk) = nmRow
                bkKind(nBlk) = ClassifyBlockKind(ws, nmRow, nmRow)
            End If
        Next rr
        If nBlk >= 2 Then
            For i = 1 To nBlk
                If i < nBlk Then bkBottom(i) = bkTop(i + 1) - 1 Else bkBottom(i) = lastR
                If bkBottom(i) < bkTop(i) Then bkBottom(i) = bkTop(i)
            Next i
            Exit Sub
        End If
        nBlk = 0
    End If
    If nHM = 0 Then Exit Sub

    Dim topRow As Long, nextTop As Long
    For i = 1 To nHM
        nBlk = nBlk + 1
        If nBlk > UBound(bkTop) Then GrowBlockArrays bkTop, bkBottom, bkKind, bkHmRow
        ' top = tieu de tong the gan nhat phia tren dong HANG MUC (tim nguoc <=5 dong)
        topRow = FindOverallTitleRow(ws, idx, WorksheetMax(1, hmR(i) - 6), hmR(i))
        If topRow = 0 Then topRow = hmR(i)
        bkTop(nBlk) = topRow
        bkHmRow(nBlk) = hmR(i)
        bkKind(nBlk) = ClassifyBlockKind(ws, hmR(i), hmR(i))
    Next i

    ' bottom moi khoi = truoc top khoi sau; khoi cuoi = lastR
    For i = 1 To nBlk
        If i < nBlk Then
            bkBottom(i) = bkTop(i + 1) - 1
        Else
            bkBottom(i) = lastR
        End If
        If bkBottom(i) < bkTop(i) Then bkBottom(i) = bkTop(i)
    Next i
End Sub

Private Function IsHangMucLine(up As String) As Boolean
    IsHangMucLine = False
    ' Truong hop 1: co tien to "HANG MUC:" / "TEN HANG MUC:"
    If (Left$(up, 9) = "HANG MUC:") Or (Left$(up, 13) = "TEN HANG MUC:") _
       Or (Left$(up, 10) = "HANG MUC :") Or (Left$(up, 14) = "TEN HANG MUC :") Then
        IsHangMucLine = True: Exit Function
    End If

    ' Truong hop 2: dong TEN HANG MUC TRAN (khong tien to) - vd file Solar:
    '   "CAP DIEN MAT TROI - LAP DAT THIET BI" / "CAP DIEN MAT TROI - XAY DUNG"
    ' Dieu kien chat de KHONG bat nham cac dong khac:
    '   - Loai cac dong tieu de / ten du an / ten cong trinh
    '   - Phai KET THUC bang " - LAP DAT THIET BI" hoac " - XAY DUNG"
    '     (hau to loai hang muc nam o CUOI dong)
    If InStr(up, "TONG HOP") > 0 Then Exit Function
    If InStr(up, "DON GIA XAY DUNG CHI TIET") > 0 Then Exit Function
    If InStr(up, "DU TOAN CHI PHI XAY DUNG") > 0 Then Exit Function
    If Left$(up, 14) = "TEN CONG TRINH" Then Exit Function   ' "Ten cong trinh: XAY DUNG NHA..."
    If Left$(up, 15) = "TEN CONG TRINH:" Then Exit Function
    If Left$(up, 10) = "TEN DU AN" Then Exit Function
    If Left$(up, 11) = "CONG TRINH:" Then Exit Function

    ' Phai ket thuc bang hau to loai hang muc
    If RightMatch(up, " - LAP DAT THIET BI") Then IsHangMucLine = True: Exit Function
    If RightMatch(up, "- LAP DAT THIET BI") Then IsHangMucLine = True: Exit Function
    If RightMatch(up, " - XAY DUNG") Then IsHangMucLine = True: Exit Function
    If RightMatch(up, "- XAY DUNG") Then IsHangMucLine = True: Exit Function
    IsHangMucLine = False
End Function

' Kiem tra chuoi 's' co KET THUC bang 'suffix' khong (da chuan hoa hoa)
Private Function RightMatch(s As String, suffix As String) As Boolean
    If Len(s) < Len(suffix) Then RightMatch = False: Exit Function
    RightMatch = (Right$(s, Len(suffix)) = suffix)
End Function

Private Function ClassifyBlockKind(ws As Worksheet, r1 As Long, r2 As Long) As String
    Dim r As Long, c As Long, up As String
    ' Uu tien tuyet doi: dong DAU khoi (r1) - phan loai theo hau to loai hang muc
    For c = 1 To 6
        up = NormUp2(SafeText2(ws, r1, c))
        If Len(up) > 0 Then
            If InStr(up, "LAP DAT THIET BI") > 0 Then ClassifyBlockKind = "TB": Exit Function
            If RightMatch(up, "XAY DUNG") And up <> "" And Left$(up, 14) <> "TEN CONG TRINH" _
               And InStr(up, "TONG HOP") = 0 And InStr(up, "CHI TIET") = 0 Then ClassifyBlockKind = "XD": Exit Function
        End If
    Next c
    ' Quet ca vung: uu tien dong co "HANG MUC" (co tien to)
    For r = r1 To r2
        For c = 1 To 6
            up = NormUp2(SafeText2(ws, r, c))
            If InStr(up, "HANG MUC") > 0 Then
                If InStr(up, "LAP DAT THIET BI") > 0 Then ClassifyBlockKind = "TB": Exit Function
                If InStr(up, "THIET BI") > 0 Then ClassifyBlockKind = "TB": Exit Function
                If InStr(up, "XAY DUNG") > 0 Then ClassifyBlockKind = "XD": Exit Function
            End If
        Next c
    Next r
    ' Du phong cuoi: dong tran chua hau to
    For r = r1 To r2
        For c = 1 To 6
            up = NormUp2(SafeText2(ws, r, c))
            If InStr(up, "TONG HOP") = 0 And InStr(up, "DON GIA XAY DUNG CHI TIET") = 0 _
               And InStr(up, "DU TOAN CHI PHI") = 0 And Left$(up, 14) <> "TEN CONG TRINH" Then
                If RightMatch(up, "LAP DAT THIET BI") Then ClassifyBlockKind = "TB": Exit Function
                If RightMatch(up, "- XAY DUNG") Then ClassifyBlockKind = "XD": Exit Function
            End If
        Next c
    Next r
    ClassifyBlockKind = ""
End Function

Private Function FindOverallTitleRow(ws As Worksheet, idx As Integer, r1 As Long, r2 As Long) As Long
    Dim titles() As String: titles = GetTitleMarkers(idx)
    Dim r As Long, t As Integer, up As String
    For r = r1 To r2
        up = NormUp2(SafeText2(ws, r, 1))
        For t = LBound(titles) To UBound(titles)
            If Len(titles(t)) > 0 Then
                If InStr(up, titles(t)) > 0 Then FindOverallTitleRow = r: Exit Function
            End If
        Next t
    Next r
    FindOverallTitleRow = 0
End Function

Private Function SheetLastRow2(ws As Worksheet, idx As Integer) As Long
    Dim lastR As Long, lr2 As Long
    lastR = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lr2 = ws.Cells(ws.Rows.Count, GetMarkerCol2Letter(idx)).End(xlUp).Row
    If lr2 > lastR Then lastR = lr2
    ' Cong trinh: du lieu o cot E/F -> kiem them
    If idx = 4 Then
        Dim lrF As Long: lrF = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
        If lrF > lastR Then lastR = lrF
        Dim lrE As Long: lrE = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
        If lrE > lastR Then lastR = lrE
    End If
    If lastR < 1 Then lastR = 1
    SheetLastRow2 = lastR
End Function

Private Function WorksheetMax(a As Long, b As Long) As Long
    WorksheetMax = IIf(a > b, a, b)
End Function

Private Sub GrowBlockArrays(ByRef t() As Long, ByRef b() As Long, _
        ByRef k() As String, ByRef h() As Long)
    Dim n As Long: n = UBound(t) + 10
    ReDim Preserve t(1 To n): ReDim Preserve b(1 To n)
    ReDim Preserve k(1 To n): ReDim Preserve h(1 To n)
End Sub

Private Sub HideOtherBlock(ws As Worksheet, idx As Integer, hm As Integer)
    On Error Resume Next

    ' QUAN TRONG: unhide TOAN BO dong TRUOC khi do khoi.
    ' Neu con dong an tu lan in hang muc truoc, ScanBlocks / End(xlUp)
    ' co the do sai ranh gioi khoi -> chon nham khoi (bug "HM2 ra khoi TB").
    ws.Cells.EntireRow.Hidden = False

    Dim bkTop() As Long, bkBottom() As Long, bkKind() As String, bkHmRow() As Long, nBlk As Integer
    ScanBlocks ws, idx, bkTop, bkBottom, bkKind, bkHmRow, nBlk
    If nBlk < 2 Then Exit Sub   ' khong phai 2 khoi -> khong an gi

    ' An moi khoi khac 'hm'
    Dim i As Integer
    For i = 1 To nBlk
        If i <> hm Then
            If bkBottom(i) >= bkTop(i) Then
                ws.Rows(bkTop(i) & ":" & bkBottom(i)).Hidden = True
            End If
        End If
    Next i

    On Error GoTo 0
End Sub

Private Function GetMarkerCol2Letter(idx As Integer) As String
    Select Case idx
        Case 4: GetMarkerCol2Letter = "E"
        Case Else: GetMarkerCol2Letter = "A"
    End Select
End Function

Private Function GetTitleMarkers(idx As Integer) As String()
    Dim m() As String
    Select Case idx
        Case 1 ' THKP
            m = Split("BANG TONG HOP DU TOAN|TONG HOP DU TOAN CHI PHI", "|")
        Case 2 ' Gia tong hop
            m = Split("GIA XAY DUNG TONG HOP", "|")
        Case 3 ' Don gia chi tiet
            m = Split("DON GIA XAY DUNG CHI TIET", "|")
        Case 4 ' Cong trinh
            m = Split("BANG DU TOAN HANG MUC|DU TOAN HANG MUC CONG TRINH", "|")
        Case Else
            m = Split("", "|")
    End Select
    GetTitleMarkers = m
End Function

Private Sub UnhideAllBlocks(ws() As Worksheet)
    On Error Resume Next
    Dim k As Integer
    For k = 1 To HM_SHEET_COUNT
        If Not ws(k) Is Nothing Then
            ws(k).Cells.EntireRow.Hidden = False
        End If
    Next k
    On Error GoTo 0
End Sub

Private Sub ApplyPageSetup2(idx As Integer, ws As Worksheet, hm As Integer)
    Dim paTop As Long, paBottom As Long, titleRow As Long, titleRow2 As Long
    Dim endCol As String
    Dim titleRowsStr As String   ' luu gia tri PrintTitleRows, se gan LAI o ngoai (xem ForcePrintTitleRows)

    ' Xac dinh vung in cho dung khoi
    GetBlockRange ws, idx, hm, paTop, paBottom, titleRow, titleRow2

    ' FIX TIEU DE MAT 1 DONG: cac sheet co header 2 dong kieu merge "hon
    ' hop" (da so cot merge xuyen 2 dong, rieng 1-2 nhom cot chi merge
    ' ngang o dong tren, dong duoi tach rieng cac o con) deu bi Excel
    ' "quen" mat dong tach khi lap lai tieu de (Print Titles) sang trang
    ' 2 tro di. Xoa merge roi merge lai chinh xac TRUOC khi gan PrintArea/
    ' PrintTitleRows, giong het cach da fix cho case 2/4/8 trong
    ' PrepareSheetForPrint (nhanh 1 hang muc). Ap dung cho ca 3 sheet co
    ' kieu header nay trong nhanh 2 hang muc (truoc day chi fix idx=6,
    ' bo sot idx=2 va idx=4 dung loi tuong tu):
    '   idx=2 Gia tong hop : Don gia (F:H) / Thanh tien (I:K)
    '   idx=4 Cong trinh   : Kich thuoc (I:K)
    '   idx=6 TH chi phi TB: Thue VAT (G:H)
    If titleRow > 0 And titleRow2 > 0 Then
        Select Case idx
            Case 2
                RefreshTwoRowHeaderMerge ws, titleRow, titleRow2, 1, 12, Array(6, 8, 9, 11)
            Case 4
                RefreshTwoRowHeaderMerge ws, titleRow, titleRow2, 1, 16, Array(9, 11)
            Case 6
                RefreshTwoRowHeaderMerge ws, titleRow, titleRow2, 1, 10, Array(7, 8)
        End Select
    End If

    ' Cot cuoi vung in theo tung loai sheet
    Select Case idx
        Case 1: endCol = "H"                                   ' THKP
        Case 2: endCol = IIf(gGTH_InTongCong2, "L", "K")       ' Gia tong hop
        Case 3: endCol = "H"                                   ' Don gia chi tiet
        Case 4: endCol = "P"                                   ' Cong trinh
        Case 5: endCol = "K"                                   ' Gia thang (chung)
        Case 6: endCol = "J"                                   ' TH chi phi TB (chung)
        Case Else: endCol = "H"
    End Select

    Application.PrintCommunication = False
    With ws.PageSetup
        .PaperSize          = xlPaperA4
        .Zoom               = False
        .FitToPagesWide     = 1
        .FitToPagesTall     = IIf(idx = 1, 1, False)   ' THKP: bang ngan -> gon 1 trang
        .LeftMargin         = Application.CentimetersToPoints(2.5)
        .RightMargin        = Application.CentimetersToPoints(1)
        .TopMargin          = Application.CentimetersToPoints(2.5)
        .BottomMargin       = Application.CentimetersToPoints(1.5)
        .HeaderMargin       = Application.CentimetersToPoints(1)
        .FooterMargin       = Application.CentimetersToPoints(1)
        .BlackAndWhite      = True
        .PrintQuality       = 600
        .CenterHorizontally = True
        .CenterVertically   = False
        .PrintGridlines     = False
        .PrintHeadings      = False
        .Draft              = False
        .Order              = xlDownThenOver
        .FirstPageNumber    = xlAutomatic
        .PrintComments      = xlPrintNoComments
        .PrintErrors        = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter    = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter       = True
        .AlignMarginsHeaderFooter       = True

        .PrintArea = "$A$" & paTop & ":$" & endCol & "$" & paBottom

        ' PrintTitleRows: lap dong header cot cua khoi dang in
        ' (chi TINH gia tri o day; gan THUC SU se lam o NGOAI khoi
        '  PrintCommunication=False qua ForcePrintTitleRows, tranh bi
        '  Excel "nuot lenh" lam mat 1 dong tieu de khi sang trang 2)
        If titleRow > 0 Then
            If titleRow2 > 0 And titleRow2 >= titleRow Then
                titleRowsStr = "$" & titleRow & ":$" & titleRow2
            Else
                titleRowsStr = "$" & titleRow & ":$" & titleRow
            End If
        Else
            titleRowsStr = ""
        End If

        ' Huong in
        Select Case idx
            Case 1, 2, 6: .Orientation = xlLandscape   ' THKP, Gia tong hop, TH TB
            Case Else:    .Orientation = xlPortrait
        End Select
    End With
    Application.PrintCommunication = True

    ' Footer dat SAU khi PrintCommunication=True (khong bi nuot)
    On Error Resume Next
    With ws.PageSetup
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .CenterHeader = "": .LeftHeader = "": .RightHeader = ""
        .CenterFooter = "": .LeftFooter = "": .RightFooter = ""
        .EvenPage.LeftHeader.Text = "": .EvenPage.CenterHeader.Text = "": .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = "": .EvenPage.CenterFooter.Text = "": .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = "": .FirstPage.CenterHeader.Text = "": .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = "": .FirstPage.CenterFooter.Text = "": .FirstPage.RightFooter.Text = ""
    End With
    On Error GoTo 0

    ForceOri2 ws, idx

    ' V4 FIX TIEU DE MAT 1 DONG O TRANG 2: gan lai PrintTitleRows o ngoai
    ' khoi PrintCommunication=False (xem ForcePrintTitleRows).
    ForcePrintTitleRows ws, titleRowsStr

    AutoFitNum2 ws, idx
    HideColsFinal2 idx, ws

    ' TH chi phi TB (idx6): an moi dong TRONG HOAN TOAN trong khoi dang in,
    ' dong bo voi Case 8 cua nhanh 1 hang muc (xem HideFullyBlankRows).
    If idx = 6 Then
        Dim dataStart6 As Long
        dataStart6 = IIf(titleRow2 > 0, titleRow2, titleRow) + 1
        HideFullyBlankRows ws, dataStart6, paBottom, 1, 10
    End If
End Sub

Private Sub GetBlockRange(ws As Worksheet, idx As Integer, hm As Integer, _
        ByRef paTop As Long, ByRef paBottom As Long, _
        ByRef titleRow As Long, ByRef titleRow2 As Long)

    Dim lastR As Long: lastR = SheetLastRow2(ws, idx)
    titleRow = 0: titleRow2 = 0

    ' Sheet chung (khong tach) -> in toan bo
    If gAssign2(idx) <> 3 Then
        paTop = 1
        paBottom = lastR
        FindHeaderRowInRange ws, idx, 1, lastR, titleRow, titleRow2
        Exit Sub
    End If

    ' Sheet tach 2 khoi -> dung ScanBlocks
    Dim bkTop() As Long, bkBottom() As Long, bkKind() As String, bkHmRow() As Long, nBlk As Integer
    ScanBlocks ws, idx, bkTop, bkBottom, bkKind, bkHmRow, nBlk

    If nBlk < hm Or nBlk < 1 Then
        paTop = 1: paBottom = lastR
        FindHeaderRowInRange ws, idx, 1, lastR, titleRow, titleRow2
        Exit Sub
    End If

    paTop = bkTop(hm)
    paBottom = bkBottom(hm)
    If paBottom < paTop Then paBottom = paTop

    FindHeaderRowInRange ws, idx, paTop, paBottom, titleRow, titleRow2
End Sub

Private Sub FindHeaderRowInRange(ws As Worksheet, idx As Integer, r1 As Long, r2 As Long, _
        ByRef titleRow As Long, ByRef titleRow2 As Long)
    titleRow = 0: titleRow2 = 0

    ' THKP (idx 1): bang ngan 1 trang -> KHONG lap dong header (giong macro mau)
    If idx = 1 Then Exit Sub

    Dim r As Long, a As String
    For r = r1 To r2
        a = UCase$(Trim(SafeText2(ws, r, 1)))
        If a = "STT" Then
            titleRow = r
            Select Case idx
                Case 2 ' Gia tong hop: header LUON 2 dong (STT o dong tren, "Vat lieu.." dong duoi)
                    titleRow2 = r + 1
                Case 4 ' Cong trinh: header 2 dong (STT + dong phu ben duoi)
                    Dim below As String
                    below = Trim(SafeText2(ws, r + 1, 1))
                    If Len(below) = 0 Then titleRow2 = r + 1
                Case 6 ' TH chi phi TB (TH thiet bi): header LUON 2 dong (dong duoi la "%"/"Thanh tien"
                       ' cua nhom "Thue VAT" G:H). FIX: truoc day thieu Case nay nen titleRow2 luon
                       ' = 0 -> PrintTitleRows chi lap 1 dong, lam mat dong "%"/"Thanh tien" khi
                       ' tieu de lap lai sang trang 2 (dung loi nguoi dung bao cao trong file PDF).
                    titleRow2 = r + 1
            End Select
            Exit Sub
        End If
    Next r

    ' Cong trinh khoi 2 (idx 4, hm 2): khong co dong "STT" trong khoi -> khong lap header
    ' (titleRow van = 0, hop le)
End Sub

Private Sub PrepareCols2(idx As Integer, ws As Worksheet)
    On Error Resume Next
    ws.Cells.ClearComments
    Select Case idx
    Case 1 ' THKP: an cot thua E,F,G (moi khoi co dong 2,5 la tieu de con -> de nguyen,
           ' vi khi tach khoi cac dong ngoai vung in se khong in)
        ws.Columns("E:G").Hidden = True
    Case 2 ' Gia tong hop
        ws.Columns("L:L").ColumnWidth = 16
        ws.Columns("L:L").Hidden = Not gGTH_InTongCong2
    Case 3 ' Don gia chi tiet: khong an gi
    Case 4 ' Cong trinh: whitelist hien A, E, F:L, O:P
        ws.Columns("A:AM").Hidden = True
        ws.Range("A:A").EntireColumn.Hidden = False
        ws.Range("E:E").EntireColumn.Hidden = False
        ws.Range("F:L").EntireColumn.Hidden = False
        ws.Range("O:P").EntireColumn.Hidden = False
    Case 5 ' Gia thang: an E, G:J
        ws.Columns("E:E").Hidden = True
        ws.Columns("G:J").Hidden = True
    Case 6 ' TH chi phi TB: khong an gi
    End Select
    On Error GoTo 0
End Sub

Private Sub HideColsFinal2(idx As Integer, ws As Worksheet)
    On Error Resume Next
    Select Case idx
    Case 1: ws.Columns("E:G").Hidden = True
    Case 2: ws.Columns("L:L").Hidden = Not gGTH_InTongCong2
    Case 4
        ws.Columns("A:AM").Hidden = True
        ws.Range("A:A").EntireColumn.Hidden = False
        ws.Range("E:E").EntireColumn.Hidden = False
        ws.Range("F:L").EntireColumn.Hidden = False
        ws.Range("O:P").EntireColumn.Hidden = False
    Case 5
        ws.Columns("E:E").Hidden = True
        ws.Columns("G:J").Hidden = True
    End Select
    On Error GoTo 0
End Sub

Private Sub AutoFitNum2(ws As Worksheet, idx As Integer)
    On Error Resume Next
    Select Case idx
    Case 1: ws.Columns("H:H").AutoFit
    Case 2
        ws.Columns("I:K").AutoFit
        If gGTH_InTongCong2 Then ws.Columns("L:L").ColumnWidth = 16
    Case 3: ws.Columns("F:H").AutoFit
    Case 4: ws.Columns("K:L").AutoFit
    Case 5: ws.Columns("F:F").AutoFit: ws.Columns("K:K").AutoFit
    Case 6: ws.Columns("D:J").AutoFit
    End Select
    On Error GoTo 0
End Sub

Private Sub ForceOri2(ws As Worksheet, idx As Integer)
    Dim ori As XlPageOrientation
    Select Case idx
        Case 1, 2, 6: ori = xlLandscape
        Case Else: ori = xlPortrait
    End Select
    Application.PrintCommunication = True
    Dim attempt As Integer
    For attempt = 1 To 3
        On Error Resume Next
        ws.PageSetup.PaperSize = xlPaperA4
        ws.PageSetup.Orientation = ori
        On Error GoTo 0
        DoEvents
        On Error Resume Next
        If ws.PageSetup.Orientation = ori Then On Error GoTo 0: Exit For
        On Error GoTo 0
    Next attempt
End Sub

Private Sub EnsureBorders2(ws As Worksheet, printRangeAddr As String)
    On Error Resume Next
    If printRangeAddr = "" Then Exit Sub
    Dim rng As Range
    Set rng = ws.Range(printRangeAddr)
    If rng Is Nothing Then Exit Sub

    Dim firstRow As Long, lastRow As Long, firstCol As Long, lastCol As Long
    firstRow = rng.Row
    lastRow = rng.Row + rng.Rows.Count - 1
    firstCol = rng.Column
    lastCol = rng.Column + rng.Columns.Count - 1

    Dim r As Long, c As Long, bdr As Border
    For r = firstRow To lastRow
        If ws.Rows(r).Hidden Then GoTo NextR
        For c = firstCol To lastCol
            Set bdr = ws.Cells(r, c).Borders(xlEdgeBottom)
            If bdr.LineStyle <> xlNone Then
                If bdr.Weight = xlHairline Then
                    bdr.LineStyle = xlContinuous
                    bdr.Weight = xlThin
                End If
            End If
        Next c
NextR:
    Next r

    ' Dam bao dong cuoi (khong an) co border bottom
    Dim lastVisible As Long: lastVisible = lastRow
    Do While lastVisible > firstRow
        If Not ws.Rows(lastVisible).Hidden Then Exit Do
        lastVisible = lastVisible - 1
    Loop
    For c = firstCol To lastCol
        With ws.Cells(lastVisible, c).Borders(xlEdgeBottom)
            If .LineStyle = xlNone Then
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End If
        End With
    Next c
    On Error GoTo 0
End Sub

Private Sub HideComments2(ws As Worksheet)
    On Error Resume Next
    Dim cm As Comment
    For Each cm In ws.Comments
        cm.Visible = False
    Next cm
    On Error GoTo 0
End Sub

Private Function Find2Sheet(wb As Workbook, idx As Integer) As Worksheet
    Dim s As Worksheet, best As Worksheet
    Dim bestSc As Long, sc As Long
    bestSc = 0: Set best = Nothing
    For Each s In wb.Worksheets
        ' sheet an = khong dung, tru Gia thang (idx 5) hay bi an trong file du toan -> hien lai de in
        If s.Visible = xlSheetVisible Or idx = 5 Then sc = Score2(s, idx) Else sc = 0
        If sc > bestSc Then bestSc = sc: Set best = s
    Next s
    If bestSc >= 50 Then
        Set Find2Sheet = best
        If idx = 5 Then
            If best.Visible <> xlSheetVisible Then
                On Error Resume Next
                best.Visible = xlSheetVisible
                On Error GoTo 0
            End If
        End If
    Else
        Set Find2Sheet = Nothing
    End If
End Function

Private Function Score2(ws As Worksheet, idx As Integer) As Long
    Dim sc As Long: sc = 0
    Dim nm As String: nm = NormUp2(ws.Name)

    ' marker tieu de (+100)
    If HeaderHasMarker2(ws, idx) Then sc = sc + 100

    ' ten tab
    Dim tabs() As String: tabs = GetTabNames2(idx)
    Dim i As Long, exact As Boolean, contains As Boolean
    For i = LBound(tabs) To UBound(tabs)
        If nm = tabs(i) Then exact = True: Exit For
    Next i
    If exact Then sc = sc + 50
    For i = LBound(tabs) To UBound(tabs)
        If Len(tabs(i)) > 0 Then
            If InStr(nm, tabs(i)) > 0 Then contains = True: Exit For
        End If
    Next i
    If contains Then sc = sc + 10

    ' phat suffix phu
    Dim padded As String: padded = " " & nm & " "
    Dim suf() As String
    suf = Split(" TL | QT | THAU | THAM | THAM DINH | GOI THAU ", "|")
    For i = LBound(suf) To UBound(suf)
        If InStr(padded, suf(i)) > 0 Then sc = sc - 30: Exit For
    Next i

    If sc < 0 Then sc = 0
    Score2 = sc
End Function

Private Function HeaderHasMarker2(ws As Worksheet, idx As Integer) As Boolean
    HeaderHasMarker2 = False
    Dim mk() As String: mk = GetHdrMarker2(idx)
    Dim r As Long, c As Long, t As String, up As String, i As Long
    For r = 1 To 5
        For c = 1 To 6
            t = SafeText2(ws, r, c)
            If Len(t) >= 10 Then
                up = NormUp2(t)
                For i = LBound(mk) To UBound(mk)
                    If Len(mk(i)) > 0 Then
                        If InStr(up, mk(i)) > 0 Then HeaderHasMarker2 = True: Exit Function
                    End If
                Next i
            End If
        Next c
    Next r
End Function

Private Function GetHdrMarker2(idx As Integer) As String()
    Dim m() As String
    Select Case idx
    Case 1: m = Split("BANG TONG HOP DU TOAN CHI PHI XAY DUNG|TONG HOP DU TOAN CHI PHI", "|")
    Case 2: m = Split("GIA XAY DUNG TONG HOP HANG MUC|GIA XAY DUNG TONG HOP", "|")
    Case 3: m = Split("DON GIA XAY DUNG CHI TIET HANG MUC|DON GIA XAY DUNG CHI TIET", "|")
    Case 4: m = Split("BANG DU TOAN HANG MUC CONG TRINH|DU TOAN HANG MUC CONG TRINH", "|")
    Case 5: m = Split("BANG GIA THONG BAO CONG TRINH|GIA THONG BAO CONG TRINH|BANG GIA THONG BAO", "|")
    Case 6: m = Split("BANG TONG HOP CHI PHI THIET BI CONG TRINH|TONG HOP CHI PHI THIET BI", "|")
    Case Else: m = Split("", "|")
    End Select
    GetHdrMarker2 = m
End Function

Private Function GetTabNames2(idx As Integer) As String()
    Dim p() As String
    Select Case idx
    Case 1: p = Split("THKP HANG MUC|THKP|THKPHANGMUC", "|")
    Case 2: p = Split("GIA TONG HOP|GIATONGHOP|GTH|BANG GIA TONG HOP", "|")
    Case 3: p = Split("DON GIA CHI TIET|DGCT|DONGIACHITIET|DON GIA XAY DUNG CHI TIET", "|")
    Case 4: p = Split("CONG TRINH|CONGTRINH|CT", "|")
    Case 5: p = Split("GIA THANG|GIATHANG|GT|BANG GIA THANG", "|")
    Case 6: p = Split("TH CHI PHI TB|THCPTB|TH CP TB|TONG HOP CHI PHI TB", "|")
    Case Else: p = Split("", "|")
    End Select
    GetTabNames2 = p
End Function

Private Function NormUp2(txt As String) As String
    Static inited As Boolean
    Static mapUp(65535) As Integer   ' 0 = khong doi; nguoc lai = ma ASCII thay the
    If Not inited Then InitVnMap mapUp: inited = True

    Dim s As String: s = UCase$(txt)   ' UCase truoc de gom hoa/thuong
    Dim res As String: res = ""
    Dim j As Long, code As Long, rep As Integer
    For j = 1 To Len(s)
        code = AscW(Mid(s, j, 1)) And &HFFFF&
        rep = mapUp(code)
        If rep <> 0 Then
            res = res & Chr(rep)
        Else
            res = res & Mid(s, j, 1)
        End If
    Next j
    ' gom khoang trang
    Do While InStr(res, "  ") > 0
        res = Replace(res, "  ", " ")
    Loop
    NormUp2 = Trim(res)
End Function

Private Sub InitVnMap(ByRef m() As Integer)
    ' A: cac bien the A HOA co dau
    SetGroup m, "A", Array( _
        &HC0, &HC1, &HC2, &HC3, &H102, &H1EA0, &H1EA2, &H1EA4, &H1EA6, &H1EA8, _
        &H1EAA, &H1EAC, &H1EAE, &H1EB0, &H1EB2, &H1EB4, &H1EB6)
    ' E HOA
    SetGroup m, "E", Array( _
        &HC8, &HC9, &HCA, &H1EB8, &H1EBA, &H1EBC, &H1EBE, &H1EC0, &H1EC2, &H1EC4, &H1EC6)
    ' I HOA
    SetGroup m, "I", Array(&HCC, &HCD, &H128, &H1EC8, &H1ECA)
    ' O HOA
    SetGroup m, "O", Array( _
        &HD2, &HD3, &HD4, &HD5, &H1A0, &H1ECC, &H1ECE, &H1ED0, &H1ED2, &H1ED4, _
        &H1ED6, &H1ED8, &H1EDA, &H1EDC, &H1EDE, &H1EE0, &H1EE2)
    ' U HOA
    SetGroup m, "U", Array( _
        &HD9, &HDA, &H168, &H1AF, &H1EE4, &H1EE6, &H1EE8, &H1EEA, &H1EEC, &H1EEE, &H1EF0)
    ' Y HOA
    SetGroup m, "Y", Array(&HDD, &H1EF2, &H1EF4, &H1EF6, &H1EF8)
    ' D gach (D HOA)
    m(&H110) = Asc("D")
End Sub

Private Sub SetGroup(ByRef m() As Integer, repChar As String, codes As Variant)
    Dim i As Long, a As Integer
    a = Asc(repChar)
    For i = LBound(codes) To UBound(codes)
        m(CLng(codes(i)) And &HFFFF&) = a
    Next i
End Sub

Private Function SafeText2(ws As Worksheet, r As Long, c As Long) As String
    On Error Resume Next
    Dim v As Variant
    v = ws.Cells(r, c).Value
    If IsError(v) Then
        SafeText2 = ""
    ElseIf IsEmpty(v) Then
        SafeText2 = ""
    Else
        SafeText2 = CStr(v)
    End If
    On Error GoTo 0
End Function

' FSO (Unicode) thay cho Dir()/Kill: duong dan co dau tieng Viet khong bi bao sai
Private Function FileExists2(p As String) As Boolean
    If Len(p) = 0 Then Exit Function
    On Error Resume Next
    FileExists2 = CreateObject("Scripting.FileSystemObject").FileExists(p)
    On Error GoTo 0
End Function

Private Sub KillFile2(p As String)
    If Len(p) = 0 Then Exit Sub
    On Error Resume Next
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(p) Then fso.DeleteFile p, True
    On Error GoTo 0
End Sub

Private Function GetFileName2(p As String) As String
    Dim i As Long
    i = InStrRev(p, "\")
    If i = 0 Then i = InStrRev(p, "/")
    If i > 0 Then GetFileName2 = Mid(p, i + 1) Else GetFileName2 = p
End Function

Private Function GetFolder2(p As String) As String
    Dim i As Long
    i = InStrRev(p, "\")
    If i = 0 Then i = InStrRev(p, "/")
    If i > 0 Then GetFolder2 = Left(p, i) Else GetFolder2 = ""
End Function

Private Function CleanName2(nm As String, fallback As String) As String
    Dim s As String: s = Trim(nm)
    If Len(s) = 0 Then CleanName2 = fallback: Exit Function
    Dim bad As Variant, ch As Variant
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|", Chr(10), Chr(13), Chr(9))
    For Each ch In bad
        s = Replace(s, CStr(ch), " ")
    Next ch
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    s = Trim(s)
    If Len(s) > 60 Then s = Left$(s, 60)
    If Len(s) = 0 Then s = fallback
    CleanName2 = s
End Function

' ================================================================
' ===================  API CHO FORM (V6)  ========================
' ================================================================
' Form frmInDuToan goi cac ham duoi day. Moi ham tra True/False va
' tra thong bao qua ByRef msg (KHONG MsgBox) de form ghi vao o nhat ky.
' ================================================================

' Nhan dien 1 file dang mo.
'   nHM        : so hang muc (>= 2 = file nhieu hang muc; 0 = 1 hang muc)
'   names()    : ten hang muc 1..nHM
'   ex1(1..8)  : sheet nao co (he 8 sheet - 1 hang muc)
'   ex2(1..6)  : sheet nao co (he 6 sheet - nhieu hang muc)
'   Tra ve so sheet tim thay theo he tuong ung (0 = khong phai file du toan)
Public Function API_Detect(wb As Workbook, ByRef nHM As Integer, ByRef names() As String, _
        ByRef ex1() As Boolean, ByRef ex2() As Boolean) As Integer
    Dim k As Integer, n1 As Integer, n2 As Integer
    ReDim ex1(1 To SHEET_COUNT)
    ReDim ex2(1 To HM_SHEET_COUNT)
    ReDim names(1 To MAX_HM)
    On Error Resume Next
    For k = 1 To SHEET_COUNT
        ex1(k) = Not (FindSheetByMarkerOrName(wb, k) Is Nothing)
        If ex1(k) Then n1 = n1 + 1
    Next k
    For k = 1 To HM_SHEET_COUNT
        ex2(k) = Not (Find2Sheet(wb, k) Is Nothing)
        If ex2(k) Then n2 = n2 + 1
    Next k
    If DetectTwoHangMuc(wb) Then nHM = gHMCount Else nHM = 0
    On Error GoTo 0
    For k = 1 To nHM: names(k) = gHMNames(k): Next k
    If nHM >= 2 Then API_Detect = n2 Else API_Detect = n1
End Function

' In / can chinh 1 file 1 HANG MUC dang mo.
' Can nap truoc: gSelSheet, gOrder, gGiaTongHopInTongCong, gOutFolder1
Public Function API_Run1HM(wb As Workbook, act As String, ByRef msg As String) As Boolean
    API_Run1HM = False
    msg = ""
    If CountSelected() = 0 Then msg = "Chua chon sheet nao.": Exit Function

    Dim pdfPath As String
    If act = "PDF" Then
        pdfPath = OutFolderFor(wb, gOutFolder1)
        If Len(pdfPath) = 0 Then msg = "File chua duoc luu - hay chon thu muc luu PDF.": Exit Function
        pdfPath = pdfPath & BaseNameOf(wb.Name) & "_FULL.pdf"
    End If

    Dim savedCalc As XlCalculation: savedCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    gLastErrDesc = ""
    Dim ok As Boolean
    ok = AlignAndExportOneWorkbook(wb, act, pdfPath)

    Application.Calculation = savedCalc
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    If Not ok Then
        msg = IIf(Len(gLastErrDesc) > 0, gLastErrDesc, "Khong tim thay sheet du toan da chon trong file.")
        Exit Function
    End If
    Select Case act
        Case "PDF":     msg = "Da xuat PDF: " & pdfPath
        Case "PRINT":   msg = "Da gui " & CountSelected() & " sheet ra may in."
        Case "PREVIEW": msg = "Da dong xem truoc."
        Case Else:      msg = "Da can chinh " & CountSelected() & " sheet (file chua luu)."
    End Select
    API_Run1HM = True
End Function

' Xuat PDF 1 file NHIEU HANG MUC dang mo.
' Can nap truoc: gSelN, gOrderN, gGTH_InTongCong2, gHM_XuatKieu,
'                gHM_TenTheoHM, gHM_OutFolder
' (Form da goi API_Detect -> gHMCount/gHMNames khop voi gSelN.)
Public Function API_RunNHM(wb As Workbook, ByRef msg As String) As Boolean
    API_RunNHM = False
    msg = ""
    If gHMCount < 2 Then msg = "File khong co nhieu khoi HANG MUC.": Exit Function
    If Len(OutFolderFor(wb, gHM_OutFolder)) = 0 Then msg = "File chua duoc luu - hay chon thu muc luu PDF.": Exit Function

    Dim savedCalc As XlCalculation: savedCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    gLastErr2 = ""
    Dim outs() As String, nOut As Integer, ok As Boolean
    ok = ProcessWorkbookNHM(wb, outs, nOut)

    Application.Calculation = savedCalc
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    If Not ok Then
        msg = IIf(Len(gLastErr2) > 0, gLastErr2, "Loi khong ro khi xuat nhieu hang muc.")
        Exit Function
    End If
    msg = "Da xuat " & nOut & " file: " & JoinOuts(outs, nOut) & "  (thu muc: " & GetFolder2(outs(1)) & ")"
    If Len(gLastErr2) > 0 Then msg = msg & vbCrLf & "  Canh bao: " & gLastErr2
    API_RunNHM = True
End Function

' Xu ly 1 file tren dia (in ca thu muc). Tu nhan dien 1HM / 2HM.
'   isPDF = True: xuat PDF ; False: in may in (file nhieu HM van xuat PDF)
'   kind tra ve: 1 = file 1 HM, 2 = file nhieu HM, 0 = khong phai file du toan
'   File nhieu HM dung mau gTplSel (xem ApplyTemplateSel).
' File dang mo san trong Excel: dung luon va KHONG dong.
Public Function API_ProcessFile(path As String, isPDF As Boolean, _
        ByRef kind As Integer, ByRef msg As String) As Boolean
    API_ProcessFile = False
    kind = 0: msg = ""
    If StrComp(path, ThisWorkbook.FullName, vbTextCompare) = 0 Then
        msg = "Bo qua (file tong)."
        Exit Function
    End If

    Dim wbk As Workbook, w As Workbook, wasOpen As Boolean
    For Each w In Application.Workbooks
        If StrComp(w.FullName, path, vbTextCompare) = 0 Then Set wbk = w: wasOpen = True: Exit For
    Next w
    If wbk Is Nothing Then
        On Error Resume Next
        Set wbk = Workbooks.Open(path, UpdateLinks:=0, ReadOnly:=True, _
                                 IgnoreReadOnlyRecommended:=True, Notify:=False, AddToMru:=False)
        On Error GoTo 0
    End If
    If wbk Is Nothing Then msg = "Khong mo duoc file.": Exit Function

    Dim savedCalc As XlCalculation: savedCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    gLastErr2 = "": gLastErrDesc = ""
    Dim ok As Boolean
    If DetectTwoHangMuc(wbk) Then
        kind = 2
        ApplyTemplateSel
        Dim outs() As String, nOut As Integer
        ok = ProcessWorkbookNHM(wbk, outs, nOut)
        If ok Then
            msg = gHMCount & " HM -> " & JoinOuts(outs, nOut)
            If Not isPDF Then msg = msg & "  (file nhieu hang muc chi xuat PDF)"
            If Len(gLastErr2) > 0 Then msg = msg & "  [canh bao: " & gLastErr2 & "]"
        Else
            msg = gHMCount & " HM loi: " & IIf(Len(gLastErr2) > 0, Replace(gLastErr2, Chr(10), " "), "khong ro")
        End If
    Else
        Dim pdfPath As String
        pdfPath = OutFolderFor(wbk, gOutFolder1) & BaseNameOf(wbk.Name) & ".pdf"
        ok = AlignAndExportOneWorkbook(wbk, IIf(isPDF, "PDF", "PRINT"), pdfPath)
        If ok Then
            kind = 1
            msg = "1 HM -> " & IIf(isPDF, GetFileName2(pdfPath), "da gui may in")
        ElseIf Len(gLastErrDesc) > 0 Then
            kind = 1
            msg = "1 HM loi: " & gLastErrDesc
        Else
            msg = "Khong phai file du toan (khong tim thay sheet da chon)."
        End If
    End If

    If Not wasOpen Then
        On Error Resume Next
        wbk.Close SaveChanges:=False
        On Error GoTo 0
    End If

    Application.Calculation = savedCalc
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    API_ProcessFile = ok
End Function

' Quet file Excel (.xls/.xlsx/.xlsm) trong thu muc. Tra mang 1-based, cnt = so file.
Public Function API_CollectFiles(ByVal folderPath As String, inclSub As Boolean, ByRef cnt As Long) As Variant
    cnt = 0
    If Len(folderPath) = 0 Then Exit Function
    If Right$(folderPath, 1) <> Application.PathSeparator Then folderPath = folderPath & Application.PathSeparator
    Dim arr() As String
    ReDim arr(1 To 500)
    CollectExcelFiles folderPath, inclSub, arr, cnt
    If cnt = 0 Then Exit Function
    ReDim Preserve arr(1 To cnt)
    API_CollectFiles = arr
End Function

Public Function API_WriteLog(folderPath As String, logText As String) As String
    Dim p As String: p = folderPath
    If Len(p) > 0 And Right$(p, 1) <> Application.PathSeparator Then p = p & Application.PathSeparator
    API_WriteLog = WriteBatchErrorLog(p, logText)
End Function

' Thu muc luu PDF: thu muc nguoi dung chon (neu ton tai) hoac canh file goc.
' Tra "" neu file chua luu va khong chon thu muc.
Private Function OutFolderFor(wb As Workbook, userFolder As String) As String
    Dim f As String: f = Trim$(userFolder)
    If Len(f) > 0 Then
        If Not FolderExists2(f) Then f = ""
    End If
    If Len(f) = 0 Then f = wb.Path
    If Len(f) = 0 Then OutFolderFor = "": Exit Function
    If Right$(f, 1) <> Application.PathSeparator Then f = f & Application.PathSeparator
    OutFolderFor = f
End Function

Private Function FolderExists2(p As String) As Boolean
    On Error Resume Next
    FolderExists2 = CreateObject("Scripting.FileSystemObject").FolderExists(p)
    On Error GoTo 0
End Function

Private Function BaseNameOf(nm As String) As String
    Dim p As Long: p = InStrRev(nm, ".")
    If p > 1 Then BaseNameOf = Left$(nm, p - 1) Else BaseNameOf = nm
End Function

Private Function JoinOuts(outs() As String, n As Integer) As String
    Dim i As Integer, s As String
    For i = 1 To n
        If i > 1 Then s = s & " | "
        s = s & GetFileName2(outs(i))
    Next i
    JoinOuts = s
End Function

' ================================================================
' ======  XUAT PDF THEO FILE TONG HOP (V6.5)  ====================
' ================================================================
' Doc file tong hop, moi dong co cong thuc lien ket ra file khac la 1 MUC.
' Tu o lien ket, di theo tung cap:
'   - Sheet 'THKP hang muc' cua file chi tiet  -> muc in du toan (TJ_THKP),
'     dong THKP cho biet HANG MUC nao.
'   - Sheet khac cua 1 file du toan (co sheet THKP):
'       'TH chi phi TB' -> chi in sheet TB (TJ_TB)
'       'TH chi phi XD' / sheet khac -> in ca file du toan (TJ_THKP, dong 0)
'   - File trung gian (khong co THKP): doc DUNG o duoc tro toi
'       + o co lien ket ngoai  -> di tiep theo tung lien ket
'       + o la SUBTOTAL/SUM/+ cac o cung sheet -> di xuong tung o con
'         (SUBTOTAL bo qua o con cung la SUBTOTAL, dung nhu Excel)
'       + khong co lien ket nao xuong duoi -> sheet do la ban chi tiet,
'         in nguyen sheet (TJ_RAW) - vd bang bao gia thiet bi
' 1 dong file tong co the ra nhieu PDF: dat ten NN.01_, NN.02_...
' PDF duoc dat vao THU MUC CON giong cau truc thu muc cua file nguon
' (tinh tu thu muc file tong), nguoi dung tu ghep theo thu muc.
' ================================================================

Private Function TongRx() As Object
    If mTRx Is Nothing Then
        Set mTRx = CreateObject("VBScript.RegExp")
        mTRx.Global = True
        mTRx.IgnoreCase = True
        ' 'D:\thu muc\[file.xls]Sheet'!$F$13   hoac   [file.xls]Sheet1!F13 (file dang mo, ten khong dau cach)
        mTRx.Pattern = "(?:'([^']*)\[([^\]]+)\]([^']+)'|\[([^\]]+)\]([A-Za-z0-9_\.]+))!\$?([A-Za-z]{1,3})\$?([0-9]+)"
    End If
    Set TongRx = mTRx
End Function

' Tham chieu noi bo: A1, $A$1, A1:B5 (khong co ten sheet dung truoc)
Private Function TongIRx() As Object
    If mIRx Is Nothing Then
        Set mIRx = CreateObject("VBScript.RegExp")
        mIRx.Global = True
        mIRx.IgnoreCase = True
        mIRx.Pattern = "(^|[^A-Za-z0-9_!'\]\.\$:])\$?([A-Za-z]{1,3})\$?([0-9]{1,7})(:\$?([A-Za-z]{1,3})\$?([0-9]{1,7}))?"
    End If
    Set TongIRx = mIRx
End Function

Private Function TongColNum(ByVal a As String) As Long
    Dim i As Long, n As Long
    a = UCase$(a)
    For i = 1 To Len(a)
        n = n * 26 + (Asc(Mid$(a, i, 1)) - 64)
    Next i
    TongColNum = n
End Function

' Lay (duong dan, sheet, dong, cot) tu 1 match cua TongRx.
' baseWb: workbook chua cong thuc - de bu duong dan khi file dich dang mo
' (Excel bo duong dan trong cong thuc khi file lien ket dang mo).
Private Function TongRefFromMatch(m As Object, baseWb As Workbook, ByRef fullPath As String, _
        ByRef shName As String, ByRef rw As Long, ByRef col As Long) As Boolean
    TongRefFromMatch = False
    Dim d As String, f As String
    If Len(CStr(m.SubMatches(1))) > 0 Then
        d = CStr(m.SubMatches(0)): f = CStr(m.SubMatches(1)): shName = CStr(m.SubMatches(2))
    Else
        d = "": f = CStr(m.SubMatches(3)): shName = CStr(m.SubMatches(4))
    End If
    If Len(f) = 0 Then Exit Function
    If Len(d) = 0 Then
        Dim w As Workbook
        On Error Resume Next
        Set w = Application.Workbooks(f)
        On Error GoTo 0
        If Not w Is Nothing Then
            fullPath = w.FullName
        ElseIf Not baseWb Is Nothing Then
            fullPath = GetFolder2(baseWb.FullName) & f
        Else
            fullPath = f
        End If
    Else
        fullPath = d & f
    End If
    shName = Replace(shName, "''", "'")
    col = TongColNum(CStr(m.SubMatches(5)))
    rw = CLng(m.SubMatches(6))
    TongRefFromMatch = True
End Function

' Tach tham chieu ngoai dau tien trong cong thuc
Private Function TongParseRef(f As String, ByRef fullPath As String, ByRef shName As String, _
        ByRef rw As Long, ByRef col As Long, Optional baseWb As Workbook = Nothing) As Boolean
    TongParseRef = False
    If InStr(f, "[") = 0 Then Exit Function
    Dim m As Object
    Set m = TongRx().Execute(f)
    If m.Count = 0 Then Exit Function
    TongParseRef = TongRefFromMatch(m(0), baseWb, fullPath, shName, rw, col)
End Function

Private Function TongOpenWb(p As String, ByRef openedHere As Boolean) As Workbook
    Dim w As Workbook
    openedHere = False
    For Each w In Application.Workbooks
        If StrComp(w.FullName, p, vbTextCompare) = 0 Then Set TongOpenWb = w: Exit Function
    Next w
    On Error Resume Next
    Set TongOpenWb = Workbooks.Open(p, UpdateLinks:=0, ReadOnly:=True, _
                                    IgnoreReadOnlyRecommended:=True, Notify:=False, AddToMru:=False)
    On Error GoTo 0
    openedHere = Not (TongOpenWb Is Nothing)
End Function

' File trung gian: mo 1 lan, giu den het lan chay
Private Function TongMidWb(p As String) As Workbook
    Dim i As Long, own As Boolean, w As Workbook
    For i = 1 To mMidN
        If mMidKey(i) = LCase$(p) Then Set TongMidWb = mMidWb(i): Exit Function
    Next i
    Set w = TongOpenWb(p, own)
    If w Is Nothing Then Exit Function
    mMidN = mMidN + 1
    ReDim Preserve mMidWb(1 To mMidN)
    ReDim Preserve mMidOwn(1 To mMidN)
    ReDim Preserve mMidKey(1 To mMidN)
    Set mMidWb(mMidN) = w: mMidOwn(mMidN) = own: mMidKey(mMidN) = LCase$(p)
    Set TongMidWb = w
End Function

Private Sub TongCloseMids()
    Dim i As Long
    On Error Resume Next
    For i = 1 To mMidN
        If mMidOwn(i) Then mMidWb(i).Close SaveChanges:=False
    Next i
    On Error GoTo 0
    mMidN = 0
    Erase mMidWb: Erase mMidOwn: Erase mMidKey
End Sub

Private Sub TongLog(logTo As Object, s As String)
    If logTo Is Nothing Then
        Application.StatusBar = "Xuat theo file tong: " & s
        Exit Sub
    End If
    On Error Resume Next
    logTo.LogLine s
    On Error GoTo 0
End Sub

' Sheet cua file tong co nhieu cong thuc lien ket ngoai nhat
Private Function TongPickSheet(wb As Workbook) As Worksheet
    Dim ws As Worksheet, v As Variant, r As Long, c As Long, n As Long, best As Long
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            n = 0
            v = ws.Range("A1:L400").Formula
            For r = 1 To 400
                For c = 4 To 12
                    If Left$(CStr(v(r, c)), 1) = "=" Then
                        If InStr(CStr(v(r, c)), "[") > 0 And InStr(CStr(v(r, c)), "!") > 0 Then n = n + 1
                    End If
                Next c
            Next r
            If n > best Then best = n: Set TongPickSheet = ws
        End If
    Next ws
End Function

' Nhan STT: "A", "B.1" la dong tieu de phan; so nguyen la nhom
Private Function TongIsPartHdr(a As String) As Boolean
    TongIsPartHdr = (a Like "[A-Z]") Or (a Like "[A-Z].#")
End Function

' File du toan = co sheet THKP
Private Function TongIsDuToan(w As Workbook) As Boolean
    Dim s As Worksheet
    For Each s In w.Worksheets
        If InStr(NormUp2(s.Name), "THKP") > 0 Then TongIsDuToan = True: Exit Function
    Next s
End Function

' Ten muc con: chu o cot B cua dong (bo trong / so) ; khong co thi dung ten cha
Private Function TongRowName(ws As Worksheet, r As Long, fallback As String) As String
    Dim t As String
    On Error Resume Next
    t = Trim$(Replace(Replace(CStr(ws.Cells(r, 2).Text), vbLf, " "), vbCr, " "))
    On Error GoTo 0
    If Len(t) = 0 Or IsNumeric(t) Then t = fallback
    TongRowName = t
End Function

Private Function TongStripStrings(f As String) As String
    Dim s As String, i As Long, inQ As Boolean, ch As String
    For i = 1 To Len(f)
        ch = Mid$(f, i, 1)
        If ch = """" Then
            inQ = Not inQ
            s = s & " "
        ElseIf inQ Then
            s = s & " "
        Else
            s = s & ch
        End If
    Next i
    TongStripStrings = s
End Function

Private Function TongAddJob(nm As String, path As String, kind As Integer, sh As String, _
        rw As Long, note As String) As Long
    mTN = mTN + 1
    If mTN > UBound(mTJ) Then ReDim Preserve mTJ(1 To mTN + 100)
    With mTJ(mTN)
        .Seq = mCurTop: .SumRow = mCurSumRow: .Label = mCurLabel
        .Tag = Format$(mCurTop, "00")
        .Nm = nm: .DetPath = path: .Kind = kind: .SheetNm = sh
        .ThkpRow = rw: .Note = note
    End With
    TongAddJob = mTN
End Function

' Di theo 1 tham chieu (file, sheet, dong, cot) -> them cac muc in vao mTJ
Private Sub TongExpand(ByVal path As String, ByVal sh As String, ByVal col As Long, ByVal rw As Long, _
        ByVal depth As Integer, ByVal nm As String)
    Dim shU As String: shU = NormUp2(sh)
    If InStr(shU, "THKP") > 0 Then
        TongAddJob nm, path, TJ_THKP, sh, rw, ""
        Exit Sub
    End If
    If depth > 6 Then
        TongAddJob nm, "", TJ_THKP, "", 0, "Lien ket qua sau (>6 cap), dung tai: " & path
        Exit Sub
    End If
    Dim key As String: key = LCase$("x|" & path & "|" & sh & "|" & rw & "|" & col)
    If mVis.Exists(key) Then Exit Sub
    mVis.Add key, 1

    Dim w As Workbook: Set w = TongMidWb(path)
    If w Is Nothing Then
        TongAddJob nm, "", TJ_THKP, "", 0, "Khong mo duoc file: " & path
        Exit Sub
    End If
    If TongIsDuToan(w) Then
        If InStr(shU, "TH CHI PHI TB") > 0 Then
            TongAddJob nm, path, TJ_TB, sh, 0, ""
        Else
            ' vd 'TH chi phi XD'!C5: ca file du toan ; ten hang muc o cot B cua dong do (de chon khoi khi file nhieu HM)
            Dim jj As Long, hmT As String, wsX As Worksheet
            jj = TongAddJob(nm, path, TJ_THKP, sh, 0, "")
            On Error Resume Next
            Set wsX = w.Worksheets(sh)
            If Not wsX Is Nothing Then hmT = Trim$(CStr(wsX.Cells(rw, 2).Text))
            On Error GoTo 0
            If Left$(hmT, 1) = "+" Then hmT = Trim$(Mid$(hmT, 2))
            If InStr(NormUp2(hmT), "HANG MUC") = 1 And InStr(hmT, ":") > 0 Then hmT = Trim$(Mid$(hmT, InStr(hmT, ":") + 1))
            If InStr(NormUp2(hmT), "TONG CONG") > 0 Then hmT = ""
            mTJ(jj).HmName = hmT
        End If
        Exit Sub
    End If

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = w.Worksheets(sh)
    On Error GoTo 0
    If ws Is Nothing Then
        TongAddJob nm, "", TJ_THKP, "", 0, "File '" & GetFileName2(path) & "' khong co sheet '" & sh & "'"
        Exit Sub
    End If
    Dim n As Long
    n = TongWalkCell(ws, rw, col, depth, nm, 0)
    If n = 0 Then
        ' khong con lien ket xuong duoi -> chinh sheet nay la ban chi tiet
        TongAddJob nm, path, TJ_RAW, ws.Name, 0, ""
    End If
End Sub

' Di xuong 1 o cua file trung gian. Tra ve so tham chieu ngoai da di theo.
Private Function TongWalkCell(ws As Worksheet, ByVal r As Long, ByVal c As Long, _
        ByVal depth As Integer, ByVal nm As String, ByVal lvl As Integer) As Long
    TongWalkCell = 0
    If lvl > 8 Then Exit Function
    mWalkBudget = mWalkBudget - 1
    If mWalkBudget < 0 Then Exit Function
    Dim key As String: key = LCase$("c|" & ws.Parent.FullName & "|" & ws.Name & "|" & r & "|" & c)
    If mVis.Exists(key) Then Exit Function
    mVis.Add key, 1

    Dim fm As String
    On Error Resume Next
    fm = CStr(ws.Cells(r, c).Formula)
    On Error GoTo 0
    If Left$(fm, 1) <> "=" Then Exit Function

    Dim rowNm As String: rowNm = TongRowName(ws, r, nm)
    Dim n As Long, m As Object, i As Long
    Dim p As String, s As String, rr As Long, cc As Long

    ' 1) lien ket ra file khac
    If InStr(fm, "[") > 0 Then
        Set m = TongRx().Execute(fm)
        For i = 0 To m.Count - 1
            If TongRefFromMatch(m(i), ws.Parent, p, s, rr, cc) Then
                TongExpand p, s, cc, rr, depth + 1, rowNm
                n = n + 1
            End If
        Next i
        If n > 0 Then TongWalkCell = n: Exit Function
    End If

    ' 2) tham chieu cung sheet: SUBTOTAL(9,G9:G35), SUM(...), =E11+E12 ...
    Dim isSub As Boolean: isSub = (InStr(UCase$(fm), "SUBTOTAL") > 0)
    Set m = TongIRx().Execute(TongStripStrings(fm))
    Dim r1 As Long, r2 As Long, c1 As Long, c2 As Long, x As Long, y As Long, f2 As String
    For i = 0 To m.Count - 1
        c1 = TongColNum(CStr(m(i).SubMatches(1))): r1 = CLng(m(i).SubMatches(2))
        If Len(CStr(m(i).SubMatches(3))) > 0 Then
            c2 = TongColNum(CStr(m(i).SubMatches(4))): r2 = CLng(m(i).SubMatches(5))
        Else
            c2 = c1: r2 = r1
        End If
        If c1 < 1 Or c1 > 200 Or c2 > 200 Then GoTo NextRef
        If r2 < r1 Then y = r1: r1 = r2: r2 = y
        If c2 < c1 Then x = c1: c1 = c2: c2 = x
        If r2 - r1 > 500 Then r2 = r1 + 500
        For y = r1 To r2
            For x = c1 To c2
                If Not (y = r And x = c) Then
                    f2 = ""
                    On Error Resume Next
                    f2 = CStr(ws.Cells(y, x).Formula)
                    On Error GoTo 0
                    If Left$(f2, 1) = "=" Then
                        If Not (isSub And InStr(UCase$(f2), "SUBTOTAL") > 0) Then
                            n = n + TongWalkCell(ws, y, x, depth, rowNm, lvl + 1)
                        End If
                    End If
                End If
            Next x
        Next y
NextRef:
    Next i
    TongWalkCell = n
End Function

' Thu muc con (tuong doi) cua file nguon so voi thu muc file tong, co "\" o cuoi
Private Function TongSubDir(detPath As String, sumFolder As String) As String
    TongSubDir = ""
    If Len(detPath) = 0 Then Exit Function
    Dim d As String: d = GetFolder2(detPath)
    If StrComp(Left$(d, Len(sumFolder)), sumFolder, vbTextCompare) = 0 Then
        TongSubDir = Mid$(d, Len(sumFolder) + 1)
    Else
        TongSubDir = "_NGOAI THU MUC TONG\"
    End If
End Function

Private Sub TongMkDirs(ByVal p As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    If Len(p) = 0 Then Exit Sub
    If fso.FolderExists(p) Then Exit Sub
    TongMkDirs fso.GetParentFolderName(p)
    On Error Resume Next
    fso.CreateFolder p
    On Error GoTo 0
End Sub

Private Function TongHmOf(ws As Worksheet, r As Long) As Integer
    TongHmOf = 0
    If ws Is Nothing Then Exit Function
    If r <= 0 Then Exit Function
    ws.Cells.EntireRow.Hidden = False
    Dim t() As Long, b() As Long, kd() As String, h() As Long, n As Integer
    ScanBlocks ws, 1, t, b, kd, h, n
    If n = 0 Then Exit Function
    Dim i As Integer, phys As Integer
    For i = 1 To n
        If r >= t(i) And r <= b(i) Then phys = i: Exit For
    Next i
    If phys = 0 Then
        For i = 1 To n
            If h(i) <= r Then phys = i
        Next i
    End If
    If phys = 0 Then Exit Function
    Dim hm As Integer
    For hm = 1 To gHMCount
        If PhysBlockForHm(ws, 1, hm) = phys Then TongHmOf = hm: Exit Function
    Next hm
End Function

' Hang muc (1..gHMCount) theo ten: trung khop truoc, roi chua nhau
Private Function TongHmByName(nm As String) As Integer
    Dim k As Integer, a As String, b As String
    a = Trim$(NormUp2(nm))
    If Len(a) = 0 Then Exit Function
    For k = 1 To gHMCount
        If Trim$(NormUp2(gHMNames(k))) = a Then TongHmByName = k: Exit Function
    Next k
    For k = 1 To gHMCount
        b = Trim$(NormUp2(gHMNames(k)))
        If Len(b) > 0 Then
            If InStr(b, a) > 0 Or InStr(a, b) > 0 Then TongHmByName = k: Exit Function
        End If
    Next k
End Function

' In 1 muc TJ_TB / TJ_RAW (1 sheet -> 1 PDF)
Private Function TongExportSheetJob(wbk As Workbook, j As Long) As Boolean
    TongExportSheetJob = False
    Dim k As Integer, ok As Boolean
    gLastErrDesc = ""
    If mTJ(j).Kind = TJ_TB Then
        For k = 1 To SHEET_COUNT
            gSelSheet(k) = False: gOrder(k) = 99
        Next k
        gSelSheet(8) = True: gOrder(8) = 1          ' TH chi phi TB
        ok = AlignAndExportOneWorkbook(wbk, "PDF", mTJ(j).OutPdf)
        If ok And FileExists2(mTJ(j).OutPdf) Then TongExportSheetJob = True: Exit Function
        ' khong nhan duoc sheet TB chuan -> in nguyen sheet
    End If
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wbk.Worksheets(mTJ(j).SheetNm)
    On Error GoTo 0
    If ws Is Nothing Then mTJ(j).Note = "Khong thay sheet '" & mTJ(j).SheetNm & "'": Exit Function
    On Error Resume Next
    If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
    KillFile2 mTJ(j).OutPdf
    Err.Clear
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=mTJ(j).OutPdf, Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
    If Err.Number <> 0 Then mTJ(j).Note = "Loi xuat PDF " & Err.Number & ": " & Err.Description
    Err.Clear
    On Error GoTo 0
    TongExportSheetJob = FileExists2(mTJ(j).OutPdf)
End Function

' Xuat tat ca muc cua 1 file chi tiet
Private Sub TongDoFile(path As String, jIdxIn() As Long, nJIn As Long, inclTong As Boolean, logTo As Object)
    Dim i As Long, k As Integer, hm As Integer, q As Long
    Dim wbk As Workbook, own As Boolean
    Set wbk = TongOpenWb(path, own)
    If wbk Is Nothing Then
        For i = 1 To nJIn: mTJ(jIdxIn(i)).Note = "Khong mo duoc file: " & path: Next i
        Exit Sub
    End If

    ' ---- muc TB / RAW: moi sheet 1 PDF ; cung sheet -> chung PDF ----
    Dim jIdx() As Long, nJ As Long
    ReDim jIdx(1 To nJIn)
    For i = 1 To nJIn
        With mTJ(jIdxIn(i))
            If .Kind = TJ_THKP Then
                nJ = nJ + 1: jIdx(nJ) = jIdxIn(i)
            Else
                For q = 1 To i - 1
                    If mTJ(jIdxIn(q)).Kind = .Kind And StrComp(mTJ(jIdxIn(q)).SheetNm, .SheetNm, vbTextCompare) = 0 _
                       And mTJ(jIdxIn(q)).DupOf = 0 Then Exit For
                Next q
                If q < i Then
                    .DupOf = jIdxIn(q)
                    .Note = "Trung sheet voi muc " & mTJ(jIdxIn(q)).Tag & " - xuat chung 1 PDF"
                Else
                    .Done = TongExportSheetJob(wbk, jIdxIn(i))
                    If Not .Done And Len(.Note) = 0 Then .Note = "Loi xuat PDF: " & IIf(Len(gLastErrDesc) > 0, gLastErrDesc, "khong ro")
                End If
            End If
        End With
    Next i
    If nJ = 0 Then GoTo CloseIt

    gLastErr2 = "": gLastErrDesc = ""
    gGiaTongHopInTongCong = inclTong
    gGTH_InTongCong2 = inclTong
    For hm = 1 To MAX_HM: gForcePdf(hm) = "": Next hm

    If DetectTwoHangMuc(wbk) Then
        ' ---- file nhieu hang muc: moi muc 1 PDF ----
        Dim wsT As Worksheet: Set wsT = Find2Sheet(wbk, 1)
        Dim owner(1 To 20) As Long        ' hang muc -> muc dau tien dung no
        Dim lstJ(1 To 20) As Long, lstHm(1 To 20) As Integer, lstPdf(1 To 20) As String, nL As Long
        Dim wholeJ As Long                ' muc lien ket vao dong TONG ('TH chi phi XD' dong tong) -> ca file
        For i = 1 To nJ
            With mTJ(jIdx(i))
                hm = TongHmOf(wsT, .ThkpRow)
                If hm = 0 And Len(.HmName) > 0 Then hm = TongHmByName(.HmName)
                If hm = 0 And .ThkpRow = 0 Then
                    If wholeJ = 0 Then
                        wholeJ = jIdx(i)
                    Else
                        .DupOf = wholeJ
                        .Note = "Ca file - xuat chung voi muc " & mTJ(wholeJ).Tag
                    End If
                ElseIf hm = 0 Then
                    .Note = "Khong xac dinh duoc hang muc (THKP dong " & .ThkpRow & ") trong file nhieu hang muc"
                ElseIf owner(hm) <> 0 Then
                    .Hm = hm: .DupOf = owner(hm)
                    .Note = "Trung hang muc voi muc " & mTJ(owner(hm)).Tag & " - xuat chung 1 PDF"
                Else
                    .Hm = hm: owner(hm) = jIdx(i)
                    nL = nL + 1: lstJ(nL) = jIdx(i): lstHm(nL) = hm: lstPdf(nL) = .OutPdf
                End If
            End With
        Next i
        ' Muc "ca file": moi hang muc chua co PDF rieng -> <ten>_HMk.pdf (nguoi dung tu ghep)
        If wholeJ > 0 Then
            Dim baseP As String: baseP = Left$(mTJ(wholeJ).OutPdf, Len(mTJ(wholeJ).OutPdf) - 4)
            Dim nW As Integer
            For hm = 1 To gHMCount
                If owner(hm) = 0 And nL < 20 Then
                    nW = nW + 1
                    nL = nL + 1: lstJ(nL) = wholeJ: lstHm(nL) = hm
                    lstPdf(nL) = baseP & "_HM" & hm & ".pdf"
                    owner(hm) = wholeJ
                End If
            Next hm
            If nW = 0 Then
                mTJ(wholeJ).Note = "Ca file - cac hang muc da co trong PDF cua muc khac"
                mTJ(wholeJ).DupOf = owner(1)
            Else
                mTJ(wholeJ).Note = "Ca file " & gHMCount & " hang muc - xuat " & nW & " PDF <ten>_HMk.pdf (hang muc da co PDF o muc khac thi khong xuat lai)"
            End If
        End If

        If nL > 0 Then
            Dim a As Integer, b As Integer
            For a = 1 To MAX_HM
                For b = 1 To HM_SHEET_COUNT: gSelN(b, a) = False: Next b
            Next a
            For i = 1 To nL
                hm = lstHm(i)
                gSelN(1, hm) = True: gSelN(2, hm) = True: gSelN(3, hm) = Not gTongNoDGCT
                gSelN(5, hm) = True               ' Gia thang -> di kem MOI hang muc
                gForcePdf(hm) = lstPdf(i)
            Next i
            gSelN(6, lstHm(1)) = True              ' TH chi phi TB -> hang muc dau
            gOrderN(6) = 1: gOrderN(1) = 2: gOrderN(2) = 3: gOrderN(3) = 4: gOrderN(5) = 5: gOrderN(4) = 99
            gHM_XuatKieu = "TACH"
            gHM_TenTheoHM = False

            Dim outs() As String, nOut As Integer, ok As Boolean
            ok = ProcessWorkbookNHM(wbk, outs, nOut)
            For i = 1 To nL
                With mTJ(lstJ(i))
                    If lstPdf(i) = .OutPdf Then
                        .Done = FileExists2(.OutPdf)
                    Else
                        If FileExists2(lstPdf(i)) Then .Done = True
                    End If
                    If Not .Done Then .Note = "Loi xuat PDF: " & IIf(Len(gLastErr2) > 0, Replace(gLastErr2, vbLf, " "), "khong ro")
                    If .Done And Len(gLastErr2) > 0 Then .Note = .Note & IIf(Len(.Note) > 0, " ; ", "") & "Canh bao: " & Replace(gLastErr2, vbLf, " ")
                End With
            Next i
        End If
        For hm = 1 To MAX_HM: gForcePdf(hm) = "": Next hm
    Else
        ' ---- file 1 hang muc: tat ca muc tro cung 1 PDF ----
        Dim first As Long: first = jIdx(1)
        For i = 2 To nJ
            mTJ(jIdx(i)).DupOf = first
            mTJ(jIdx(i)).Note = "File chi co 1 hang muc - xuat chung PDF voi muc " & mTJ(first).Tag
        Next i
        For k = 1 To SHEET_COUNT
            gSelSheet(k) = False: gOrder(k) = 99
        Next k
        gSelSheet(8) = True: gOrder(8) = 1      ' TH chi phi TB
        gSelSheet(1) = True: gOrder(1) = 2      ' THKP
        gSelSheet(2) = True: gOrder(2) = 3      ' Gia tong hop
        gSelSheet(3) = Not gTongNoDGCT: gOrder(3) = 4      ' Don gia chi tiet
        gSelSheet(5) = True: gOrder(5) = 5      ' Gia thang
        Dim okE As Boolean
        okE = AlignAndExportOneWorkbook(wbk, "PDF", mTJ(first).OutPdf)
        mTJ(first).Done = okE And FileExists2(mTJ(first).OutPdf)
        If Not mTJ(first).Done Then
            mTJ(first).Note = "Loi xuat PDF: " & IIf(Len(gLastErrDesc) > 0, gLastErrDesc, "khong tim thay sheet du toan")
        End If
    End If

CloseIt:
    If own Then
        On Error Resume Next
        wbk.Close SaveChanges:=False
        On Error GoTo 0
    End If
End Sub

' Xuat PDF theo file tong hop.
'   sumPath      : file tong hop (.xlsx/.xls) - chua cac cong thuc lien ket
'   outFolder    : thu muc luu PDF (rong = <thu muc file tong>\PDF)
'   inclTongCong : in cot Tong cong o Gia tong hop
'   fromLeaf     : chay tu muc so NN (so thu tu trong ten PDF; 1 = tu dau) - de chay tiep khi bi ngat
'   toLeaf       : chay den muc so NN (0 = den het)
'   logTo        : form co Sub LogLine(s) de hien nhat ky (co the Nothing)
' Tra ve so PDF da xuat; report = bao cao day du (da ghi _THU_TU_XUAT.txt)
Public Function API_RunTong(sumPath As String, ByVal outFolder As String, inclTongCong As Boolean, _
        ByRef report As String, Optional fromLeaf As Long = 1, Optional toLeaf As Long = 0, _
        Optional logTo As Object = Nothing) As Long
    API_RunTong = TongRun(sumPath, outFolder, inclTongCong, report, fromLeaf, toLeaf, logTo)
    gTongReport = report
End Function

' Chi xuat cac muc co file chi tiet nam trong duong dan chua chuoi nay (vd "\1. DHKK\"); rong = tat ca
Public Sub API_SetTongFilter(ByVal s As String)
    gTongFilter = s
End Sub

Public Sub API_SetTongNoDGCT(ByVal b As Boolean)
    gTongNoDGCT = b
End Sub

' True = chi do lien ket + ghi bao cao, KHONG xuat PDF (xem truoc danh sach)
Public Sub API_SetTongDryRun(ByVal b As Boolean)
    gTongDryRun = b
End Sub

Public Function API_TongReport() As String
    API_TongReport = gTongReport
End Function

Private Function TongRun(sumPath As String, ByVal outFolder As String, inclTongCong As Boolean, _
        ByRef report As String, fromLeaf As Long, toLeaf As Long, logTo As Object) As Long
    TongRun = 0
    report = ""
    If Not FileExists2(sumPath) Then report = "Khong thay file tong hop: " & sumPath: Exit Function

    outFolder = Trim$(outFolder)
    If Len(outFolder) = 0 Then outFolder = GetFolder2(sumPath) & "PDF"
    If Right$(outFolder, 1) = "\" Then outFolder = Left$(outFolder, Len(outFolder) - 1)
    If Not FolderExists2(outFolder) Then
        TongMkDirs outFolder
        If Not FolderExists2(outFolder) Then report = "Khong tao duoc thu muc luu PDF: " & outFolder: Exit Function
    End If
    outFolder = outFolder & "\"
    Dim sumFolder As String: sumFolder = GetFolder2(sumPath)

    Dim savedCalc As XlCalculation: savedCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    Dim wbS As Workbook, ownS As Boolean
    Set wbS = TongOpenWb(sumPath, ownS)
    If wbS Is Nothing Then
        report = "Khong mo duoc file tong hop."
        GoTo Done
    End If
    Dim ws As Worksheet: Set ws = TongPickSheet(wbS)
    If ws Is Nothing Then
        report = "File tong hop khong co sheet nao chua cong thuc lien ket toi file khac."
        GoTo Done
    End If
    TongLog logTo, "Doc sheet '" & ws.Name & "' cua file tong hop..."

    ' ---- 1) doc cac dong co lien ket, di theo tung cap ----
    Dim lastRow As Long
    lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If lastRow > 1500 Then lastRow = 1500
    Dim fm As Variant: fm = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 12)).Formula
    Dim r As Long, c As Long, a As String, bnm As String, hdr As String
    Dim p As String, s As String, rw As Long, cl As Long, hit As Boolean
    Dim nTop As Long, firstJ As Long, cntJ As Long, mm As Object, q As Long
    mTN = 0
    ReDim mTJ(1 To 200)
    ' cot lien ket = cot (D..L) co nhieu cong thuc tro ra file khac nhat (vd cot G, H)
    Dim cnt(4 To 12) As Long, linkCol As Long, bestCnt As Long
    For r = 1 To lastRow
        For c = 4 To 12
            If InStr(CStr(fm(r, c)), "[") > 0 And InStr(CStr(fm(r, c)), "!") > 0 Then cnt(c) = cnt(c) + 1
        Next c
    Next r
    linkCol = 7
    For c = 4 To 12
        If cnt(c) > bestCnt Then bestCnt = cnt(c): linkCol = c
    Next c
    For r = 1 To lastRow
        hit = TongParseRef(CStr(fm(r, linkCol)), p, s, rw, cl, wbS)
        a = Replace(Trim$(ws.Cells(r, 1).Text), ",", ".")
        If hit Then
            If toLeaf > 0 And nTop >= toLeaf Then Exit For
            nTop = nTop + 1
            bnm = Trim$(Replace(Replace(ws.Cells(r, 2).Text, vbLf, " "), vbCr, " "))
            mCurTop = nTop: mCurSumRow = r
            mCurLabel = IIf(Len(hdr) > 0, hdr & "-", "") & a
            If nTop < fromLeaf Then
                TongAddJob bnm, "", TJ_THKP, "", 0, "Ngoai pham vi chay"
            Else
                TongLog logTo, "Do lien ket muc " & Format$(nTop, "00") & " " & bnm
                DoEvents
                Set mVis = CreateObject("Scripting.Dictionary")
                mWalkBudget = 5000
                firstJ = mTN + 1
                Set mm = TongRx().Execute(CStr(fm(r, linkCol)))
                For q = 0 To mm.Count - 1
                    If TongRefFromMatch(mm(q), wbS, p, s, rw, cl) Then TongExpand p, s, cl, rw, 0, bnm
                Next q
                If mTN < firstJ Then
                    TongAddJob bnm, "", TJ_THKP, "", 0, "Khong co file du toan chi tiet (lien ket noi bo / gia tri nhap tay)"
                End If
                cntJ = mTN - firstJ + 1
                For q = firstJ To mTN
                    With mTJ(q)
                        If cntJ > 1 Then .Tag = Format$(nTop, "00") & "." & Format$(q - firstJ + 1, "00")
                        If Len(gTongFilter) > 0 And Len(.DetPath) > 0 Then
                            If InStr(1, .DetPath, gTongFilter, vbTextCompare) = 0 Then
                                .DetPath = "": .Note = "Ngoai pham vi chay"
                            End If
                        End If
                    End With
                Next q
            End If
        ElseIf TongIsPartHdr(a) Then
            hdr = a
        End If
    Next r
    If mTN = 0 Then report = "Khong co dong lien ket nao trong file tong hop.": GoTo Done

    ' ---- 2) ten PDF (trong thu muc con giong nguon) + gom theo file chi tiet ----
    Dim i As Long, j As Long, sd As String
    Dim files() As String, nF As Long
    ReDim files(1 To mTN)
    For i = 1 To mTN
        With mTJ(i)
            sd = TongSubDir(.DetPath, sumFolder)
            .OutPdf = outFolder & sd & .Tag & "_" & CleanName2(.Label & " " & .Nm, "HM") & ".pdf"
            If Len(.DetPath) > 0 Then
                If Len(sd) > 0 Then TongMkDirs outFolder & sd
                For j = 1 To nF
                    If StrComp(files(j), .DetPath, vbTextCompare) = 0 Then Exit For
                Next j
                If j > nF Then nF = nF + 1: files(nF) = .DetPath
            End If
        End With
    Next i
    TongLog logTo, nTop & " muc trong file tong -> " & mTN & " muc in, " & nF & " file chi tiet can xu ly."

    ' ---- 3) xu ly tung file chi tiet ----
    Dim jIdx() As Long, nJ As Long, f As Long
    ReDim jIdx(1 To mTN)
    If gTongDryRun Then nF = 0                ' chay kho: khong xuat
    For f = 1 To nF
        nJ = 0
        For i = 1 To mTN
            If StrComp(mTJ(i).DetPath, files(f), vbTextCompare) = 0 Then nJ = nJ + 1: jIdx(nJ) = i
        Next i
        TongLog logTo, "[" & f & "/" & nF & "] " & GetFileName2(files(f)) & "  (" & nJ & " muc)"
        DoEvents
        TongDoFile files(f), jIdx, nJ, inclTongCong, logTo
        For i = 1 To nJ
            With mTJ(jIdx(i))
                If .Done Then TongLog logTo, "     OK  " & Mid$(.OutPdf, Len(outFolder) + 1)
            End With
        Next i
    Next f

    ' ---- 4) bao cao ----
    Dim nOK As Long, nBad As Long, rp As String, lastTop As Long, kindTxt As String
    rp = "XUAT PDF THEO FILE TONG HOP (V6.5) - " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & _
         "File tong : " & sumPath & vbCrLf & "Sheet     : " & ws.Name & vbCrLf & _
         "Thu muc   : " & outFolder & vbCrLf & String$(70, "=") & vbCrLf
    For i = 1 To mTN
        With mTJ(i)
            If .Seq <> lastTop Then
                rp = rp & Format$(.Seq, "00") & " | " & .Label & " | " & Trim$(Replace(Replace(CStr(ws.Cells(.SumRow, 2).Text), vbLf, " "), vbCr, " ")) & vbCrLf
                lastTop = .Seq
            End If
            rp = rp & "   " & .Tag & " " & .Nm & vbCrLf
            If .Done Then
                nOK = nOK + 1
                rp = rp & "        -> " & Mid$(.OutPdf, Len(outFolder) + 1) & IIf(Len(.Note) > 0, "   [" & .Note & "]", "") & vbCrLf
            ElseIf .DupOf > 0 Then
                rp = rp & "        -> (chung PDF voi muc " & mTJ(.DupOf).Tag & ")" & vbCrLf
            ElseIf .Note = "Ngoai pham vi chay" Then
                rp = rp & "        -> (khong chay: ngoai pham vi)" & vbCrLf
            ElseIf gTongDryRun And Len(.DetPath) > 0 Then
                rp = rp & "        -> (chay kho) " & Mid$(.OutPdf, Len(outFolder) + 1) & IIf(Len(.HmName) > 0, "   [HM: " & .HmName & "]", "") & vbCrLf
            Else
                nBad = nBad + 1
                rp = rp & "        -> KHONG XUAT: " & .Note & vbCrLf
            End If
            If Len(.DetPath) > 0 Then
                Select Case .Kind
                    Case TJ_TB: kindTxt = "  (sheet " & .SheetNm & ")"
                    Case TJ_RAW: kindTxt = "  (in nguyen sheet " & .SheetNm & ")"
                    Case Else: kindTxt = IIf(.ThkpRow > 0, "  (THKP dong " & .ThkpRow & ")", "  (ca file)")
                End Select
                rp = rp & "           nguon: " & .DetPath & kindTxt & vbCrLf
            End If
        End With
    Next i
    rp = rp & String$(70, "=") & vbCrLf & nTop & " muc file tong -> " & mTN & " muc in. Da xuat " & nOK & " PDF, khong xuat " & nBad & " muc."
    report = rp
    On Error Resume Next
    Dim ts As Object
    Set ts = CreateObject("Scripting.FileSystemObject").CreateTextFile(outFolder & "_THU_TU_XUAT.txt", True, True)
    ts.Write rp: ts.Close
    On Error GoTo 0
    TongRun = nOK

Done:
    TongCloseMids
    Set mVis = Nothing
    On Error Resume Next
    If Not wbS Is Nothing Then
        If ownS Then wbS.Close SaveChanges:=False
    End If
    On Error GoTo 0
    Application.Calculation = savedCalc
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Function
