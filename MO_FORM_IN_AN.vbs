' MO_FORM_IN_AN.vbs - mo thang form "In an du toan G8", khong can mo file .xlsm bang tay.
' Dat cung thu muc voi FILE_TONG_IN_G8_V*.xlsm (lay ban moi nhat). Chay qua MO_FORM_IN_AN.bat.
' - Excel dang mo: dung luon (form thay cac file du toan dang mo).
' - File tool mo CHI DOC va AN cua so; dong form -> tu dong file tool.
Option Explicit
Dim fso, here, f, best, bestDate, xl, wb, w, created, opened
Set fso = CreateObject("Scripting.FileSystemObject")
here = fso.GetParentFolderName(WScript.ScriptFullName)

For Each f In fso.GetFolder(here).Files
    If LCase(fso.GetExtensionName(f.Name)) = "xlsm" And InStr(1, f.Name, "FILE_TONG_IN_G8_V", 1) = 1 _
       And InStr(1, f.Name, "_test", 1) = 0 And Left(f.Name, 2) <> "~$" Then
        If IsEmpty(bestDate) Then
            best = f.Path: bestDate = f.DateLastModified
        ElseIf f.DateLastModified > bestDate Then
            best = f.Path: bestDate = f.DateLastModified
        End If
    End If
Next
If IsEmpty(best) Then
    MsgBox "Khong tim thay FILE_TONG_IN_G8_V*.xlsm trong thu muc:" & vbCrLf & here, vbExclamation, "In an du toan G8"
    WScript.Quit 1
End If

On Error Resume Next
Set xl = GetObject(, "Excel.Application")
On Error GoTo 0
created = False
If IsEmpty(xl) Or (TypeName(xl) <> "Application") Then
    Set xl = CreateObject("Excel.Application")
    created = True
End If

Set wb = Nothing
' Excel khong mo duoc 2 file trung ten -> file tool cung ten dang mo (o dau cung vay) thi dung luon
For Each w In xl.Workbooks
    If LCase(w.Name) = LCase(fso.GetFileName(best)) Then Set wb = w
Next
opened = False
If wb Is Nothing Then
    opened = True
    xl.EnableEvents = False
    Set wb = xl.Workbooks.Open(best, 0, True)
    xl.EnableEvents = True
End If

xl.Visible = True
If created Then xl.WindowState = -4137     ' xlMaximized
If opened Then
    xl.Run "'" & wb.Name & "'!MoFormTuBat"       ' an file tool, dong form -> dong file tool
Else
    xl.Run "'" & wb.Name & "'!MoBangDieuKhien"   ' file tool nguoi dung tu mo: giu nguyen
End If
