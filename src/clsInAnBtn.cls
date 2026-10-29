' Nut gia (Label) cua frmInDuToan: chuyen Click / MouseMove ve form.
Public WithEvents L As MSForms.Label
Public Nm As String
Public Host As Object

Private Sub L_Click()
    Host.OnBtn Nm
End Sub

Private Sub L_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Host.OnHover Nm
End Sub
