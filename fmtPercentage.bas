Attribute VB_Name = "fmtPercentage"
Public Function fmtPercentage(ctl As TextBox) As Boolean
    Dim strTemp As String
    Dim i As Integer
    
    fmtPercentage = True
    
    On Error Resume Next
    strTemp = ctl.Text
    i = InStr(strTemp, "%")
    If i > 0 Then strTemp = Mid(strTemp, 1, i - 1)
    ctl.Text = Format(CDbl(strTemp) / 100, "Percent")
    If Err.Number <> 0 Then
        fmtPercentage = False
        MsgBox "Invalid percent value entered: " & ctl.Text & ". Please enter a valid percentage.", vbExclamation
        TextSelected
        Exit Function
    End If
    
    If CDbl(strTemp) < 0 Then
        ctl.ForeColor = vbRed
    Else
        ctl.ForeColor = vbBlack
    End If
End Function
Public Function DecodePercentage(ByVal strPercentage As String) As Double
    Dim i As Integer
    
    i = InStr(strPercentage, "%")
    If i > 0 Then strPercentage = Mid(strPercentage, 1, i - 1)
    DecodePercentage = CDbl(strPercentage) / 100
End Function
