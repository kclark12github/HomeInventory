Attribute VB_Name = "fmtCurrency"
Public Function fmtCurrency(ctl As TextBox) As Boolean
    FormatCurrency = True
    On Error Resume Next
    ctl = Format(CCur(ctl.Text), "Currency")
    If Err.Number <> 0 Then
        FormatCurrency = False
        MsgBox "Invalid currency value entered: " & ctl.Text & ". Please enter a valid dollar amount.", vbExclamation
        TextSelected
        Exit Function
    End If
    
    If CCur(ctl.Text) < 0 Then
        ctl.ForeColor = vbRed
    Else
        ctl.ForeColor = vbBlack
    End If
End Function
