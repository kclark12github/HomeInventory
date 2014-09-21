Attribute VB_Name = "libKeyPress"
Public Sub KeyPressUcase(KeyAscii As Integer)
    Dim Char As String
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
End Sub
Public Sub KeyPressInteger(KeyAscii As Integer)
    If KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 45 And (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 7
    End If
End Sub
Public Sub KeyPressReal(KeyAscii As Integer)
    If KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 45 And (KeyAscii < 46 Or KeyAscii > 57) Then
        KeyAscii = 7
    End If
End Sub
Public Sub ValidateCurrency(strField As String, Cancel As Integer)
    If strField = vbNullString Then strField = Format(0, "Currency")
    If Not IsNumeric(strField) Then
        MsgBox "Invalid currency value entered.", vbExclamation, Screen.ActiveForm.Caption
        TextSelected
        Cancel = True
    End If
End Sub
