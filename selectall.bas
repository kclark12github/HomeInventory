Attribute VB_Name = "SelectAll"
Public Sub TextSelected()
    Dim i As Integer
    Dim ctl As Control
   
    Select Case TypeName(Screen.ActiveControl)
        Case "TextBox", "ComboBox", "DataCombo"
            Set ctl = Screen.ActiveControl
            i = Len(ctl.Text)
            ctl.SelStart = 0
            ctl.SelLength = i
    End Select
End Sub
