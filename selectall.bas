    '*********************************************
    '* Submitted by Mike Shaffer
    '*********************************************
Public Sub TextSelected()
 ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'
 ':::                                                                   :::'
 ':::   Selects all of the text in the current textbox                  :::'
 ':::   (call from the textbox GetFocus event)                          :::'
 ':::                                                                   :::'
 ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'
   '
   Dim i As Integer
   Dim oTextBox As TextBox
   '
   If TypeOf Screen.ActiveControl Is TextBox Then
      Set oTextBox = Screen.ActiveControl
      i = Len(oTextBox.Text)
      oTextBox.SelStart = 0
      oTextBox.SelLength = i
   End If
   '
End Sub
