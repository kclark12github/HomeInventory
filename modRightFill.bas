Attribute VB_Name = "modRightFill"
Option Explicit
Public Function RightFill(strInput As String, iLength As Integer) As String
    Dim i As Integer
    Dim s As String
    
    s = Trim(strInput)
    
    If Len(s) = iLength Then
        RightFill = s
    Else
        RightFill = String(iLength - Len(s), " ") & s
    End If
End Function
