Attribute VB_Name = "SYSGetComputerName"
'---insert this code in the declarations section of a module or class
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Function XGetComputerName() As String
    Dim Computer As String
    Dim BufSize As Long
    Dim RetCode As Long
    Dim NullCharPos As Long
     
    Computer = Space(80)
    BufSize = Len(Computer)
    
    '---call WINAPI
    RetCode = GetComputerName(Computer, BufSize)
    
    '---search for the end of the string
    NullCharPos = InStr(Computer, Chr(0))
    If NullCharPos > 0 Then
        Computer = Left(Computer, NullCharPos - 1)
    Else
        Computer = ""
    End If

    XGetComputerName = Computer
End Function

