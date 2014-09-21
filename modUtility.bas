Attribute VB_Name = "modUtility"
Option Explicit

' Trace Objects...
Public Enum trcLevel
    trcEnter = -1
    trcBody = 0
    trcExit = 1
End Enum
Public gfTraceMode As Boolean
Public Sub Trace(trcTraceLevel As trcLevel, strMessage As String)
    Static indentLevel As Integer
    Dim i As Integer
    Dim strTabs As String
    Dim strTraceMessage As String
    
    If Not gfTraceMode Then Exit Sub
    
    If trcTraceLevel = trcExit Then indentLevel = indentLevel - 1
    For i = 1 To indentLevel
        strTabs = strTabs & vbTab
    Next i
    
    If trcTraceLevel = trcEnter Then
        strTraceMessage = Now & ": " & strTabs & "Entering " & strMessage
    ElseIf trcTraceLevel = trcExit Then
        strTraceMessage = Now & ": " & strTabs & "Exiting " & strMessage
    Else
        strTraceMessage = Now & ": " & strTabs & strMessage
    End If
    
    Debug.Print strTraceMessage
    'Print #gintTraceUnit, strTraceMessage

    If trcTraceLevel = trcEnter Then indentLevel = indentLevel + 1
End Sub
Public Function SQLQuote(strSource) As String
    Dim i As Integer
    Dim strTemp As String
    
    strTemp = strSource
    'Eliminate all double - single quotes
    Do
        i = InStr(1, strTemp, "''")
        If i <> 0 Then
            strTemp = Left(strTemp, i - 1) & "'" & Mid(strTemp, i + 2)
        End If
    Loop Until i = 0
    'Replace all single quotes with double - single quotes
    i = InStr(1, strTemp, "'")
    Do While i > 0
        strTemp = Left(strTemp, i - 1) & "''" & Mid(strTemp, i + 1)
        i = InStr(i + 2, strTemp, "'")
    Loop
    SQLQuote = strTemp
End Function
Public Function URLencode(vURL As Variant) As Variant
    Dim i As Integer
    
    URLencode = vURL
    If vURL = "" Then
        URLencode = Null
        Exit Function
    End If
    
    While True
        i = InStr(URLencode, " ")
        If i > 0 Then
            URLencode = Mid(URLencode, 1, i - 1) & "%20" & Mid(URLencode, i + 1)
        Else
            Exit Function
        End If
    Wend
End Function
Public Function URLdecode(vURL As Variant) As String
    Dim i As Integer
    
    If IsNull(vURL) Then
        URLdecode = ""
        Exit Function
    End If
    
    URLdecode = vURL
    While True
        i = InStr(URLdecode, "%20")
        If i > 0 Then
            URLdecode = Mid(URLdecode, 1, i - 1) & " " & Mid(URLdecode, i + 3)
        Else
            Exit Function
        End If
    Wend
End Function
Public Function VBencode(vString As Variant) As Variant
    Dim i As Integer
    Dim Start As Integer
    
    VBencode = vString
    If vString = "" Then
        VBencode = Null
        Exit Function
    End If
    
    Start = 1
    While True
        i = InStr(Start, VBencode, "&")
        If i > 0 Then
            VBencode = Mid(VBencode, 1, i - 1) & "&&" & Mid(VBencode, i + 1)
            Start = i + 2
        Else
            Exit Function
        End If
    Wend
End Function
Public Function VBdecode(vString As Variant) As String
    Dim i As Integer
    
    If IsNull(vString) Then
        VBdecode = ""
        Exit Function
    End If
    
    VBdecode = vString
    While True
        i = InStr(VBdecode, "&&")
        If i > 0 Then
            VBdecode = Mid(VBdecode, 1, i - 1) & Mid(VBdecode, i + 1)
        Else
            Exit Function
        End If
    Wend
End Function
Public Function IsButton(ByRef Node As ComctlLib.Node) As Boolean
    IsButton = False
    If Node Is Nothing Then Exit Function
    If UCase(Left(Node.Tag, 7)) = "BUTTON:" Then IsButton = True
End Function
Public Function IsGroup(ByRef Node As ComctlLib.Node) As Boolean
    IsGroup = False
    If Node Is Nothing Then Exit Function
    If UCase(Left(Node.Tag, 6)) = "GROUP:" Then IsGroup = True
End Function
Public Function IsLink(ByRef Node As ComctlLib.Node) As Boolean
    IsLink = False
    If Node Is Nothing Then Exit Function
    If UCase(Left(Node.Tag, 5)) = "LINK:" Then IsLink = True
End Function
