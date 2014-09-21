Attribute VB_Name = "libISBN"
'Algorithm taken from web site:
'http://www.isbn.org/standards/home/isbn/international/hyphenation-instructions.asp
Option Explicit
Public Function FormatISBN(sISBN As String) As String
    Dim s As String
    Dim Group As String
    
    s = Replace(sISBN, "-", vbNullString)
    Group = Left(s, 1)
    Select Case Group
        Case "0"
            Select Case Mid(s, 2, 2)
                Case "00" To "19"
                    FormatISBN = Format(s, "&-&&-&&&&&&-&")
                Case "20" To "69"
                    FormatISBN = Format(s, "&-&&&-&&&&&-&")
                Case "70" To "84"
                    FormatISBN = Format(s, "&-&&&&-&&&&-&")
                Case "85" To "89"
                    FormatISBN = Format(s, "&-&&&&&-&&&-&")
                Case "90" To "94"
                    FormatISBN = Format(s, "&-&&&&&&-&&-&")
                Case "95" To "99"
                    FormatISBN = Format(s, "&-&&&&&&&-&-&")
            End Select
        Case "1"
            Select Case Mid(s, 2, 2)
                Case "00" To "09"
                    FormatISBN = Format(s, "&-&&-&&&&&&-&")
                Case "10" To "39"
                    FormatISBN = Format(s, "&-&&&-&&&&&-&")
                Case "40" To "54"
                    FormatISBN = Format(s, "&-&&&&-&&&&-&")
                Case Else
                    Select Case Mid(s, 2, 4)
                        Case "5500" To "8697"
                            FormatISBN = Format(s, "&-&&&&&-&&&-&")
                        Case "8698" To "9989"
                            FormatISBN = Format(s, "&-&&&&&&-&&-&")
                        Case "9990" To "9999"
                            FormatISBN = Format(s, "&-&&&&&&&-&-&")
                    End Select
            End Select
    End Select

ExitSub:
    Exit Function
End Function
