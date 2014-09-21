Attribute VB_Name = "libParseStr"
'libParseStr - ParseStr.bas
'   Library ParseStr Module...
'Public domain, taken from "The Waite Group's Visual Basic Source Library"/SAMS Publishing...
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Problem:    Programmer:     Description:
'   02/25/99    None        Ken Clark       Incorporated into FiRRe;
'=================================================================================================================================
Option Explicit
Public Enum OpMode
  StringBinaryCompare = vbBinaryCompare + 1
  StringTextCompare = vbTextCompare + 1
  StringDataBaseCompare = vbDatabaseCompare + 1
  CharacterBinaryCompare = -(vbBinaryCompare + 1)
  CharacterTextCompare = -(vbTextCompare + 1)
  CharacterDataBaseCompare = -(vbDatabaseCompare + 1)
End Enum
Public Function ReplaceCS(ByVal strWork As String, ByVal strOld As String, ByVal strNew As String, Optional ByVal intOPMode As OpMode, Optional blnUpdated As Boolean) As String

  '*******************************************************************************
  '
  ' DESCRIPTION
  '     Replace a string or specific character(s) within a string. This routine
  '     can also be used to strip characters.
  '
  ' ARGUMENTS
  '     strWork    = String to work on.
  '
  '     strOld     = If intOPMode is negative, defines character(s) to replace.
  '                = If intOPMode is positive, defines string to replace.
  '
  '     strNew     = New character (or string) to substitute.
  '
  '     intOPMode  = Sets operation by defining the "replace" mode and "compare"
  '                  mode. Valid parameters are:
  '
  '                  StringBinaryCompare (Default if not specified.)
  '                  StringTextCompare
  '                  StringDataBaseCompare
  '                  CharacterBinaryCompare
  '                  CharacterTextCompare
  '                  CharacterDataBaseCompare
  '
  '     blnUpdated = Optional. Returns TRUE if string was modified
  '
  ' RETURNS
  '     Returns new string.
  '
  ' DEPENDENCIES
  '     None
  '
  ' REMARKS
  '     To strip a string or character(s), set strNew to vbNullString or "".
  '
  '
  '*******************************************************************************
  
  On Local Error Resume Next
  
  Dim intOldLen As Integer
  Dim intNewLen As Integer
  Dim intSPos As Long
  Dim intN As Integer
  
  If intOPMode = 0 Then
    intOPMode = StringBinaryCompare
  End If
  
  intNewLen = Len(strNew)
  intOldLen = Len(strOld)
    
  intSPos = 1
  blnUpdated = False
    
  If intOPMode < 0 Then
    intOPMode = Abs(intOPMode) - 1
    For intN = 1 To intOldLen
      intSPos = 1
      Do
        intSPos = InStr(intSPos, strWork, Mid(strOld, intN, 1), intOPMode)
        If intSPos Then
          strWork = Left(strWork, intSPos - 1) & strNew & Mid(strWork, intSPos + 1)
          intSPos = intSPos + intNewLen
          blnUpdated = True
        End If
      Loop While intSPos
    Next
  Else
    intOPMode = intOPMode - 1
    Do
      intSPos = InStr(intSPos, strWork, strOld, intOPMode)
      If intSPos Then
        strWork = Left(strWork, intSPos - 1) & strNew & Mid(strWork, intSPos + intOldLen)
        intSPos = intSPos + intNewLen
        blnUpdated = True
      End If
    Loop While intSPos
  End If
  
  ReplaceCS = strWork
  
End Function
Public Function TokenCount(strWork As String, strDelimiter As String) As Long

  '*******************************************************************************
  '
  ' DESCRIPTION
  '     Counts number of tokens in a string.  Function can also be used to
  '     determine the number of "delimiter" characters found in the string by
  '     subtracting 1 from the returned value.
  '
  ' ARGUMENTS
  '     strWork      = String to work on
  '     strDelimiter = String Delimiter
  '
  ' RETURNS
  '     Number of tokens found.
  '
  ' DEPENDENCIES
  '     None
  '
  ' REMARKS
  '
  '*******************************************************************************
  
  On Local Error Resume Next
  
  Dim lngN As Long
  Dim lngCPos As Long
  Dim lngSPos As Long
  Dim lngCharLen As Long
    
  If Len(strWork) = 0 Then Exit Function
  
  lngCharLen = Len(strDelimiter)
  lngSPos = 1
    
  Do
    lngCPos = InStr(lngSPos, strWork, strDelimiter)
    If lngCPos Then
      lngN = lngN + 1
      lngSPos = lngCharLen + lngCPos
    End If
  Loop While lngCPos
   
  If Right(strWork, lngCharLen) <> strDelimiter Then
    TokenCount = lngN + 1
  Else
    TokenCount = lngN
  End If
      
End Function
Public Function ParseStr(ByVal strWork As String, intTokenNum As Integer, strDelimitChr As String, Optional ByVal strEncapChr As String) As String

  '*******************************************************************************
  '
  ' DESCRIPTION
  '     Retrieve specified token of string.
  '
  ' ARGUMENTS
  '     strWork       = String to work on.
  '     intTokenNum   = If > 0,  returns specified token in string. If 0, returns
  '                     next token in string each time function is called. (If
  '                     no more tokens are found, function will return 0.) To
  '                     reset counter to 0, call routine as ParseStr ("", 0, "").
  '
  '     strDelimitChr = Token delimiter.
  '     strEncapChr   = Optional. Allows for tokens to return strings
  '                     encapsulated with "strDelimitChr" characters.
  '
  ' RETURNS
  '     Returns string token.  If none is found, will return "".
  '
  ' DEPENDENCIES
  '     ReplaceCS
  '
  ' REMARKS
  '
  '     If you are in "auto-mode" (intTokenNum=0) and are going to auto
  '     process another string, make sure to reset it as follows:
  '
  '     CALL ParseStr("", 0, "") or ParseStr "", 0, ""
  '
  '*******************************************************************************
    
  On Local Error Resume Next
  
  Dim blnExitDo As Boolean
  Dim intDPos As Integer
  Dim intSPtr As Integer
  Dim intEPtr As Integer
  Dim intCurrentTokenNum As Integer
  Dim intWorkStrLen As Integer
  Dim intEncapStatus As Integer
  Static intSPos As Integer
  Dim strTemp As String
  Static intDelimitLen As Integer

  intWorkStrLen = Len(strWork)
    
  If Len(strEncapChr) Then
    intEncapStatus = Len(strEncapChr)
  End If

  If intWorkStrLen = 0 Or (intSPos > intWorkStrLen And intTokenNum = 0) Then
    intSPos = 0
    Exit Function
  ElseIf intTokenNum > 0 Or intSPos = 0 Then
    intSPos = 1
    intDelimitLen = Len(strDelimitChr)
  End If

  Do
    
    intDPos = InStr(intSPos, strWork, strDelimitChr)

    If intEncapStatus Then
      intSPtr = InStr(intSPos, strWork, strEncapChr)
      intEPtr = InStr(intSPtr + 1, strWork, strEncapChr)
      If intDPos > intSPtr And intDPos < intEPtr Then
        intDPos = InStr(intEPtr, strWork, strDelimitChr)
      End If
    End If

    If intDPos < intSPos Then
      intDPos = intWorkStrLen + intDelimitLen
    End If

    If intDPos Then
      If intTokenNum Then
        intCurrentTokenNum = intCurrentTokenNum + 1
        If intCurrentTokenNum = intTokenNum Then
          strTemp = Mid(strWork, intSPos, intDPos - intSPos)
          blnExitDo = True
        Else
          blnExitDo = False
        End If
      Else
        strTemp = Mid(strWork, intSPos, intDPos - intSPos)
          blnExitDo = True
      End If
      intSPos = intDPos + intDelimitLen
    Else
      intSPos = 0
      blnExitDo = True
    End If
  Loop Until blnExitDo

  If intEncapStatus Then
    ParseStr = ReplaceCS(strTemp, strEncapChr, "", StringBinaryCompare)
  Else
    ParseStr = strTemp
  End If

End Function


