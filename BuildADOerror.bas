Attribute VB_Name = "libBuildADOerror"
Option Explicit
Public Function BuildADOerror(ByRef cn As ADODB.Connection, ByRef errorCode As Long) As String
    Dim adoError As ADODB.Error
    BuildADOerror = ""
    errorCode = 0
    For Each adoError In cn.Errors
        errorCode = adoError.NativeError             'For additional error processing by caller...
        If Trim(adoError.Description) = "" Then
            BuildADOerror = BuildADOerror & "System Error (" & Hex(adoError.Number) & ")" & vbCr
        Else
            BuildADOerror = BuildADOerror & adoError.Description & "(" & Hex(adoError.Number) & ")" & vbCr
        End If
        BuildADOerror = BuildADOerror & vbTab & "Source: " & adoError.Source & vbCr & _
            vbTab & "SQL State: " & adoError.SQLState & vbCr & _
            vbTab & "Native Error: " & adoError.NativeError & vbCr
        If adoError.HelpFile = "" Then
            BuildADOerror = BuildADOerror & vbCr & vbTab & "No Help file available"
        Else
            BuildADOerror = BuildADOerror & vbTab & "HelpFile: " & adoError.HelpFile & vbCr & _
                vbTab & "HelpContext: " & adoError.HelpContext
        End If
        BuildADOerror = BuildADOerror & vbCr & vbCr
    Next
    If Trim(BuildADOerror) = "" Then
        BuildADOerror = "Unable to determine error, no ADO errors registered." & vbCr & vbCr
    End If
End Function

