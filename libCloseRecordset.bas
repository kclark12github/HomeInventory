Attribute VB_Name = "libCloseRecordset"
Option Explicit
Public Sub CloseRecordset(adoRS As ADODB.Recordset, Destroy As Boolean)
    'On Error Resume Next
    If Not adoRS Is Nothing Then
        If (adoRS.State And adStateOpen) = adStateOpen Then
            adoRS.CancelUpdate
            adoRS.Close
        End If
        If Destroy Then Set adoRS = Nothing
    End If
End Sub

