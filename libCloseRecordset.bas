Attribute VB_Name = "libCloseRecordset"
Option Explicit
Public Sub CloseRecordset(adoRS As ADODB.Recordset, Destroy As Boolean)
    'On Error Resume Next
    If Not adoRS Is Nothing Then
        If (adoRS.State And adStateOpen) = adStateOpen Then
            If Not (adoRS.EOF Or adoRS.BOF) Then
                If adoRS.EditMode <> adEditNone Then adoRS.CancelUpdate
            End If
            adoRS.Close
        End If
        If Destroy Then Set adoRS = Nothing
    End If
End Sub

