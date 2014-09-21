Attribute VB_Name = "libMakeVirtualRecordset"
Option Explicit
Public Function MakeVirtualRecordset(ByRef ADOConnection As ADODB.Connection, RS As ADODB.Recordset, ByRef vRS As ADODB.Recordset) As Boolean
    Dim adoRS As New ADODB.Recordset
    Dim fld As ADODB.Field
    Dim iPos As Integer
    Dim SQLsource As String
    
    On Error GoTo ErrorHandler
    MakeVirtualRecordset = True
    
    'If the recordset has a filter on it already, SCR won't respect it, so include
    'it in the virtual recordset's Source...
    If RS.Filter <> 0 And RS.Filter <> "" Then
        iPos = InStr(UCase(RS.Source), " WHERE ")
        If iPos > 0 Then
            SQLsource = Mid(RS.Source, 1, iPos + Len(" WHERE ") - 1) & RS.Filter & " and " & Mid(RS.Source, iPos + Len(" WHERE "))
        Else
            iPos = InStr(UCase(RS.Source), " ORDER BY ")
            If iPos > 0 Then
                SQLsource = Mid(RS.Source, 1, iPos - 1) & " where " & RS.Filter & Mid(RS.Source, iPos)
            End If
        End If
    Else
        SQLsource = RS.Source
    End If
    adoRS.Open SQLsource, ADOConnection, adOpenForwardOnly, adLockReadOnly
        
    If Not vRS Is Nothing Then
        On Error Resume Next
        If vRS.State = adStateOpen Then vRS.Close
        Set vRS = Nothing
        On Error GoTo ErrorHandler
    End If
    Set vRS = New ADODB.Recordset ' Set the object variable.
    
    For Each fld In adoRS.Fields
        vRS.Fields.Append fld.Name, fld.Type, fld.DefinedSize, fld.Attributes
    Next fld
    vRS.CursorType = adOpenStatic    'Updatable snapshot
    vRS.LockType = adLockOptimistic  'Allow updates
    vRS.Open
    
    'Copy the data from the real recordset to the virtual one...
    If Not (adoRS.BOF And adoRS.EOF) Then
        adoRS.MoveFirst
        While Not adoRS.EOF
            'Populate the grid with the recordset data...
            vRS.AddNew
            For Each fld In adoRS.Fields
                vRS(fld.Name) = adoRS(fld.Name)
            Next fld
            vRS.Update
            adoRS.MoveNext
        Wend
        vRS.MoveFirst
    End If
    adoRS.Close
    Set adoRS = Nothing
    
    Exit Function
    
ErrorHandler:
    Dim errorCode As Long
    MakeVirtualRecordset = False
    MsgBox BuildADOerror(ADOConnection, errorCode), vbCritical, "MakeVirtualRecordset"
End Function

