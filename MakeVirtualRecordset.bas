Attribute VB_Name = "libMakeVirtualRecordset"
Option Explicit
Public Function MakeVirtualRecordset(ByRef ADOConnection As ADODB.Connection, RS As ADODB.Recordset, ByRef vRS As ADODB.Recordset, Optional HiddenFieldName As Variant) As Boolean
    Dim adoRS As New ADODB.Recordset
    Dim FieldList As String
    Dim TableList As String
    Dim WhereClause As String
    Dim OrderByClause As String
    Dim fld As ADODB.Field
    Dim iPos As Integer
    Dim SQLsource As String
    
    Call Trace(trcEnter, "MakeVirtualRecordset(ADOConnection, RS, vRS, """ & HiddenFieldName & """)")
    On Error GoTo ErrorHandler
    MakeVirtualRecordset = True
    
    'If the recordset has a filter on it already, SCR won't respect it, so include
    'it in the virtual recordset's Source...
    SQLsource = RS.Source
    ParseSQLSelect SQLsource, FieldList, TableList, WhereClause, OrderByClause
    If RS.Filter <> 0 And RS.Filter <> "" Then
        If WhereClause <> vbNullString Then WhereClause = WhereClause & " And "
        WhereClause = WhereClause & RS.Filter
    End If
    SQLsource = "Select " & FieldList & " From " & TableList
    If WhereClause <> vbNullString Then SQLsource = SQLsource & " Where " & WhereClause
    If OrderByClause <> vbNullString Then SQLsource = SQLsource & " Order By " & OrderByClause
    
    adoRS.Open SQLsource, ADOConnection, adOpenForwardOnly, adLockReadOnly
    If Not vRS Is Nothing Then
        CloseRecordset vRS, False
    Else
        Set vRS = New ADODB.Recordset
    End If
    
    For Each fld In adoRS.Fields
        vRS.Fields.Append fld.Name, fld.Type, fld.DefinedSize, fld.Attributes
    Next fld
    'Add the hidden field (assuming the value does not matter - usually used for Grids)...
    If Not IsMissing(HiddenFieldName) Then vRS.Fields.Append HiddenFieldName, adVarChar, 1
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
                vRS(fld.Name).Value = adoRS(fld.Name).Value
            Next fld
            vRS.Update
            adoRS.MoveNext
        Wend
        vRS.MoveFirst
    End If
    adoRS.Close
    Set adoRS = Nothing
    
    Call Trace(trcExit, "MakeVirtualRecordset")
    Exit Function
    
ErrorHandler:
    Dim errorCode As Long
    MakeVirtualRecordset = False
    MsgBox BuildADOerror(ADOConnection, errorCode), vbCritical, "MakeVirtualRecordset"
End Function

