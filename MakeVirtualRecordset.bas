Attribute VB_Name = "libMakeVirtualRecordset"
'libMakeVirtualRecordset - libMakeVirtualRecordset.bas
'   Library Module Handling Creation of Virtual ADODB Recordsets...
'   Copyright © 1999-2002, Ken Clark
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Description:
'   08/20/02    Started History;
'=================================================================================================================================
Option Explicit
Public Function MakeVirtualRecordsetFromRS(ByRef ADOConnection As ADODB.Connection, RS As ADODB.Recordset, ByRef vRS As ADODB.Recordset, Optional HiddenFieldName As Variant) As Boolean
    Dim adoRS As New ADODB.Recordset
    Dim FieldList As String
    Dim TableList As String
    Dim WhereClause As String
    Dim OrderByClause As String
    Dim fld As ADODB.Field
    Dim iPos As Integer
    Dim SQLSource As String
    
    On Error GoTo ErrorHandler
    If IsMissing(HiddenFieldName) Then
        Call Trace(trcEnter, "MakeVirtualRecordsetFromRS(ADOConnection, RS, vRS, , )")
    Else
        Call Trace(trcEnter, "MakeVirtualRecordsetFromRS(ADOConnection, RS, vRS, """ & HiddenFieldName & """)")
    End If
    MakeVirtualRecordsetFromRS = True
    
    'If the recordset has a filter on it already, SCR won't respect it, so include
    'it in the virtual recordset's Source...
    SQLSource = RS.Source
    ParseSQLSelect SQLSource, FieldList, TableList, WhereClause, OrderByClause
    If RS.Filter <> 0 And RS.Filter <> "" Then
        If WhereClause <> vbNullString Then WhereClause = WhereClause & " And "
        WhereClause = WhereClause & RS.Filter
    End If
    SQLSource = "Select " & FieldList & " From " & TableList
    If WhereClause <> vbNullString Then SQLSource = SQLSource & " Where " & WhereClause
    If OrderByClause <> vbNullString Then SQLSource = SQLSource & " Order By " & OrderByClause
    
    adoRS.Open SQLSource, ADOConnection, adOpenForwardOnly, adLockReadOnly
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
    
    Call Trace(trcExit, "MakeVirtualRecordsetFromRS")
    Exit Function
    
ErrorHandler:
    Dim errorCode As Long
    MakeVirtualRecordsetFromRS = False
    MsgBox BuildADOerror(ADOConnection, errorCode), vbCritical, "MakeVirtualRecordsetFromRS"
End Function
Public Function MakeVirtualRecordsetFromSQL(ByRef ADOConnection As ADODB.Connection, ByVal SQLSource As String, ByRef vRS As ADODB.Recordset, Optional HiddenFieldName As Variant) As Boolean
    Dim adoRS As New ADODB.Recordset
    Dim FieldList As String
    Dim TableList As String
    Dim WhereClause As String
    Dim OrderByClause As String
    Dim fld As ADODB.Field
    Dim iPos As Integer
    Dim SQLSource As String
    
    On Error GoTo ErrorHandler
    If IsMissing(HiddenFieldName) Then
        Call Trace(trcEnter, "MakeVirtualRecordsetFromSQL(ADOConnection, RS, vRS, , )")
    Else
        Call Trace(trcEnter, "MakeVirtualRecordsetFromSQL(ADOConnection, RS, vRS, """ & HiddenFieldName & """)")
    End If
    MakeVirtualRecordsetFromSQL = True
    
    'If the recordset has a filter on it already, SCR won't respect it, so include
    'it in the virtual recordset's Source...
    SQLSource = RS.Source
    ParseSQLSelect SQLSource, FieldList, TableList, WhereClause, OrderByClause
    If RS.Filter <> 0 And RS.Filter <> "" Then
        If WhereClause <> vbNullString Then WhereClause = WhereClause & " And "
        WhereClause = WhereClause & RS.Filter
    End If
    SQLSource = "Select " & FieldList & " From " & TableList
    If WhereClause <> vbNullString Then SQLSource = SQLSource & " Where " & WhereClause
    If OrderByClause <> vbNullString Then SQLSource = SQLSource & " Order By " & OrderByClause
    
    adoRS.Open SQLSource, ADOConnection, adOpenForwardOnly, adLockReadOnly
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
    
    Call Trace(trcExit, "MakeVirtualRecordsetFromSQL")
    Exit Function
    
ErrorHandler:
    Dim errorCode As Long
    MakeVirtualRecordsetFromSQL = False
    MsgBox BuildADOerror(ADOConnection, errorCode), vbCritical, "MakeVirtualRecordsetFromSQL"
End Function


