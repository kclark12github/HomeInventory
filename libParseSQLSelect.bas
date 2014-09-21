Attribute VB_Name = "libParseSQLSelect"
Option Explicit
Public Sub ParseSQLSelect(ByVal SQLstatement As String, FieldList As String, TableList As String, WhereClause As String, OrderByClause As String)
    Dim iFrom As Integer
    Dim iWhere As Integer
    Dim iOrderBy As Integer
    
    FieldList = vbNullString
    TableList = vbNullString
    WhereClause = vbNullString
    OrderByClause = vbNullString
    
    'Gotta have a SELECT and FROM...
    SQLstatement = Trim(SQLstatement)
    If UCase(Mid(SQLstatement, 1, 6)) <> "SELECT" Then Exit Sub
    SQLstatement = Trim(Right(SQLstatement, Len(SQLstatement) - Len("SELECT")))
    iFrom = InStr(UCase(SQLstatement), " FROM ")
    If iFrom = 0 Then Exit Sub
    
    'Parse FieldList...
    FieldList = Trim(Left(SQLstatement, iFrom))
    SQLstatement = Trim(Mid(SQLstatement, iFrom + Len(" FROM ")))
    iWhere = InStr(UCase(SQLstatement), " WHERE ")
    
    'Parse TableList...
    If iWhere = 0 Then
        iOrderBy = InStr(UCase(SQLstatement), " ORDER BY ")
        If iOrderBy = 0 Then
            TableList = Trim(SQLstatement)
            Exit Sub
        Else
            TableList = Trim(Left(SQLstatement, iOrderBy))
            SQLstatement = Trim(Mid(SQLstatement, iOrderBy + Len(" ORDER BY ")))
            OrderByClause = Trim(SQLstatement)
            Exit Sub
        End If
    Else
        TableList = Trim(Left(SQLstatement, iWhere))
        SQLstatement = Trim(Mid(SQLstatement, iWhere + Len(" WHERE ")))
        iOrderBy = InStr(UCase(SQLstatement), " ORDER BY ")
        
        If iOrderBy = 0 Then
            WhereClause = Trim(SQLstatement)
            Exit Sub
        Else
            WhereClause = Trim(Left(SQLstatement, iOrderBy))
            SQLstatement = Trim(Mid(SQLstatement, iOrderBy + Len(" ORDER BY ")))
            OrderByClause = Trim(SQLstatement)
            Exit Sub
        End If
    End If
End Sub

