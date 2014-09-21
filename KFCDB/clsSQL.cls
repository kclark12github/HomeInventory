VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'clsSQL - clsSQL.cls
'   SQL Module...
'   Copyright © 2002, SunGard Investor Accounting Systems
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Problem:    Programmer:     Description:
'   12/13/01    None        JAD/Ken Clark   Corrected implementation of ALTER TABLE DROP COLUMN <ColumnName> for SQL Server;
'   12/03/01    None        JAD             PRIMARY KEY CONSTRAINT vs INDEX
'   10/30/01    None        Ken Clark       Added dberrInvalidObject support in ExecuteSQL;
'   10/25/01    None        JAD             Excluded Temporary Report tables (non-dbparmOwner) from XSYS tables
'   09/28/01    58219       Ken Clark       Replaced initDBErrors with initDBEngineParms;
'                                           Replaced strSQL* codes with modDBEngine objDBE.dbparm* values;
'   09/26/01    58219       Ken Clark       Made various supporting table owner properties;
'   09/25/01    58219       Ken Clark       Removed 'ByVal' from doSQL's SQLSource argument so error messages can be returned;
'                                           SQL Server owner corrections ('SYSADM' vs 'dbo' case sensitivity issues);
'                                           Added use of vrsXSYSColumns to cache REMARKS when necessary;
'   09/23/01    58219       Ken Clark       Added logic to doCreateTable to properly handle Owner;
'   09/21/01    58219       Ken Clark       Moved Connect, Disconnect, and ExecuteSQL from modMain;
'   09/07/01    None        Ken Clark       Implemented Refresh Schema command for SQL Server;
'                                           Renamed remaining XSYS-table indexes from FIRRE to XSYS prefixes;
'   09/01/01    None        Ken Clark       Added DBENgine argument to doDataType calls;
'   08/31/01    None        Ken Clark       Removed specification of precision in REAL data types;
'   07/21/01    None        Ken Clark       Consolidated FIRRE table management through doSQL;
'   07/20/01    None        Ken Clark       Created;
'=================================================================================================================================
Option Explicit
Private mConstants As SIASDB.clsSIASDBConstants
Private objDBE As SIASDB.clsDBEngine
Private objUI As SIASUTL.clsUI

Public vrsXSYSTables As ADODB.Recordset
Public vrsXSYSColumns As ADODB.Recordset
Private Sub Class_Initialize()
    Set mConstants = New SIASDB.clsSIASDBConstants
    Set objDBE = New SIASDB.clsDBEngine
    Set objUI = New SIASUTL.clsUI
    Set vrsXSYSTables = Nothing
    Set vrsXSYSColumns = Nothing
End Sub
Private Sub Class_Terminate()
    Set mConstants = Nothing
    Set objDBE = Nothing
    Set objUI = Nothing
    Set vrsXSYSTables = Nothing
    Set vrsXSYSColumns = Nothing
End Sub
Public Function Connect(DSN As String) As Boolean
    Dim ErrorCode As Long
    Dim ErrorString As String
    Dim i As Integer
    Dim saveErrNumber As Long
    Dim saveErrString As String
    Dim saveMousePointer As Integer
    
    On Error Resume Next
    Connect = False
    objDBE.ClearError
    Call objDBE.ConnectByFileDSN(DSN, vbNullString, vbNullString)
    
    'Set DBEngine-specific parameters (to be used throughout this utility)...
    Call objDBE.initDBEngineParms(objDBE.DBEngine)

    Connect = True
    
ExitSub:
    Exit Function
    
ErrorHandler:
    ErrorCode = objDBE.LastErrorCode
    ErrorString = objDBE.LastErrorString
    If ErrorCode = 0 And ErrorString = vbNullString Then
        ErrorString = BuildADOerror(objDBE.ActiveConnection, ErrorCode)
        If ErrorCode = 0 Then
            ErrorCode = err.Number
            ErrorString = err.Description
        End If
    End If
    
    Select Case ErrorCode
        Case objDBE.dberrDBLicenseExceeded
            ErrorString = "Maximum number of clients exceeded; try again later."
        Case objDBE.dberrInvalidUserID
            ErrorString = "The supplied UserID is not valid; try again."
        Case objDBE.dberrInvalidPassword
            ErrorString = "The supplied Password is not valid; try again."
        Case objDBE.dberrDBNotConfigured
            ErrorString = _
                "Database was not found in the SQL.INI file on this PC. " & vbCrLf & vbCrLf & _
                "Please contact your Help Desk or System Administrator to verify your SQL.INI file settings."
        Case Else
            If ErrorCode <> 0 Or saveErrString = vbNullString Then
                ErrorString = "Error encountered connecting to the " & objDBE.DatabaseName & " database." & vbCr & vbCr & _
                    "Connection String: """ & objDBE.ActiveConnection.ConnectionString & """" & vbCr & vbCr & _
                    ErrorString
            Else
                ErrorString = "Error encountered connecting to the " & objDBE.DatabaseName & " database." & vbCr & vbCr & _
                    saveErrString & " (Error #" & saveErrNumber & ")"
            End If
    End Select
    App.LogEvent "Connection to " & objDBE.ServerName & "\" & objDBE.DatabaseName & " failed. UserID: " & objDBE.UserID & ";" & ErrorString, vbLogEventTypeError
    Call objDBE.Disconnect(True)
    err.Raise ErrorCode, mConstants.App.EXEName & " (clsSQL.Connect)", ErrorString
    Resume Next
End Function
Public Function CreateFilesBookTables() As Boolean
    Dim SQLsource As String
    Dim ErrorCode As Long
    Dim fNeedToDrop As Boolean
    
    Call Trace(trcEnter, "CreateFilesBookTables()", trcSQL)
    
    CreateFilesBookTables = False
    
    'Before we get started, see if the FiRRe meta-data tables exist...
    On Error Resume Next
    Set vrsXSYSTables = New ADODB.Recordset
    fNeedToDrop = False
    vrsXSYSTables.Open "Select NAME From XSYSTABLES Where NAME='XSYSTABLES' And OWNER='" & objDBE.dbparmOwner & "'", objDBE.ActiveConnection, adOpenKeyset, adLockReadOnly
    ErrorCode = err.Number
    err.Clear
    Call CloseRecordset(vrsXSYSTables, True)
    If ErrorCode = 0 Then
        'Cache any existing default record information and store for later use...
        Set vrsXSYSTables = New ADODB.Recordset
        Call objDBE.MakeVirtualRecordset("Select * From XSYSTABLES Where OWNER='" & objDBE.dbparmOwner & "' Order By NAME", vrsXSYSTables)
        fNeedToDrop = True
        'Cache any existing Remarks information and store for later use...
        Set vrsXSYSColumns = New ADODB.Recordset
        Call objDBE.MakeVirtualRecordset("Select * From XSYSCOLUMNS Where TBOWNER='" & objDBE.dbparmOwner & "' Order By TBNAME,NAME", vrsXSYSColumns)
    End If
    
    Call objUI.Display(vbTab & "XSYSTABLES", deBothStatusBarAndTextBox)
    If Not ExecuteSQL("DROP TABLE " & objDBE.dbparmOwner & "." & "XSYSTABLES") Then GoTo ExitSub
    If Not ExecuteSQL( _
            "CREATE TABLE " & objDBE.dbparmOwner & "." & "XSYSTABLES(" & _
                "NAME VARCHAR(80)," & _
                "COLCOUNT SMALLINT," & _
                "TYPE CHAR(1)," & _
                "SYSTEM CHAR(1)," & _
                "SNUM INTEGER," & _
                "DATABASE_NAME VARCHAR(80)," & _
                "OWNER VARCHAR(80)," & _
                "DEFAULT_FIELD VARCHAR(80)," & _
                "DEFAULT_MIN INTEGER," & _
                "DEFAULT_MAX INTEGER," & _
                "REMARKS VARCHAR(254))") Then GoTo ExitSub
    If Not ExecuteSQL( _
            "CREATE UNIQUE INDEX XSYSTABLESBYNAME ON " & objDBE.dbparmOwner & "." & "XSYSTABLES (" & _
                "NAME" & objDBE.dbparmAsc & ", " & _
                "OWNER" & objDBE.dbparmAsc & ", " & _
                "SNUM" & objDBE.dbparmAsc & ")") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSTABLES (NAME, COLCOUNT, REMARKS, TYPE, SYSTEM, SNUM, DATABASE_NAME, OWNER, DEFAULT_FIELD, DEFAULT_MIN, DEFAULT_MAX) Values ('XSYSTABLES',6,'Documented database table information to be displayed in the FilesBook.rpt','T','Y',1,'" & objDBE.DatabaseName & "','" & objDBE.dbparmOwner & "',NULL,NULL,NULL)") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSTABLES (NAME, COLCOUNT, REMARKS, TYPE, SYSTEM, SNUM, DATABASE_NAME, OWNER, DEFAULT_FIELD, DEFAULT_MIN, DEFAULT_MAX) Values ('XSYSCOLUMNS',7,'Documented database column information to be displayed in the FilesBook.rpt','T','Y',2,'" & objDBE.DatabaseName & "','" & objDBE.dbparmOwner & "',NULL,NULL,NULL)") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSTABLES (NAME, COLCOUNT, REMARKS, TYPE, SYSTEM, SNUM, DATABASE_NAME, OWNER, DEFAULT_FIELD, DEFAULT_MIN, DEFAULT_MAX) Values ('XSYSINDEXES',5,'Documented database index information to be displayed in the FilesBook.rpt','T','Y',3,'" & objDBE.DatabaseName & "','" & objDBE.dbparmOwner & "',NULL,NULL,NULL)") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSTABLES (NAME, COLCOUNT, REMARKS, TYPE, SYSTEM, SNUM, DATABASE_NAME, OWNER, DEFAULT_FIELD, DEFAULT_MIN, DEFAULT_MAX) Values ('XSYSKEYS',5,'Documented database key detail information to be displayed in the FilesBook.rpt','T','Y',4,'" & objDBE.DatabaseName & "','" & objDBE.dbparmOwner & "',NULL,NULL,NULL)") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSTABLES (NAME, COLCOUNT, REMARKS, TYPE, SYSTEM, SNUM, DATABASE_NAME, OWNER, DEFAULT_FIELD, DEFAULT_MIN, DEFAULT_MAX) Values ('XSYSCONSTRAINTS',4,'Documented database primary/foreign key information to be displayed in the FilesBook.rpt','T','Y',5,'" & objDBE.DatabaseName & "','" & objDBE.dbparmOwner & "',NULL,NULL,NULL)") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSTABLES (NAME, COLCOUNT, REMARKS, TYPE, SYSTEM, SNUM, DATABASE_NAME, OWNER, DEFAULT_FIELD, DEFAULT_MIN, DEFAULT_MAX) Values ('XSYSFKCONSTRAINTS',6,'Documented database foreign key detail information to be displayed in the FilesBook.rpt','T','Y',6,'" & objDBE.DatabaseName & "','" & objDBE.dbparmOwner & "',NULL,NULL,NULL)") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSTABLES (NAME, COLCOUNT, REMARKS, TYPE, SYSTEM, SNUM, DATABASE_NAME, OWNER, DEFAULT_FIELD, DEFAULT_MIN, DEFAULT_MAX) Values ('XSYSPKCONSTRAINTS',3,'Documented database primary key detail information to be displayed in the FilesBook.rpt','T','Y',7,'" & objDBE.DatabaseName & "','" & objDBE.dbparmOwner & "',NULL,NULL,NULL)") Then GoTo ExitSub
    
    Call objUI.Display(vbTab & "XSYSCOLUMNS", deBothStatusBarAndTextBox)
    If Not ExecuteSQL("DROP TABLE " & objDBE.dbparmOwner & "." & "XSYSCOLUMNS") Then GoTo ExitSub
    If Not ExecuteSQL( _
            "CREATE TABLE " & objDBE.dbparmOwner & "." & "XSYSCOLUMNS(" & _
                "NAME VARCHAR(80)," & _
                "TBNAME VARCHAR(80)," & _
                "TBOWNER VARCHAR(80)," & _
                "COLNO SMALLINT," & _
                "COLTYPE VARCHAR(32)," & _
                "STORAGE SMALLINT," & _
                "LENGTH SMALLINT," & _
                "SCALE SMALLINT," & _
                "NULLS CHAR(1)," & _
                "REMARKS VARCHAR(254))") Then GoTo ExitSub
    If Not ExecuteSQL( _
            "CREATE UNIQUE INDEX XSYSCOLUMNSBYNAME ON " & objDBE.dbparmOwner & "." & "XSYSCOLUMNS (" & _
                "TBNAME" & objDBE.dbparmAsc & ", " & _
                "TBOWNER" & objDBE.dbparmAsc & ", " & _
                "NAME" & objDBE.dbparmAsc & ")") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSTABLES','NAME','" & objDBE.dbparmOwner & "',1,'VARCHAR',80,80,0,'Y','Table Name')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSTABLES','COLCOUNT','" & objDBE.dbparmOwner & "',2,'SMALLINT',2,0,0,'Y','Number of Columns in the Table')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSTABLES','TYPE','" & objDBE.dbparmOwner & "',3,'CHAR',1,1,0,'Y','Type of Table; V = View; T = Table')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSTABLES','SYSTEM','" & objDBE.dbparmOwner & "',4,'CHAR',1,1,0,'Y','Does this table represent a System Table? (Y/N)')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSTABLES','SNUM','" & objDBE.dbparmOwner & "',5,'INTEGER',4,0,0,'Y','Ordinal number of table (based on order tables are created)')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSTABLES','DATABASE_NAME','" & objDBE.dbparmOwner & "',6,'VARCHAR',80,80,0,'Y','Database name (used in instances where a full table specification is required (i.e. <database>.<owner>.<table>).')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSTABLES','OWNER','" & objDBE.dbparmOwner & "',7,'VARCHAR',80,80,0,'Y','Owner/Creator of this table.')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSTABLES','DEFAULT_FIELD','" & objDBE.dbparmOwner & "',8,'VARCHAR',80,80,0,'Y','Key field to be used to identify an initial set of records when creating new databases based on an existing schema (Null indicates no initial records).')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSTABLES','DEFAULT_MIN','" & objDBE.dbparmOwner & "',9,'INTEGER',4,0,0,'Y','Minimum key value used in identifying the range of initial records when creating new databases based on an existing schema (Null denotes no set minimum range).')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSTABLES','DEFAULT_MAX','" & objDBE.dbparmOwner & "',10,'INTEGER',4,0,0,'Y','Maximum key value used in identifying the range of initial records when creating new databases based on an existing schema (Null denotes no set maximum range).')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSTABLES','REMARKS','" & objDBE.dbparmOwner & "',11,'VARCHAR',254,254,0,'Y','Table Description')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSCOLUMNS','NAME','" & objDBE.dbparmOwner & "',1,'VARCHAR',80,80,0,'Y','Column or Field Name')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSCOLUMNS','TBNAME','" & objDBE.dbparmOwner & "',2,'VARCHAR',80,80,0,'Y','Parent Table Name')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSCOLUMNS','TBOWNER','" & objDBE.dbparmOwner & "',3,'VARCHAR',80,80,0,'Y','Parent Table Owner')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSCOLUMNS','COLNO','" & objDBE.dbparmOwner & "',4,'SMALLINT',2,0,0,'Y','Ordinal number of Column or Field')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSCOLUMNS','COLTYPE','" & objDBE.dbparmOwner & "',5,'VARCHAR',32,32,0,'Y','Data Type of Column or Field Name')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSCOLUMNS','STORAGE','" & objDBE.dbparmOwner & "',6,'SMALLINT',2,0,0,'Y','Size in Bytes of data contained in Column or Field Name')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSCOLUMNS','LENGTH','" & objDBE.dbparmOwner & "',7,'SMALLINT',2,0,0,'Y','Number of characters in string data or total decimal digits in decimal data')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSCOLUMNS','SCALE','" & objDBE.dbparmOwner & "',8,'SMALLINT',2,0,0,'Y','Number of decimal digits stored to the right of the decimal point in decimal data')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSCOLUMNS','NULLS','" & objDBE.dbparmOwner & "',9,'CHAR',1,1,0,'Y','Does Column or Field Name allow Null Values? (Y/N)')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSCOLUMNS','REMARKS','" & objDBE.dbparmOwner & "',10,'VARCHAR',254,254,0,'Y','Column or Field Name Description')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSINDEXES','NAME','" & objDBE.dbparmOwner & "',1,'VARCHAR',80,80,0,'Y','Index Name')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSINDEXES','TBNAME','" & objDBE.dbparmOwner & "',2,'VARCHAR',80,80,0,'Y','Parent Table Name')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSINDEXES','TBOWNER','" & objDBE.dbparmOwner & "',3,'VARCHAR',80,80,0,'Y','Parent Table Owner')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSINDEXES','UNIQUERULE','" & objDBE.dbparmOwner & "',4,'CHAR',1,1,0,'Y','Does this index allow duplicate entries? (D = Duplicates; U = Unique)')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSINDEXES','COLCOUNT','" & objDBE.dbparmOwner & "',5,'SMALLINT',2,0,0,'Y','Number of Columns contained in the index')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSINDEXES','SYSTEM','" & objDBE.dbparmOwner & "',6,'CHAR',1,1,0,'Y','Does this index represent a System Index? (Y/N)')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSKEYS','IXNAME','" & objDBE.dbparmOwner & "',1,'VARCHAR',80,80,0,'Y','Parent index name')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSKEYS','COLNAME','" & objDBE.dbparmOwner & "',2,'VARCHAR',80,80,0,'Y','Name of Column name represented by this key segment')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSKEYS','COLNO','" & objDBE.dbparmOwner & "',3,'SMALLINT',2,0,0,'Y','Ordinal number Column represented by this key segment')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSKEYS','COLSEQ','" & objDBE.dbparmOwner & "',4,'SMALLINT',2,0,0,'Y','Ordinal number Column represented by this key segment')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSKEYS','ORDERING','" & objDBE.dbparmOwner & "',5,'CHAR',1,1,0,'Y','Is this key ascending or descending? (A/D)')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSCONSTRAINTS','TBNAME','" & objDBE.dbparmOwner & "',1,'VARCHAR',80,80,0,'Y','Table name of the referring column or field')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSCONSTRAINTS','TBOWNER','" & objDBE.dbparmOwner & "',2,'VARCHAR',80,80,0,'Y','Table owner of the referring column or field')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSCONSTRAINTS','CONSTRAINTNAME','" & objDBE.dbparmOwner & "',3,'VARCHAR',80,80,0,'Y','Name of the primary or foreign key')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSCONSTRAINTS','TYPE','" & objDBE.dbparmOwner & "',4,'CHAR',1,1,0,'Y','Type of constraint: Primary or Foreign (P/F)')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSCONSTRAINTS','DELETERULE','" & objDBE.dbparmOwner & "',5,'CHAR',1,1,0,'Y','I have no idea... values seem to be R/C (Row/Column?)')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSFKCONSTRAINTS','TBNAME','" & objDBE.dbparmOwner & "',1,'VARCHAR',80,80,0,'Y','Table name of the referring column or field')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSFKCONSTRAINTS','TBOWNER','" & objDBE.dbparmOwner & "',2,'VARCHAR',80,80,0,'Y','Table owner of the referring column or field')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSFKCONSTRAINTS','CONSTRAINTNAME','" & objDBE.dbparmOwner & "',3,'VARCHAR',80,80,0,'Y','Name of the foreign key')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSFKCONSTRAINTS','COLSEQ','" & objDBE.dbparmOwner & "',4,'SMALLINT',2,0,0,'Y','Ordinal number of this column or field within the foreign key')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSFKCONSTRAINTS','REFSCOLUMN','" & objDBE.dbparmOwner & "',5,'VARCHAR',80,80,0,'Y','Referring column or field name')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSFKCONSTRAINTS','REFDTBNAME','" & objDBE.dbparmOwner & "',6,'VARCHAR',80,80,0,'Y','Referred table name')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSFKCONSTRAINTS','REFDTBOWNER','" & objDBE.dbparmOwner & "',6,'VARCHAR',80,80,0,'Y','Referred table owner')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSFKCONSTRAINTS','REFDCOLUMN','" & objDBE.dbparmOwner & "',8,'VARCHAR',80,80,0,'Y','Referred column or field name')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSPKCONSTRAINTS','TBNAME','" & objDBE.dbparmOwner & "',1,'VARCHAR',80,80,0,'Y','Table name of this primary key')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSPKCONSTRAINTS','TBOWNER','" & objDBE.dbparmOwner & "',2,'VARCHAR',80,80,0,'Y','Table owner of this primary key')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSPKCONSTRAINTS','COLSEQ','" & objDBE.dbparmOwner & "',3,'SMALLINT',2,0,0,'Y','Ordinal number of this column or field within the key')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values ('XSYSPKCONSTRAINTS','COLNAME','" & objDBE.dbparmOwner & "',4,'VARCHAR',80,80,0,'Y','Name of the column or field making up this segment of the primary key')") Then GoTo ExitSub
    
    Call objUI.Display(vbTab & "XSYSINDEXES", deBothStatusBarAndTextBox)
    If Not ExecuteSQL("DROP TABLE " & objDBE.dbparmOwner & "." & "XSYSINDEXES") Then GoTo ExitSub
    If Not ExecuteSQL( _
            "CREATE TABLE " & objDBE.dbparmOwner & "." & "XSYSINDEXES(" & _
                "NAME VARCHAR(80)," & _
                "TBNAME VARCHAR(80)," & _
                "TBOWNER VARCHAR(80)," & _
                "UNIQUERULE CHAR(1)," & _
                "COLCOUNT SMALLINT," & _
                "SYSTEM CHAR(1))") Then GoTo ExitSub
    If Not ExecuteSQL( _
            "CREATE UNIQUE INDEX XSYSIDXBYTABLE ON " & objDBE.dbparmOwner & "." & "XSYSINDEXES (" & _
                "TBNAME" & objDBE.dbparmAsc & ", " & _
                "TBOWNER" & objDBE.dbparmAsc & ", " & _
                "NAME" & objDBE.dbparmAsc & ")") Then GoTo ExitSub
    If Not ExecuteSQL( _
            "CREATE UNIQUE INDEX XSYSIDXBYNAME ON " & objDBE.dbparmOwner & "." & "XSYSINDEXES (" & _
                "NAME" & objDBE.dbparmAsc & ", " & _
                "TBNAME" & objDBE.dbparmAsc & ", " & _
                "TBOWNER" & objDBE.dbparmAsc & ")") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSINDEXES (TBNAME, NAME, TBOWNER, UNIQUERULE, COLCOUNT, SYSTEM) Values ('XSYSTABLES','XSYSTABLESBYNAME','" & objDBE.dbparmOwner & "','D',2,'Y')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSINDEXES (TBNAME, NAME, TBOWNER, UNIQUERULE, COLCOUNT, SYSTEM) Values ('XSYSCOLUMNS','XSYSCOLUMNSBYNAME','" & objDBE.dbparmOwner & "','D',2,'Y')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSINDEXES (TBNAME, NAME, TBOWNER, UNIQUERULE, COLCOUNT, SYSTEM) Values ('XSYSINDEXES','XSYSIDXBYTABLE','" & objDBE.dbparmOwner & "','D',2,'Y')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSINDEXES (TBNAME, NAME, TBOWNER, UNIQUERULE, COLCOUNT, SYSTEM) Values ('XSYSINDEXES','XSYSIDXBYNAME','" & objDBE.dbparmOwner & "','D',2,'Y')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSINDEXES (TBNAME, NAME, TBOWNER, UNIQUERULE, COLCOUNT, SYSTEM) Values ('XSYSKEYS','XSYSKEYSBYCOL','" & objDBE.dbparmOwner & "','D',2,'Y')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSINDEXES (TBNAME, NAME, TBOWNER, UNIQUERULE, COLCOUNT, SYSTEM) Values ('XSYSKEYS','XSYSKEYSBYSEQ','" & objDBE.dbparmOwner & "','D',2,'Y')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSINDEXES (TBNAME, NAME, TBOWNER, UNIQUERULE, COLCOUNT, SYSTEM) Values ('XSYSCONSTRAINTS','XSYSCONSTRAINTS','" & objDBE.dbparmOwner & "','D',2,'Y')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSINDEXES (TBNAME, NAME, TBOWNER, UNIQUERULE, COLCOUNT, SYSTEM) Values ('XSYSFKCONSTRAINTS','XSYSFKEYSBYNAME','" & objDBE.dbparmOwner & "','D',2,'Y')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSINDEXES (TBNAME, NAME, TBOWNER, UNIQUERULE, COLCOUNT, SYSTEM) Values ('XSYSFKCONSTRAINTS','XSYSFKEYSBYREFS','" & objDBE.dbparmOwner & "','D',2,'Y')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSINDEXES (TBNAME, NAME, TBOWNER, UNIQUERULE, COLCOUNT, SYSTEM) Values ('XSYSFKCONSTRAINTS','XSYSFKEYSBYREFD','" & objDBE.dbparmOwner & "','D',2,'Y')") Then GoTo ExitSub
    
    Call objUI.Display(vbTab & "XSYSKEYS", deBothStatusBarAndTextBox)
    If Not ExecuteSQL("DROP TABLE " & objDBE.dbparmOwner & "." & "XSYSKEYS") Then GoTo ExitSub
    If Not ExecuteSQL( _
            "CREATE TABLE " & objDBE.dbparmOwner & "." & "XSYSKEYS(" & _
                "IXNAME VARCHAR(80)," & _
                "COLNAME VARCHAR(80)," & _
                "COLNO SMALLINT," & _
                "COLSEQ SMALLINT," & _
                "ORDERING CHAR(1))") Then GoTo ExitSub
    If Not ExecuteSQL( _
            "CREATE UNIQUE INDEX XSYSKEYSBYCOL ON " & objDBE.dbparmOwner & "." & "XSYSKEYS (" & _
                "IXNAME" & objDBE.dbparmAsc & ", " & _
                "COLNAME" & objDBE.dbparmAsc & ")") Then GoTo ExitSub
    If Not ExecuteSQL( _
            "CREATE UNIQUE INDEX XSYSKEYSBYSEQ ON " & objDBE.dbparmOwner & "." & "XSYSKEYS (" & _
                "IXNAME" & objDBE.dbparmAsc & ", " & _
                "COLSEQ" & objDBE.dbparmAsc & ")") Then GoTo ExitSub
    
    If Not ExecuteSQL("Insert Into XSYSKEYS (IXNAME, COLNAME, COLNO, COLSEQ, ORDERING) Values ('XSYSTABLESBYNAME','NAME',1,1,'A')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSKEYS (IXNAME, COLNAME, COLNO, COLSEQ, ORDERING) Values ('XSYSTABLESBYNAME','OWNER',7,2,'A')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSKEYS (IXNAME, COLNAME, COLNO, COLSEQ, ORDERING) Values ('XSYSTABLESBYNAME','SNUM',5,3,'A')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSKEYS (IXNAME, COLNAME, COLNO, COLSEQ, ORDERING) Values ('XSYSCOLUMNSBYNAME','TBNAME',2,1,'A')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSKEYS (IXNAME, COLNAME, COLNO, COLSEQ, ORDERING) Values ('XSYSCOLUMNSBYNAME','TBOWNER',3,2,'A')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSKEYS (IXNAME, COLNAME, COLNO, COLSEQ, ORDERING) Values ('XSYSCOLUMNSBYNAME','NAME',1,3,'A')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSKEYS (IXNAME, COLNAME, COLNO, COLSEQ, ORDERING) Values ('XSYSIDXBYTABLE','TBNAME',2,1,'A')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSKEYS (IXNAME, COLNAME, COLNO, COLSEQ, ORDERING) Values ('XSYSIDXBYTABLE','TBOWNER',3,2,'A')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSKEYS (IXNAME, COLNAME, COLNO, COLSEQ, ORDERING) Values ('XSYSIDXBYTABLE','NAME',1,3,'A')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSKEYS (IXNAME, COLNAME, COLNO, COLSEQ, ORDERING) Values ('XSYSIDXBYNAME','NAME',1,1,'A')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSKEYS (IXNAME, COLNAME, COLNO, COLSEQ, ORDERING) Values ('XSYSIDXBYNAME','TBNAME',2,2,'A')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSKEYS (IXNAME, COLNAME, COLNO, COLSEQ, ORDERING) Values ('XSYSIDXBYNAME','TBOWNER',3,3,'A')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSKEYS (IXNAME, COLNAME, COLNO, COLSEQ, ORDERING) Values ('XSYSKEYSBYCOL','IXNAME',1,1,'A')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSKEYS (IXNAME, COLNAME, COLNO, COLSEQ, ORDERING) Values ('XSYSKEYSBYCOL','COLNAME',2,2,'A')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSKEYS (IXNAME, COLNAME, COLNO, COLSEQ, ORDERING) Values ('XSYSKEYSBYSEQ','IXNAME',1,1,'A')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSKEYS (IXNAME, COLNAME, COLNO, COLSEQ, ORDERING) Values ('XSYSKEYSBYSEQ','COLSEQ',4,2,'A')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSKEYS (IXNAME, COLNAME, COLNO, COLSEQ, ORDERING) Values ('XSYSCONSTRAINTS','TBNAME',1,1,'A')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSKEYS (IXNAME, COLNAME, COLNO, COLSEQ, ORDERING) Values ('XSYSCONSTRAINTS','TBOWNER',2,2,'A')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSKEYS (IXNAME, COLNAME, COLNO, COLSEQ, ORDERING) Values ('XSYSCONSTRAINTS','CONSTRAINTNAME',3,3,'A')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSKEYS (IXNAME, COLNAME, COLNO, COLSEQ, ORDERING) Values ('XSYSFKEYSBYNAME','TBNAME',1,1,'A')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSKEYS (IXNAME, COLNAME, COLNO, COLSEQ, ORDERING) Values ('XSYSFKEYSBYNAME','TBOWNER',2,2,'A')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSKEYS (IXNAME, COLNAME, COLNO, COLSEQ, ORDERING) Values ('XSYSFKEYSBYNAME','CONSTRAINTNAME',3,3,'A')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSKEYS (IXNAME, COLNAME, COLNO, COLSEQ, ORDERING) Values ('XSYSFKEYSBYREFS','TBNAME',1,1,'A')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSKEYS (IXNAME, COLNAME, COLNO, COLSEQ, ORDERING) Values ('XSYSFKEYSBYREFS','TBOWNER',2,2,'A')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSKEYS (IXNAME, COLNAME, COLNO, COLSEQ, ORDERING) Values ('XSYSFKEYSBYREFS','REFSCOLUMN',5,3,'A')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSKEYS (IXNAME, COLNAME, COLNO, COLSEQ, ORDERING) Values ('XSYSFKEYSBYREFD','TBNAME',1,1,'A')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSKEYS (IXNAME, COLNAME, COLNO, COLSEQ, ORDERING) Values ('XSYSFKEYSBYREFD','REFSCOLUMN',5,2,'A')") Then GoTo ExitSub
    If Not ExecuteSQL("Insert Into XSYSKEYS (IXNAME, COLNAME, COLNO, COLSEQ, ORDERING) Values ('XSYSFKEYSBYREFD','REFDCOLUMN',8,3,'A')") Then GoTo ExitSub

    Call objUI.Display(vbTab & "XSYSCONSTRAINTS", deBothStatusBarAndTextBox)
    If Not ExecuteSQL("DROP TABLE " & objDBE.dbparmOwner & "." & "XSYSCONSTRAINTS") Then GoTo ExitSub
    If Not ExecuteSQL( _
            "CREATE TABLE " & objDBE.dbparmOwner & "." & "XSYSCONSTRAINTS(" & _
                "TBNAME VARCHAR(80)," & _
                "TBOWNER VARCHAR(80)," & _
                "CONSTRAINTNAME VARCHAR(80)," & _
                "TYPE CHAR(1)," & _
                "DELETERULE CHAR(1))") Then GoTo ExitSub
    If Not ExecuteSQL( _
            "CREATE UNIQUE INDEX XSYSCONSTRAINTS ON " & objDBE.dbparmOwner & "." & "XSYSCONSTRAINTS (" & _
                "TBNAME" & objDBE.dbparmAsc & ", " & _
                "TBOWNER" & objDBE.dbparmAsc & ", " & _
                "CONSTRAINTNAME" & objDBE.dbparmAsc & ")") Then GoTo ExitSub

    Call objUI.Display(vbTab & "XSYSFKCONSTRAINTS", deBothStatusBarAndTextBox)
    If Not ExecuteSQL("DROP TABLE " & objDBE.dbparmOwner & "." & "XSYSFKCONSTRAINTS") Then GoTo ExitSub
    If Not ExecuteSQL( _
            "CREATE TABLE " & objDBE.dbparmOwner & "." & "XSYSFKCONSTRAINTS(" & _
                "TBNAME VARCHAR(80)," & _
                "TBOWNER VARCHAR(80)," & _
                "CONSTRAINTNAME VARCHAR(80)," & _
                "COLSEQ SMALLINT," & _
                "REFSCOLUMN VARCHAR(80)," & _
                "REFDTBNAME VARCHAR(80)," & _
                "REFDTBOWNER VARCHAR(80)," & _
                "REFDCOLUMN VARCHAR(80))") Then GoTo ExitSub
    If Not ExecuteSQL( _
            "CREATE UNIQUE INDEX XSYSFKEYSBYNAME ON " & objDBE.dbparmOwner & "." & "XSYSFKCONSTRAINTS (" & _
                "TBNAME" & objDBE.dbparmAsc & ", " & _
                "TBOWNER" & objDBE.dbparmAsc & ", " & _
                "CONSTRAINTNAME" & objDBE.dbparmAsc & ")") Then GoTo ExitSub
    If Not ExecuteSQL( _
            "CREATE UNIQUE INDEX XSYSFKEYSBYREFS ON " & objDBE.dbparmOwner & "." & "XSYSFKCONSTRAINTS (" & _
                "TBNAME" & objDBE.dbparmAsc & ", " & _
                "TBOWNER" & objDBE.dbparmAsc & ", " & _
                "REFSCOLUMN" & objDBE.dbparmAsc & ")") Then GoTo ExitSub
    'Note: This index hasn't the parts you'd think it should because its overall length
    '      would exceed SQL Server's limit...
    If Not ExecuteSQL( _
            "CREATE INDEX XSYSFKEYSBYREFD ON " & objDBE.dbparmOwner & "." & "XSYSFKCONSTRAINTS (" & _
                "TBNAME" & objDBE.dbparmAsc & ", " & _
                "REFSCOLUMN" & objDBE.dbparmAsc & ", " & _
                "REFDCOLUMN" & objDBE.dbparmAsc & ")") Then GoTo ExitSub

    Call objUI.Display(vbTab & "XSYSPKCONSTRAINTS", deBothStatusBarAndTextBox)
    If Not ExecuteSQL("DROP TABLE " & objDBE.dbparmOwner & "." & "XSYSPKCONSTRAINTS") Then GoTo ExitSub
    If Not ExecuteSQL( _
            "CREATE TABLE " & objDBE.dbparmOwner & "." & "XSYSPKCONSTRAINTS(" & _
                "TBNAME VARCHAR(80)," & _
                "TBOWNER VARCHAR(80)," & _
                "COLSEQ SMALLINT," & _
                "COLNAME VARCHAR(80))") Then GoTo ExitSub
    
    CreateFilesBookTables = True

ExitSub:
    Call Trace(trcExit, "CreateFilesBookTables()", trcSQL)
    Exit Function
End Function
Public Function doDataType(DBEngine As SIASDB.dbeDBEngineEnum, ColumnName As String, DataType As String, strDefault As String, strStorage As String, strPrecision As String, strScale As String) As Boolean
    Dim iPos As Integer
    Dim baseDataType As String
    
    Call Trace(trcEnter, "doDataType(""" & DBEngine & """,""" & ColumnName & """,""" & DataType & """,""" & strDefault & """,""" & strStorage & """,""" & strPrecision & """,""" & strScale & ""","")", trcSQL)
    Call objDBE.initDBEngineParms(DBEngine)
    
    doDataType = False
    strDefault = vbNullString
    strStorage = "0"
    strPrecision = "0"
    strScale = "0"
    DataType = Trim(DataType)
    baseDataType = ParseStr(DataType, 1, "(")

    Select Case UCase(Trim(baseDataType))
        '============================================================================================
        'Integer data types...
        Case "BIT"
            'SQL Server Only...
            'This is tricky, because SQL Server will pack bits together so if the bit string grows
            'beyond a single byte, this one-byte assumption will be inaccurate...
            '(OK, not so much tricky as lazy)...
            strDefault = "0"
            strStorage = "1"
        Case "INT", "INTEGER"
            strDefault = "0"
            strStorage = "4"
        Case "SMALLINT"
            strDefault = "0"
            strStorage = "2"
        Case "TINYINT"
            strDefault = "0"
            strStorage = "1"                         'this is really one byte... no growing allowed...
        '============================================================================================
        'Decimal/numeric data types...
        Case "DECIMAL", "DEC", "NUMERIC"
            Select Case DBEngine
                Case dbeAccess
                    DataType = "Access DECIMAL data type not [yet] supported"
                    GoTo ExitSub
                Case dbeSQLBase
                    DataType = "SQLBase DECIMAL data type not [yet] supported"
                    GoTo ExitSub
                Case dbeSQLServer
            End Select
            strDefault = "0"
            strStorage = "5"                            'Assume the smallest...
            iPos = InStr(DataType, "(")
            If iPos > 0 Then
                strPrecision = ParseStr(DataType, 2, "(")
                strPrecision = Left(strPrecision, Len(strPrecision) - 1)
                strScale = ParseStr(strPrecision, 2, ",")
                strPrecision = ParseStr(strPrecision, 1, ",")
                If Val(strPrecision) >= 1 And Val(strPrecision) <= 9 Then
                    strStorage = "5"
                ElseIf Val(strPrecision) >= 10 And Val(strPrecision) <= 19 Then
                    strStorage = "9"
                ElseIf Val(strPrecision) >= 20 And Val(strPrecision) <= 28 Then
                    strStorage = "13"
                ElseIf Val(strPrecision) >= 29 And Val(strPrecision) <= 38 Then
                    strStorage = "17"
                End If
                DataType = ParseStr(DataType, 1, "(")
            End If
        '============================================================================================
        'Currency data types...
        Case "MONEY"
            'SQL Server data type...
            Select Case DBEngine
                Case dbeAccess
                    DataType = "CURRENCY"
                    strDefault = "0.0"
                    strStorage = "8"
                Case dbeSQLBase
                    DataType = "FLOAT"
                    strDefault = "0.0"
                    strStorage = "8"
                Case dbeSQLServer
                    strDefault = "0.0"
                    strStorage = "8"
            End Select
        Case "SMALLMONEY"
            'SQL Server data type...
            Select Case DBEngine
                Case dbeAccess
                    DataType = "???"                'Do we care?
                    strDefault = "0.0"
                    strStorage = "4"
                Case dbeSQLBase
                    DataType = "REAL"
                    strDefault = "0.0"
                    strStorage = "4"
                Case dbeSQLServer
                    strDefault = "0.0"
                    strStorage = "4"
            End Select
        '============================================================================================
        'Floating-Point data types...
        Case "FLOAT", "DOUBLE PRECISION"
            iPos = InStr(DataType, "(")
            If iPos > 0 Then
                strPrecision = ParseStr(DataType, 2, "(")
                strPrecision = Left(strPrecision, Len(strPrecision) - 1)
                DataType = ParseStr(DataType, 1, "(")
            End If
            Select Case DBEngine
                Case dbeAccess
                    DataType = "CURRENCY"
                    strDefault = "0.0"
                    strStorage = "8"
                Case dbeSQLBase
                    strDefault = "0.0"
                    strStorage = "8"
                Case dbeSQLServer
                    'FLOAT(n) [for n=1-7]=REAL
                    strDefault = "0.0"
                    If Val(strPrecision) >= 1 And Val(strPrecision) <= 7 Then
                        strStorage = "4"
                        DataType = "REAL"
                    Else
                        strStorage = "8"
                    End If
            End Select
        Case "REAL"
            Select Case DBEngine
                Case dbeAccess
                    DataType = "???"                'Do we care?
                    strDefault = "0.0"
                    strStorage = "4"
                Case dbeSQLBase
                    strDefault = "0.0"
                    strStorage = "4"
                Case dbeSQLServer
                    strDefault = "0.0"
                    strStorage = "4"
            End Select
        '============================================================================================
        'Date/Time data types...
        Case "DATETIME"
            'SQL Server data type...
            Select Case DBEngine
                Case dbeAccess
                    DataType = "DATE"
                    strStorage = "4"
                Case dbeSQLBase
                    DataType = "TIMESTAMP"
                    strStorage = "10"
                Case dbeSQLServer
                    DataType = "DATETIME"
                    strStorage = "8"
            End Select
            strDefault = objDBE.dbparmTimeStamp
        Case "SMALLDATETIME"
            'SQL Server data type...
            Select Case DBEngine
                Case dbeAccess
                    DataType = "DATE"
                    strStorage = "4"
                Case dbeSQLBase
                    DataType = "DATE"
                    strStorage = "4"
                Case dbeSQLServer
                    strStorage = "4"
            End Select
            strDefault = objDBE.dbparmTimeStamp
        Case "DATE"
            'SQLBase data type...
            Select Case DBEngine
                Case dbeAccess
                    strStorage = "4"
                Case dbeSQLBase
                    strStorage = "4"
                Case dbeSQLServer
                    'Even though this is stored as a 4-byte date data type in other
                    'engines, the SQL Server SMALLDATETIME datatype only supports a
                    'date range of 01/01/1900 through 06/06/2079... Since we need to
                    'go from 01/01/1900 through 12/31/2999, we need DATETIME which
                    'costs 4 more bytes, but gives us a range of 01/01/1753 through
                    '12/31/9999 which is plenty... (not to mention the time component)...
                    'DataType = "SMALLDATETIME"
                    'strStorage = "4"
                    DataType = "DATETIME"
                    strStorage = "8"
            End Select
            strDefault = objDBE.dbparmTimeStamp
        Case "TIMESTMP", "TIMESTAMP"
            Select Case DBEngine
                Case dbeAccess
                    DataType = "DATE"
                    strStorage = "4"
                    strDefault = objDBE.dbparmTimeStamp
                Case dbeSQLBase
                    DataType = "TIMESTAMP"
                    strStorage = "10"
                    strDefault = objDBE.dbparmTimeStamp
                Case dbeSQLServer
                    If ColumnName <> "ROWID" Then
                        'SQL Server uses TIMESTAMP fields as ROWID's, but this is taken care of
                        'by the calling routine (only ROWID field can be a TIMESTAMP)... So if we
                        'find one here, it must not be for the ROWID purpose, so use DATETIME...
                        DataType = "DATETIME"
                        strStorage = "8"
                    Else
                        strStorage = "8"
                    End If
            End Select
        '============================================================================================
        'Other Numeric data types...
        Case "CURSOR"
            'SQL Server data type...
            DataType = "SQL Server CURSOR data type not supported"
            GoTo ExitSub
        Case "UNIQUEIDENTIFIER"
            'SQL Server data type...
            DataType = "SQL Server UNIQUEIDENTIFIER data type not supported"
            GoTo ExitSub
        '============================================================================================
        'Character String data types...
        Case "CHAR", "CHARACTER"
            strDefault = "' '"
            iPos = InStr(DataType, "(")
            If iPos > 0 Then
                strPrecision = ParseStr(DataType, 2, "(")
                strPrecision = Left(strPrecision, Len(strPrecision) - 1)
                DataType = ParseStr(DataType, 1, "(")
            Else
                strPrecision = "1"
            End If
            strStorage = strPrecision
        Case "VARCHAR", "CHAR VARYING", "CHARACTER VARYING"
            strDefault = "' '"
            iPos = InStr(DataType, "(")
            If iPos > 0 Then
                strPrecision = ParseStr(DataType, 2, "(")
                strPrecision = Left(strPrecision, Len(strPrecision) - 1)
                DataType = ParseStr(DataType, 1, "(")
            Else
                strPrecision = "1"
            End If
            strStorage = strPrecision
        Case "TEXT", "MEMO", "LONG"
            Select Case DBEngine
                Case dbeAccess
                    DataType = "MEMO"
                Case dbeSQLBase
                    DataType = "LONG VARCHAR"
                Case dbeSQLServer
                    DataType = "TEXT"
            End Select
            strStorage = "-1"
        '============================================================================================
        'Unicode Character String data types...
        Case "NCHAR", "NATIONAL CHARACTER", "NATIONAL CHAR"
            strDefault = "' '"
            iPos = InStr(DataType, "(")
            If iPos > 0 Then
                strPrecision = ParseStr(DataType, 2, "(")
                strPrecision = Left(strPrecision, Len(strPrecision) - 1)
                DataType = ParseStr(DataType, 1, "(")
            Else
                strPrecision = "1"
            End If
            strStorage = 2 * Val(strPrecision)
        Case "NVARCHAR", "NATIONAL CHARACTER VARYING", "NATIONAL CHAR VARYING"
            strDefault = "' '"
            iPos = InStr(DataType, "(")
            If iPos > 0 Then
                strPrecision = ParseStr(DataType, 2, "(")
                strPrecision = Left(strPrecision, Len(strPrecision) - 1)
                DataType = ParseStr(DataType, 1, "(")
            Else
                strPrecision = "1"
            End If
            strStorage = 2 * Val(strPrecision)
        Case "NTEXT", "NATIONAL TEXT"
            'SQL Server data type...
            DataType = "SQL Server NTEXT data type not supported"
            GoTo ExitSub
        '============================================================================================
        'Binary String data types...
        Case "BINARY"
            'SQL Server data type...
            DataType = "SQL Server BINARY data type not supported"
            GoTo ExitSub
        Case "VARBINARY", "BINARY VARYING"
            'SQL Server data type...
            DataType = "SQL Server VARBINARY data type not supported"
            GoTo ExitSub
        Case "IMAGE"
            Select Case DBEngine
                Case dbeSQLBase
                    DataType = "LONG VARCHAR"
            End Select
            strStorage = "-1"
        Case "LONGVAR", "LONG VARCHAR"
            Select Case DBEngine
                Case dbeAccess
                    DataType = " IMAGE"  '?
                Case dbeSQLBase
                    DataType = "LONG VARCHAR"
                Case dbeSQLServer
                    DataType = "IMAGE"
            End Select
            strStorage = "-1"
        Case Else
            'Error...
            DataType = "Unrecognized data type: """ & baseDataType & """"
            GoTo ExitSub
    End Select
    doDataType = True

ExitSub:
    Call Trace(trcExit, "doDataType(""" & DBEngine & """,""" & ColumnName & """,""" & DataType & """,""" & strDefault & """,""" & strStorage & """,""" & strPrecision & """,""" & strScale & ""","")", trcSQL)
    Exit Function
End Function
'Public Sub Disconnect(fClearFileDSN As Boolean)
'    On Error Resume Next
'    Call objDBE.Disconnect(True)
'    If fClearFileDSN Then gFiRRe.frmMain.cboFileDSN.ListIndex = -1
'    gFiRRe.frmMain.sbStatus.Panels("DBEngine").Visible = False
'    gFiRRe.frmMain.sbStatus.Panels("DBEngine").Text = vbNullString
'    gFiRRe.frmMain.sbStatus.Panels("Server").Visible = False
'    gFiRRe.frmMain.sbStatus.Panels("Server").Text = vbNullString
'    gFiRRe.frmMain.sbStatus.Panels("Database").Visible = False
'    gFiRRe.frmMain.sbStatus.Panels("Database").Text = vbNullString
'End Sub
Private Function doAlterTable(SQLsource As String) As Boolean
    Dim Operation As String
    Dim fSuccess As Boolean
    Dim uSQLsource As String
    
    'The purpose of this routine is two-fold...
    'First we need to translate a SQLBase syntax command into something SQL Server will understand
    '(if necessary)... (don't know the long term impact of using SQLBase as the base syntax, but
    'we're all familiar with it, so we'll use it for now)...
    'Second, we need to disect the SQL statement to know how to maintain the XSYS system tables used
    'to generate our FilesBook report.
    
    Call Trace(trcEnter, "doAlterTable: " & SQLsource, trcSQL)
    
    fSuccess = False
    uSQLsource = Trim(UCase(SQLsource))
    'Need to take different ALTER TABLE syntaxes into consideration here... (or remove it from the
    'calling routine (like modPopulateDB.DoConstraints) and put that logic here where it belongs...
    Operation = ParseStr(uSQLsource, 4, " ")
    Select Case Operation
        Case "ADD"
            fSuccess = doAlterTableAdd(SQLsource)
        Case "DROP"
            fSuccess = doAlterTableDrop(SQLsource)
        Case "MODIFY"
            fSuccess = doAlterTableModify(SQLsource)
        Case "RENAME"
            fSuccess = doAlterTableRename(SQLsource)
        Case "PRIMARY"
            fSuccess = doAlterTableAddPrimaryKey(SQLsource)
        Case "FOREIGN"
            fSuccess = doAlterTableAddForeignKey(SQLsource)
        Case Else
            'Error...
            SQLsource = "Unrecognized ALTER TABLE command:" & vbCrLf & SQLsource
            GoTo ExitSub
    End Select
    
ExitSub:
    doAlterTable = fSuccess
    DoEvents
    Call Trace(trcExit, "doAlterTable", trcSQL)
    Exit Function
End Function
Private Function doAlterTableAdd(SQLsource As String) As Boolean
    Dim adoRS As ADODB.Recordset
    Dim iPos As Integer
    Dim TableName As String
    Dim DataType As String
    Dim ColumnName As String
    Dim ColumnNumber As Integer
    Dim fSuccess As Boolean
    Dim uSQLsource As String
    Dim NewSQLSource As String
    Dim strDefault As String
    Dim strPrecision As String
    Dim strScale As String
    Dim strStorage As String
    Dim strNulls As String
    Dim strRemarks As String
    Dim strOwner As String
    
    'SQLBase:    ALTER TABLE <TableName> ADD <ColumnName> <DataType> [[] | [NOT NULL] | [NOT NULL WITH DEFAULT]]
    'SQL Server: ALTER TABLE <TableName> ADD <ColumnName> <DataType [NULL | NOT NULL]
    
    Call Trace(trcEnter, "doAlterTableAdd: " & SQLsource, trcSQL)
    
    fSuccess = False
    uSQLsource = Trim(UCase(SQLsource))
    TableName = ParseStr(uSQLsource, 3, " ")
    strOwner = objDBE.dbparmOwner
    If InStr(TableName, ".") Then
        strOwner = ParseStr(TableName, 1, ".")
        If strOwner <> objDBE.dbparmOwner And strOwner = UCase(objDBE.dbparmOwner) Then strOwner = objDBE.dbparmOwner
        TableName = ParseStr(TableName, 2, ".")
    End If
    
    ColumnName = ParseStr(uSQLsource, 5, " ")
    DataType = ParseStr(uSQLsource, 6, " ")
    
    'Get the next sequential column number for this table...
    Set adoRS = New ADODB.Recordset
    adoRS.Open "Select max(COLNO)+1 From XSYSCOLUMNS Where TBNAME='" & TableName & "' And TBOWNER='" & strOwner & "'", objDBE.ActiveConnection, adOpenKeyset, adLockReadOnly
    ColumnNumber = adoRS(0)
    Call CloseRecordset(adoRS, True)
    
    If Not doDataType(objDBE.DBEngine, ColumnName, DataType, strDefault, strStorage, strPrecision, strScale) Then
        'Note: Errors are returned in DataType field...
        SQLsource = "Error: '" & DataType & "' encountered in ALTER TABLE command:" & vbCrLf & SQLsource
        GoTo ExitSub
    End If
    
    If Right(uSQLsource, Len("NOT NULL")) = "NOT NULL" Then
        strNulls = "N"
    ElseIf InStr(uSQLsource, " NOT NULL") > 0 And InStr(uSQLsource, " DEFAULT") > 0 Then
        strNulls = "D"
    Else
        strNulls = "Y"
    End If
    
    'Update meta-data...
    strRemarks = " "
    If ColumnName = "ROWID" Then strRemarks = "Unique record identifier used to handle concurrency."
    If Not doSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values " & "('" & TableName & "','" & ColumnName & "','" & strOwner & "'," & ColumnNumber & ",'" & DataType & "'," & strStorage & "," & strPrecision & "," & strScale & ",'" & strNulls & "','" & FixQuotes(strRemarks) & "')") Then GoTo ExitSub
    
    Select Case objDBE.DBEngine
        Case dbeSQLBase
            NewSQLSource = SQLsource
        Case dbeSQLServer
            Select Case strNulls
                Case "N"
                    strNulls = "NOT NULL"
                Case "D"
                    'This should never happen, but might with unusual data types...
                    If strDefault <> vbNullString Then
                        strNulls = "NOT NULL " & objDBE.dbparmDefault & strDefault
                    Else
                        strNulls = "NOT NULL"
                    End If
                Case Else
                    strNulls = "NULL"
            End Select
            DataType = ReassembleDataType(DataType, strPrecision, strScale)
            NewSQLSource = "ALTER TABLE " & TableName & " ADD " & ColumnName & " " & DataType & " " & strNulls
        Case Else
    End Select
    
    'Execute the change to the database...
    fSuccess = ExecuteSQL(NewSQLSource)
    SQLsource = NewSQLSource
            
ExitSub:
    doAlterTableAdd = fSuccess
    DoEvents
    Call Trace(trcExit, "doAlterTableAdd", trcSQL)
    Exit Function
End Function
Private Function doAlterTableAddForeignKey(SQLsource As String) As Boolean
    Dim adoRS As New ADODB.Recordset
    Dim i As Integer
    Dim TableName As String
    Dim ParentTable As String
    Dim ParentColumnList As String
    Dim ColumnName As String
    Dim ColumnList As String
    Dim ConstraintName As String
    Dim DeleteRule As String
    Dim fSuccess As Boolean
    Dim uSQLsource As String
    Dim NewSQLSource As String
    Dim strOwner As String

    'SQLBase:    ALTER TABLE <TableName> FOREIGN KEY <FKey>(<ColumnList>) REFERENCES <ParentTable> [ON DELETE [RESTRICT | CASCADE | SET NULL]]
    'SQL Server: ALTER TABLE <TableName> ADD CONSTRAINT <FKey> FOREIGN KEY (<ColumnList>) REFERENCES <ParentTable> (<ParentColumnList>)
    'MS Access:  ALTER TABLE <TableName> ADD CONSTRAINT <FKey> FOREIGN KEY (<ColumnList>) REFERENCES <ParentTable> (<ParentColumnList>)
    
    Call Trace(trcEnter, "doAlterTableAddForeignKey: " & SQLsource, trcSQL)
    
    fSuccess = False
    uSQLsource = Trim(UCase(SQLsource))
    TableName = ParseStr(uSQLsource, 3, " ")
    strOwner = objDBE.dbparmOwner
    If InStr(TableName, ".") Then
        strOwner = ParseStr(TableName, 1, ".")
        If strOwner <> objDBE.dbparmOwner And strOwner = UCase(objDBE.dbparmOwner) Then strOwner = objDBE.dbparmOwner
        TableName = ParseStr(TableName, 2, ".")
    End If
    ConstraintName = ParseStr(uSQLsource, 6, " ")
    i = InStr(ConstraintName, "(")
    If i > 0 Then ConstraintName = Mid(ConstraintName, 1, i - 1)
    
    ColumnList = ParseStr(uSQLsource, 2, "(")
    ColumnList = Mid(ColumnList, 1, InStr(ColumnList, ")") - 1)
    ParentTable = Mid(uSQLsource, InStr(uSQLsource, " REFERENCES ") + Len(" REFERENCES "))
    ParentTable = ParseStr(ParentTable, 1, " ")
    If InStr(ParentTable, ".") Then ParentTable = ParseStr(ParentTable, 2, ".")
    If Right(uSQLsource, Len("ON DELETE RESTRICT")) = "ON DELETE RESTRICT" Then
        DeleteRule = "R"
    ElseIf Right(uSQLsource, Len("ON DELETE CASCADE")) = "ON DELETE CASCADE" Then
        DeleteRule = "C"
    ElseIf Right(uSQLsource, Len("ON DELETE SET NULL")) = "ON DELETE SET NULL" Then
        DeleteRule = "N"
    Else
        DeleteRule = " "
    End If
    
    If Not doSQL("Insert Into XSYSCONSTRAINTS (TBNAME, TBOWNER, CONSTRAINTNAME, TYPE, DELETERULE) Values ('" & TableName & "','" & strOwner & "','" & ConstraintName & "','F','" & DeleteRule & "')") Then GoTo ExitSub
    
    For i = 1 To TokenCount(ColumnList, ",")
        'Find the Parent Table's PRIMARY KEY field...
        Call objDBE.OpenRecordset(adoRS, "Select COLNAME From XSYSPKCONSTRAINTS Where TBNAME='" & ParentTable & "' And TBOWNER='" & strOwner & "' And COLSEQ=" & i, adOpenKeyset, adLockReadOnly)
        ColumnName = ParseStr(ColumnList, i, ",", """")
        If Not doSQL("Insert Into XSYSFKCONSTRAINTS (TBNAME, TBOWNER, CONSTRAINTNAME, COLSEQ, REFSCOLUMN, REFDTBNAME, REFDTBOWNER, REFDCOLUMN) Values ('" & _
            TableName & "','" & strOwner & "','" & ConstraintName & "'," & i & ",'" & ColumnName & "','" & ParentTable & "','" & strOwner & "','" & adoRS("COLNAME") & "')") Then GoTo ExitSub
        ParentColumnList = ParentColumnList & adoRS("COLNAME")
        If i < TokenCount(ColumnList, ",") Then ParentColumnList = ParentColumnList & ","
        Call CloseRecordset(adoRS, False)
    Next i
            
    Select Case objDBE.DBEngine
        Case dbeAccess
            NewSQLSource = "ALTER TABLE " & strOwner & "." & TableName & " ADD CONSTRAINT " & ConstraintName & " FOREIGN KEY (" & ColumnList & ") REFERENCES " & strOwner & "." & ParentTable & "(" & ParentColumnList & ")"
        Case dbeSQLBase
            NewSQLSource = SQLsource
        Case dbeSQLServer
            NewSQLSource = "ALTER TABLE " & strOwner & "." & TableName & " ADD CONSTRAINT " & ConstraintName & " FOREIGN KEY (" & ColumnList & ") REFERENCES " & strOwner & "." & ParentTable & "(" & ParentColumnList & ")"
        Case Else
    End Select
    
    'Execute the change to the database...
    fSuccess = ExecuteSQL(NewSQLSource)
    SQLsource = NewSQLSource
            
ExitSub:
    Call CloseRecordset(adoRS, True)
    doAlterTableAddForeignKey = fSuccess
    DoEvents
    Call Trace(trcExit, "doAlterTableAddForeignKey", trcSQL)
    Exit Function
End Function
Private Function doAlterTableAddPrimaryKey(SQLsource As String) As Boolean
    Dim adoRS As ADODB.Recordset
    Dim i As Integer
    Dim TableName As String
    Dim ColumnList As String
    Dim ColumnName As String
    Dim ConstraintName As String
    Dim fSuccess As Boolean
    Dim uSQLsource As String
    Dim NewSQLSource As String
    Dim strOwner As String

    'SQLBase:    ALTER TABLE <TableName> PRIMARY KEY(<ColumnList>)
    'SQL Server: ALTER TABLE <TableName> ADD CONSTRAINT <TableName>_PK PRIMARY KEY [CLUSTERED] (<ColumnList>)
    'MS Access:  ALTER TABLE <TableName> ADD CONSTRAINT <TableName>_PK PRIMARY KEY [CLUSTERED] (<ColumnList>)
    
    Call Trace(trcEnter, "doAlterTableAddPrimaryKey: " & SQLsource, trcSQL)
    
    fSuccess = False
    uSQLsource = Trim(UCase(SQLsource))
    TableName = ParseStr(uSQLsource, 3, " ")
    strOwner = objDBE.dbparmOwner
    If InStr(TableName, ".") Then
        strOwner = ParseStr(TableName, 1, ".")
        If strOwner <> objDBE.dbparmOwner And strOwner = UCase(objDBE.dbparmOwner) Then strOwner = objDBE.dbparmOwner
        TableName = ParseStr(TableName, 2, ".")
    End If
    Select Case objDBE.DBEngine
        Case dbeSQLBase
            ConstraintName = "PRIMARY"
        Case dbeSQLServer, dbeAccess
            'ConstraintName = TableName & "_PK"
            ConstraintName = ParseStr(uSQLsource, 6, " ")
            i = InStr(ConstraintName, "(")
            If i > 0 Then ConstraintName = Mid(ConstraintName, 1, i - 1)
        Case Else
    End Select
    ColumnList = ParseStr(uSQLsource, 2, "(")
    ColumnList = Mid(ColumnList, 1, InStr(ColumnList, ")") - 1)
            
    If Not doSQL("Insert Into XSYSCONSTRAINTS (TBNAME, TBOWNER, CONSTRAINTNAME, TYPE, DELETERULE) Values ('" & TableName & "','" & strOwner & "','" & ConstraintName & "','P',' ')") Then GoTo ExitSub
    
    For i = 1 To TokenCount(ColumnList, ",")
        ColumnName = ParseStr(ColumnList, i, ",", """")
        If Not doSQL("Insert Into XSYSPKCONSTRAINTS (TBNAME, TBOWNER, COLNAME, COLSEQ) Values ('" & TableName & "','" & strOwner & "','" & ColumnName & "'," & i & ")") Then GoTo ExitSub
    Next i
    
    Select Case objDBE.DBEngine
        Case dbeAccess
            NewSQLSource = "ALTER TABLE " & strOwner & "." & TableName & " ADD CONSTRAINT " & ConstraintName & " PRIMARY KEY (" & ColumnList & ")"
        Case dbeSQLBase
            'NewSQLSource = SQLsource
            NewSQLSource = "ALTER TABLE " & strOwner & "." & TableName & " PRIMARY KEY (" & ColumnList & ")"
        Case dbeSQLServer
            'NewSQLSource = "ALTER TABLE " & strOwner & "." & TableName & " ADD CONSTRAINT " & ConstraintName & " PRIMARY KEY (" & ColumnList & ")"
            NewSQLSource = "ALTER TABLE " & strOwner & "." & TableName & " ADD CONSTRAINT " & ConstraintName & " PRIMARY KEY CLUSTERED (" & ColumnList & ")"
        Case Else
    End Select
    
    'Execute the change to the database...
    fSuccess = ExecuteSQL(NewSQLSource)
    SQLsource = NewSQLSource
            
ExitSub:
    doAlterTableAddPrimaryKey = fSuccess
    DoEvents
    Call Trace(trcExit, "doAlterTableAddPrimaryKey", trcSQL)
    Exit Function
End Function
Private Function doAlterTableDrop(SQLsource As String) As Boolean
    Dim adoRS As ADODB.Recordset
    Dim iPos As Integer
    Dim TableName As String
    Dim ColumnName As String
    Dim ConstraintName As String
    Dim fSuccess As Boolean
    Dim uSQLsource As String
    Dim NewSQLSource As String
    Dim strOwner As String

    'SQLBase:    ALTER TABLE <TableName> DROP PRIMARY KEY <PKey>
    '            ALTER TABLE <TableName> DROP FOREIGN KEY <FKey>
    '            ALTER TABLE <TableName> DROP <ColumnName>
    'SQL Server: ALTER TABLE <TableName> DROP CONSTRAINT <TableName>_PK
    '            ALTER TABLE <TableName> DROP CONSTRAINT <ConstraintName>
    '            ALTER TABLE <TableName> DROP COLUMN <ColumnName>
    
    Call Trace(trcEnter, "doAlterTableDrop: " & SQLsource, trcSQL)
    
    fSuccess = False
    uSQLsource = Trim(UCase(SQLsource))
    TableName = ParseStr(uSQLsource, 3, " ")
    strOwner = objDBE.dbparmOwner
    If InStr(TableName, ".") Then
        strOwner = ParseStr(TableName, 1, ".")
        If strOwner <> objDBE.dbparmOwner And strOwner = UCase(objDBE.dbparmOwner) Then strOwner = objDBE.dbparmOwner
        TableName = ParseStr(TableName, 2, ".")
    End If
            
    iPos = InStr(uSQLsource, " DROP ") + Len(" DROP ")
    
    If Mid(uSQLsource, iPos, Len("PRIMARY KEY")) = "PRIMARY KEY" Then
        ConstraintName = "PRIMARY"
        If Not doSQL("Delete From XSYSCONSTRAINTS Where TBNAME='" & TableName & "' And TBOWNER='" & strOwner & "' And TYPE='P'") Then GoTo ExitSub
        If Not doSQL("Delete From XSYSPKCONSTRAINTS Where TBNAME='" & TableName & "' And TBOWNER='" & strOwner & "'") Then GoTo ExitSub
        Select Case objDBE.DBEngine
            Case dbeAccess
                    NewSQLSource = "ALTER TABLE " & strOwner & "." & TableName & " DROP PRIMARY KEY"
            Case dbeSQLBase
                NewSQLSource = SQLsource
            Case dbeSQLServer
                'NewSQLSource = "ALTER TABLE " & strOwner & "." & TableName & " DROP CONSTRAINT " & TableName & "_PK"
                NewSQLSource = "ALTER TABLE " & strOwner & "." & TableName & " DROP CONSTRAINT " & ConstraintName
            Case Else
        End Select
    ElseIf Mid(uSQLsource, iPos, Len("FOREIGN KEY")) = "FOREIGN KEY" Then
        ConstraintName = Mid(uSQLsource, iPos + 12)
        If Not doSQL("Delete From XSYSCONSTRAINTS Where TBNAME='" & TableName & "' And TBOWNER='" & strOwner & "' And TYPE='F' And CONSTRAINTNAME='" & ConstraintName & "'") Then GoTo ExitSub
        If Not doSQL("Delete From XSYSFKCONSTRAINTS Where TBNAME='" & TableName & "' And TBOWNER='" & strOwner & "' And CONSTRAINTNAME='" & ConstraintName & "'") Then GoTo ExitSub
        Select Case objDBE.DBEngine
            Case dbeAccess
                NewSQLSource = "ALTER TABLE " & strOwner & "." & TableName & " DROP FOREIGN KEY " & ConstraintName
            Case dbeSQLBase
                NewSQLSource = SQLsource
            Case dbeSQLServer
                NewSQLSource = "ALTER TABLE " & strOwner & "." & TableName & " DROP CONSTRAINT " & ConstraintName
            Case Else
        End Select
    Else
        'ColumnName = Mid(uSQLSource, iPos + 1)
        ColumnName = Mid(uSQLsource, iPos)
        If Not doSQL("Delete From XSYSCOLUMNS Where TBNAME='" & TableName & "' And TBOWNER='" & strOwner & "' And NAME='" & ColumnName & "'") Then GoTo ExitSub
        Select Case objDBE.DBEngine
            Case dbeSQLBase
                NewSQLSource = SQLsource
            Case dbeSQLServer
                NewSQLSource = "ALTER TABLE " & strOwner & "." & TableName & " DROP COLUMN " & ColumnName
            Case Else
        End Select
    End If
    
    'Execute the change to the database...
    fSuccess = ExecuteSQL(NewSQLSource)
    SQLsource = NewSQLSource
            
ExitSub:
    doAlterTableDrop = fSuccess
    DoEvents
    Call Trace(trcExit, "doAlterTableDrop", trcSQL)
    Exit Function
End Function
Private Function doAlterTableModify(SQLsource As String) As Boolean
    Dim adoRS As ADODB.Recordset
    Dim iPos As Integer
    Dim TableName As String
    Dim DataType As String
    Dim ColumnName As String
    Dim ConstraintName As String
    Dim fSuccess As Boolean
    Dim uSQLsource As String
    Dim NewSQLSource As String
    Dim strDefault As String
    Dim strPrecision As String
    Dim strScale As String
    Dim strStorage As String
    Dim strNulls As String
    Dim strOwner As String
            
    'SQLBase:    ALTER TABLE <TableName> MODIFY <ColumnName> [<DataType>] [NULL | NOT NULL | NOT NULL WITH DEFAULT]
    'SQL Server: ALTER TABLE <TableName> ALTER COLUMN <ColumnName> <DataType> [NULL | NOT NULL]
    
    Call Trace(trcEnter, "doAlterTableModify: " & SQLsource, trcSQL)
    
    fSuccess = False
    uSQLsource = Trim(UCase(SQLsource))
    TableName = ParseStr(uSQLsource, 3, " ")
    strOwner = objDBE.dbparmOwner
    If InStr(TableName, ".") Then
        strOwner = ParseStr(TableName, 1, ".")
        If strOwner <> objDBE.dbparmOwner And strOwner = UCase(objDBE.dbparmOwner) Then strOwner = objDBE.dbparmOwner
        TableName = ParseStr(TableName, 2, ".")
    End If
    ColumnName = ParseStr(uSQLsource, 5, " ")
    DataType = ParseStr(uSQLsource, 6, " ")
    
    Select Case DataType
        Case "NULL", "NOT"
            DataType = vbNullString
        Case Else
            If Not doDataType(objDBE.DBEngine, ColumnName, DataType, strDefault, strStorage, strPrecision, strScale) Then
                'Note: Errors are returned in DataType field...
                SQLsource = "Error: '" & DataType & "' encountered in ALTER TABLE command:" & vbCrLf & SQLsource
                GoTo ExitSub
            End If
            If Not doSQL("Update XSYSCOLUMNS Set COLTYPE='" & DataType & "', STORAGE=" & strStorage & ", LENGTH=" & strPrecision & ", SCALE=" & strScale & " Where TBNAME='" & TableName & "' And TBOWNER='" & strOwner & "' And NAME='" & ColumnName & "'") Then GoTo ExitSub
    End Select
    
    If Right(uSQLsource, Len("NOT NULL")) = "NOT NULL" Then
        strNulls = "N"
    ElseIf InStr(uSQLsource, " NOT NULL") > 0 And InStr(uSQLsource, " DEFAULT") > 0 Then
        strNulls = "D"
    Else
        strNulls = "Y"
    End If
    If Not doSQL("Update XSYSCOLUMNS Set NULLS='" & strNulls & "' Where TBNAME='" & TableName & "' And TBOWNER='" & strOwner & "' And NAME='" & ColumnName & "'") Then GoTo ExitSub
    
    Select Case objDBE.DBEngine
        Case dbeSQLBase
            NewSQLSource = SQLsource
        Case dbeSQLServer
            'As best I can tell, SQL Server's ALTER COLUMN clause requires the datatype... so
            'if we weren't passed the information we need (like in the case where we're just
            'setting the field to NOT NULL, for example), we have to go get data type information
            'here before continuing...
            If DataType = vbNullString Then
                Set adoRS = New ADODB.Recordset
                adoRS.Open "Select COLTYPE, LENGTH, SCALE From XSYSCOLUMNS Where TBNAME='" & TableName & "' And TBOWNER='" & strOwner & "' And NAME='" & ColumnName & "'", objDBE.ActiveConnection, adOpenKeyset, adLockReadOnly
                DataType = adoRS("COLTYPE")
                strPrecision = adoRS("LENGTH")
                strScale = adoRS("SCALE")
                Call CloseRecordset(adoRS, False)
            End If
            If Right(uSQLsource, Len("NOT NULL")) = "NOT NULL" Then
                strNulls = "NOT NULL"
            ElseIf Right(uSQLsource, Len("NOT NULL WITH DEFAULT")) = "NOT NULL WITH DEFAULT" Then
                'SQL Server doesn't support ALTER COLUMN with DEFAULT specification...
                'or does it...?
                'The syntax of ALTER TABLE (T-SQL) implies that a DEFAULT is simply a built-in constraint
                'and that it must be removed prior to reassignment, so maybe you could (with a little work),
                'but we won't for now (not until we really need it)...
                strNulls = "NOT NULL"
            Else
                strNulls = "NULL"
            End If
            DataType = ReassembleDataType(DataType, strPrecision, strScale)
            NewSQLSource = "ALTER TABLE " & strOwner & "." & TableName & " ALTER COLUMN " & ColumnName & " " & DataType & " " & strNulls
        Case Else
    End Select
    
    'Execute the change to the database...
    fSuccess = ExecuteSQL(NewSQLSource)
    SQLsource = NewSQLSource
            
ExitSub:
    doAlterTableModify = fSuccess
    DoEvents
    Call Trace(trcExit, "doAlterTableModify", trcSQL)
    Exit Function
End Function
Private Function doAlterTableRename(SQLsource As String) As Boolean
    Dim adoRS As ADODB.Recordset
    Dim TableName As String
    Dim ColumnName As String
    Dim fSuccess As Boolean
    Dim uSQLsource As String
    Dim NewSQLSource As String
    Dim NewName As String
    Dim strOwner As String

    'SQLBase:    ALTER TABLE <TableName> RENAME <ColumnName> <NewName>
    '            ALTER TABLE <TableName> RENAME TABLE <NewName>
    'SQL Server: sp_rename '<TableName>', '<NewName>'
    '            sp_rename '<TableName>.<ColumnName>', '<NewName>', 'COLUMN'
    
    Call Trace(trcEnter, "doAlterTableRename: " & SQLsource, trcSQL)
    
    fSuccess = False
    uSQLsource = Trim(UCase(SQLsource))
    TableName = ParseStr(uSQLsource, 3, " ")
    strOwner = objDBE.dbparmOwner
    If InStr(TableName, ".") Then
        strOwner = ParseStr(TableName, 1, ".")
        If strOwner <> objDBE.dbparmOwner And strOwner = UCase(objDBE.dbparmOwner) Then strOwner = objDBE.dbparmOwner
        TableName = ParseStr(TableName, 2, ".")
    End If
            
    ColumnName = ParseStr(uSQLsource, 5, " ")
    NewName = ParseStr(uSQLsource, 6, " ")
    If ColumnName = "TABLE" Then
        If Not doSQL("Update XSYSTABLES Set NAME='" & NewName & "' Where TBNAME='" & TableName & "' And OWNER='" & strOwner & "'") Then GoTo ExitSub
        Select Case objDBE.DBEngine
            Case dbeSQLBase
                NewSQLSource = SQLsource
            Case dbeSQLServer
                NewSQLSource = "sp_rename '" & strOwner & "." & TableName & "', '" & NewName & "'"
            Case Else
        End Select
    Else
        If Not doSQL("Update XSYSCOLUMNS Set NAME='" & NewName & "' Where TBNAME='" & TableName & "' And TBOWNER='" & strOwner & "' And NAME='" & ColumnName & "'") Then GoTo ExitSub
        Select Case objDBE.DBEngine
            Case dbeSQLBase
                NewSQLSource = SQLsource
            Case dbeSQLServer
                NewSQLSource = "sp_rename '" & strOwner & "." & TableName & "." & "[" & ColumnName & "]', '" & NewName & "', 'COLUMN'"
            Case Else
        End Select
    End If
    
    'Execute the change to the database...
    fSuccess = ExecuteSQL(NewSQLSource)
    SQLsource = NewSQLSource
            
ExitSub:
    doAlterTableRename = fSuccess
    DoEvents
    Call Trace(trcExit, "doAlterTableRename", trcSQL)
    Exit Function
End Function
Private Function doCreateIndex(SQLsource As String) As Boolean
    Dim adoRS As New ADODB.Recordset
    Dim i As Integer
    Dim iPos As Integer
    Dim TableName As String
    Dim UniqueRule As String
    Dim SortDirection As String
    Dim ColumnList As String
    Dim NewColumnList As String
    Dim ColumnName As String
    Dim ColumnNumber As Integer
    Dim IndexName As String
    Dim FillFactor As String
    Dim fSuccess As Boolean
    Dim uSQLsource As String
    Dim NewSQLSource As String
    Dim strOwner As String

    'SQLBase:    CREATE [UNIQUE] [CLUSTERED HASHED] INDEX <IndexName> ON <TableName> (<ColumnList-ASC/DESC>) [PCTFREE <#>] [SIZE <#> [ROWS | BUCKETS]]
    'SQL Server: CREATE [UNIQUE] [CLUSTERED | NONCLUSTERED] INDEX <IndexName> ON <TableName> (<ColumnList-ASC/DESC>) [WITH [PAD_INDEX] [[,]FILLFACTOR=<FillFactor>] [[,]IGNORE_DUP_KEY] [[,]DROP_EXISTING] [[,]STATISTICS_NORECOMPUTE]]
    
    Call Trace(trcEnter, "doCreateIndex: " & SQLsource, trcSQL)
    
    fSuccess = False
    uSQLsource = Trim(UCase(SQLsource))
    If ParseStr(uSQLsource, 2, " ") = "UNIQUE" Then UniqueRule = "U" Else UniqueRule = "D"
    
    uSQLsource = Trim(Mid(uSQLsource, InStr(uSQLsource, " INDEX ")))
    IndexName = ParseStr(uSQLsource, 2, " ")
    If InStr(IndexName, ".") Then IndexName = ParseStr(IndexName, 2, ".")
    strOwner = objDBE.dbparmOwner
    TableName = ParseStr(uSQLsource, 4, " ")
    If InStr(TableName, ".") Then
        strOwner = ParseStr(TableName, 1, ".")
        If strOwner <> objDBE.dbparmOwner And strOwner = UCase(objDBE.dbparmOwner) Then strOwner = objDBE.dbparmOwner
        TableName = ParseStr(TableName, 2, ".")
    End If
    i = InStr(TableName, "(")
    If i > 0 Then TableName = Mid(TableName, 1, i - 1)
    ColumnList = ParseStr(uSQLsource, 2, "(")
    ColumnList = Mid(ColumnList, 1, InStr(ColumnList, ")") - 1)
            
    If Not doSQL("Insert Into XSYSINDEXES (TBNAME, NAME, TBOWNER, UNIQUERULE, COLCOUNT, SYSTEM) Values ('" & TableName & "','" & IndexName & "','" & strOwner & "','" & UniqueRule & "'," & TokenCount(ColumnList, ",") & ",'N')") Then GoTo ExitSub
    
    NewColumnList = vbNullString
    For i = 1 To TokenCount(ColumnList, ",")
        ColumnName = Trim(ParseStr(ColumnList, i, ",", """"))
        SortDirection = Trim(ParseStr(ColumnName, 2, " "))
        ColumnName = Trim(ParseStr(ColumnName, 1, " "))
        Select Case SortDirection
            Case "ASC", vbNullString
                SortDirection = "A"
                NewColumnList = NewColumnList & ColumnName & objDBE.dbparmAsc
            Case "DESC"
                SortDirection = "D"
                NewColumnList = NewColumnList & ColumnName & objDBE.dbparmDesc
        End Select
        If i < TokenCount(ColumnList, ",") Then NewColumnList = NewColumnList & ","
        
        'Find the Column's COLNO...
                
        Call objDBE.OpenRecordset(adoRS, "Select COLNO From XSYSCOLUMNS Where TBNAME='" & TableName & "' And TBOWNER='" & strOwner & "' And NAME='" & ColumnName & "'", adOpenKeyset, adLockReadOnly)
        If Not doSQL("Insert Into XSYSKEYS (IXNAME, COLNAME, COLNO, COLSEQ, ORDERING) Values ('" & IndexName & "','" & ColumnName & "'," & adoRS("COLNO") & "," & i & ",'" & SortDirection & "')") Then GoTo ExitSub
        Call CloseRecordset(adoRS, False)
    Next i
    
    FillFactor = vbNullString
    iPos = InStr(uSQLsource, " PCTFREE ")
    If iPos > 0 Then
        FillFactor = Mid(uSQLsource, iPos + 1)
        FillFactor = ParseStr(FillFactor, 2, " ")
        'Since this value represents a percent-free value of the SQLBase index, FillFactor is
        'a percent-used value, so adjust it accordingly...
        FillFactor = 100 - Val(FillFactor)
    End If
    
    Select Case objDBE.DBEngine
        Case dbeSQLBase
            NewSQLSource = SQLsource
        Case dbeSQLServer
            NewSQLSource = "CREATE "
            If UniqueRule = "U" Then NewSQLSource = NewSQLSource & "UNIQUE "
            NewSQLSource = NewSQLSource & "INDEX " & IndexName & " ON " & strOwner & "." & TableName & " (" & NewColumnList & ")"
            If FillFactor <> vbNullString Then NewSQLSource = NewSQLSource & " WITH FILLFACTOR=" & FillFactor
        Case Else
    End Select
    
    'Execute the change to the database...
    fSuccess = ExecuteSQL(NewSQLSource)
    SQLsource = NewSQLSource
            
ExitSub:
    Call CloseRecordset(adoRS, True)
    doCreateIndex = fSuccess
    DoEvents
    Call Trace(trcExit, "doCreateIndex", trcSQL)
    Exit Function
End Function
Public Function doCreateTable(SQLsource As String, Description As String, DefaultField As String, DefaultMin As String, DefaultMax As String, Optional TableType As String = "T") As Boolean
    Dim adoRS As New ADODB.Recordset
    Dim i As Integer
    Dim iPos As Integer
    Dim sNum As Integer
    Dim TableName As String
    Dim DataType As String
    Dim ColumnDef As String
    Dim ColumnList As String
    Dim NewColumnList As String
    Dim ColumnName As String
    Dim ColumnNumber As Integer
    Dim FillFactor As String
    Dim fSuccess As Boolean
    Dim uSQLsource As String
    Dim NewSQLSource As String
    Dim strDefault As String
    Dim strOwner As String
    Dim strPrecision As String
    Dim strScale As String
    Dim strStorage As String
    Dim strNulls As String
    Dim strRemarks As String
    
    'SQLBase:    CREATE TABLE (<ColumnName> <DataType> [[] | [NOT NULL] | [NOT NULL WITH DEFAULT]],...) [PCTFREE <#>] !omitting referential integrity features...
    'SQL Server: CREATE TABLE [database_name.[owner]. | owner.] table_name{<column_definition> | column_name AS computed_column_expression | <table_constraint>} [,...n])
    '               <column_definition> ::= { column_name data_type } [ [ DEFAULT constant_expression ] | [ IDENTITY [(seed, increment ) [NOT FOR REPLICATION] ] ] ] [ ROWGUIDCOL ] [ <column_constraint>] [ ...n]
    '               <column_constraint> ::= [CONSTRAINT constraint_name] {[ NULL | NOT NULL ] | [    { PRIMARY KEY | UNIQUE } [CLUSTERED | NONCLUSTERED] [WITH FILLFACTOR = fillfactor] [ON {filegroup | DEFAULT} ]]

    Call Trace(trcEnter, "doCreateTable: " & SQLsource, trcSQL)
    Call objDBE.initDBEngineParms(objDBE.DBEngine)
    
    fSuccess = False
    uSQLsource = Trim(UCase(SQLsource))
    TableName = ParseStr(uSQLsource, 3, " ")
    strOwner = objDBE.dbparmOwner
    If InStr(TableName, ".") Then
        strOwner = ParseStr(TableName, 1, ".")
        If strOwner <> objDBE.dbparmOwner And strOwner = UCase(objDBE.dbparmOwner) Then strOwner = objDBE.dbparmOwner
        TableName = ParseStr(TableName, 2, ".")
        If strOwner <> objDBE.dbparmOwner Then TableType = "P"         'Assume temPorary table...
    End If
    If InStr(TableName, "(") > 0 Then TableName = ParseStr(TableName, 1, "(")
    ColumnList = Trim(Mid(uSQLsource, InStr(uSQLsource, TableName) + Len(TableName)))
    i = InStr(ColumnList, "PCTFREE")
    If i > 0 Then ColumnList = Trim(Mid(ColumnList, 1, i - 1))
    ColumnList = Mid(ColumnList, 2, Len(ColumnList) - 2)
            
    'If TableName = "FORMAT" Then Stop
    
    NewColumnList = vbNullString
    For i = 1 To TokenCount(ColumnList, ",")
        ColumnDef = Trim(ParseStr(ColumnList, i, ",", """"))
        ColumnName = ParseStr(ColumnDef, 1, " ")
        'DataType = ParseStr(ColumnDef, 2, " ")
        DataType = Trim(Mid(ColumnDef, InStr(ColumnDef, ColumnName) + Len(ColumnName)))
        If InStr(DataType, " ") Then
            'Some data types like DOUBLE PRECISION are made up of two words... try to strip these away from
            'additional NULL, or NOT NULL specifications...
            Select Case UCase(Left(ParseStr(DataType, 2, " "), 3))
                Case "PRE", "VAR", "CHA", "TEX"
                Case Else
                    DataType = ParseStr(DataType, 1, " ")
            End Select
        End If
        
        If Not doDataType(objDBE.DBEngine, ColumnName, DataType, strDefault, strStorage, strPrecision, strScale) Then
            'Note: Errors are returned in DataType field...
            SQLsource = "Error: '" & DataType & "' encountered in CREATE TABLE command:" & vbCrLf & SQLsource
            GoTo ExitSub
        End If
        NewColumnList = NewColumnList & ColumnName & " " & ReassembleDataType(DataType, strPrecision, strScale)
        
        If Right(ColumnDef, Len("NOT NULL")) = "NOT NULL" Then
            strNulls = "N"
            NewColumnList = NewColumnList & " NOT NULL"
        ElseIf InStr(ColumnDef, " NOT NULL") > 0 And InStr(ColumnDef, " DEFAULT") > 0 Then
            strNulls = "D"
            'This should never happen, but might with unusual data types...
            If strDefault <> vbNullString Then
                NewColumnList = NewColumnList & " NOT NULL " & objDBE.dbparmDefault & strDefault
            Else
                NewColumnList = NewColumnList & " NOT NULL"
            End If
        Else
            strNulls = "Y"
            NewColumnList = NewColumnList & " NULL"
        End If
        If i < TokenCount(ColumnList, ",") Then NewColumnList = NewColumnList & ","
        
        strRemarks = " "
        If ColumnName = "ROWID" Then strRemarks = "Unique record identifier used to handle concurrency."
        'if this is a Temp Table, skip this happy horseshit
        If strOwner = objDBE.dbparmOwner Then
            If Not doSQL("Insert Into XSYSCOLUMNS (TBNAME, NAME, TBOWNER, COLNO, COLTYPE, STORAGE, LENGTH, SCALE, NULLS, REMARKS) Values " & "('" & TableName & "','" & ColumnName & "','" & strOwner & "'," & i & ",'" & DataType & "'," & strStorage & "," & strPrecision & "," & strScale & ",'" & strNulls & "','" & FixQuotes(strRemarks) & "')") Then GoTo ExitSub
        End If
    Next i
    
    On Error Resume Next
    adoRS.Open "Select Max(SNUM)+1 From XSYSTABLES", objDBE.ActiveConnection, adOpenKeyset, adLockReadOnly
    If err.Number <> 0 Then
        err.Clear
        sNum = 1
    Else
        sNum = adoRS(0)
    End If
    Call CloseRecordset(adoRS, False)
    If Trim(DefaultField) = vbNullString Or DefaultField = "NULL" Or DefaultField = "''" Then DefaultField = "NULL" Else DefaultField = "'" & FixQuotes(DefaultField) & "'"
    If Trim(DefaultMin) = vbNullString Then DefaultMin = "NULL"
    If Trim(DefaultMax) = vbNullString Then DefaultMax = "NULL"
    'if this is a Temp Table, skip this happy horseshit
    If strOwner = objDBE.dbparmOwner Then
        If Not doSQL("Insert Into XSYSTABLES (NAME, COLCOUNT, REMARKS, TYPE, SYSTEM, SNUM, DATABASE_NAME, OWNER, DEFAULT_FIELD, DEFAULT_MIN, DEFAULT_MAX) Values (" & _
                "'" & TableName & "'," & TokenCount(ColumnList, ",") & ",'" & FixQuotes(Description) & "','" & TableType & "','N'," & sNum & ",'" & objDBE.DatabaseName & "','" & strOwner & "'," & DefaultField & "," & DefaultMin & "," & DefaultMax & ")") Then GoTo ExitSub
    End If
    FillFactor = vbNullString
    iPos = InStr(uSQLsource, " PCTFREE ")
    If iPos > 0 Then
        FillFactor = Mid(uSQLsource, iPos + 1)
        FillFactor = ParseStr(FillFactor, 2, " ")
        'Since this value represents a percent-free value of the SQLBase table, FillFactor is
        'a percent-used value, so adjust it accordingly...
        FillFactor = 100 - Val(FillFactor)
    End If
    
    Select Case objDBE.DBEngine
        Case dbeSQLBase
            NewSQLSource = SQLsource
        Case dbeSQLServer
            NewSQLSource = "CREATE TABLE " & strOwner & "." & TableName & " (" & NewColumnList & ")"
            'duh...
            'Turns out SQL Server provides for a FILLFACTOR to be specified when also specifying the
            'PRIMARY/Secondary KEY information... it does not apply to the data table itself...
            'If FillFactor <> vbNullString Then NewSQLSource = NewSQLSource & " WITH FILLFACTOR=" & FillFactor
        Case Else
    End Select
    
    'Execute the change to the database...
    fSuccess = ExecuteSQL(NewSQLSource)
    SQLsource = NewSQLSource
            
ExitSub:
    Call CloseRecordset(adoRS, True)
    doCreateTable = fSuccess
    DoEvents
    Call Trace(trcExit, "doCreateTable", trcSQL)
    Exit Function
End Function
Private Function doDropIndex(SQLsource As String) As Boolean
    Dim i As Integer
    Dim TableName As String
    Dim IndexName As String
    Dim fSuccess As Boolean
    Dim uSQLsource As String
    Dim NewSQLSource As String
    Dim strOwner As String
    
    'SQLBase:    DROP INDEX <IndexName>
    'SQL Server: DROP INDEX <TableName>.<IndexName>
    
    Call Trace(trcEnter, "doDropIndex: " & SQLsource, trcSQL)
    
    fSuccess = False
    uSQLsource = Trim(UCase(SQLsource))
    IndexName = ParseStr(uSQLsource, 3, " ")
    
    strOwner = ParseStr(IndexName, 1, ".")
    TableName = ParseStr(IndexName, 2, ".")
    IndexName = ParseStr(IndexName, 3, ".")
    If IndexName = vbNullString Then
        IndexName = TableName
        TableName = strOwner
        strOwner = objDBE.dbparmOwner
    End If
    If strOwner <> objDBE.dbparmOwner And strOwner = UCase(objDBE.dbparmOwner) Then strOwner = objDBE.dbparmOwner
    
    If Not doSQL("Delete From XSYSKEYS Where IXNAME='" & IndexName & "'") Then GoTo ExitSub
    If Not doSQL("Delete From XSYSINDEXES Where NAME='" & IndexName & "'") Then GoTo ExitSub
    
    Select Case objDBE.DBEngine
        Case dbeSQLBase
            NewSQLSource = "DROP INDEX " & IndexName
        Case dbeSQLServer
            NewSQLSource = "DROP INDEX " & strOwner & "." & TableName & "." & IndexName
        Case Else
    End Select
    
    'Execute the change to the database...
    fSuccess = ExecuteSQL(NewSQLSource)
    SQLsource = NewSQLSource
            
ExitSub:
    doDropIndex = fSuccess
    DoEvents
    Call Trace(trcExit, "doDropIndex", trcSQL)
    Exit Function
End Function
Private Function doDropTable(SQLsource As String) As Boolean
    Dim adoRS As New ADODB.Recordset
    Dim i As Integer
    Dim TableName As String
    Dim fSuccess As Boolean
    Dim uSQLsource As String
    Dim NewSQLSource As String
    Dim strOwner As String

    'SQLBase:    DROP TABLE <IndexName>
    'SQL Server: DROP TABLE <TableName>
    
    Call Trace(trcEnter, "doDropTable: " & SQLsource, trcSQL)
    
    On Error Resume Next
    fSuccess = False
    uSQLsource = Trim(UCase(SQLsource))
    TableName = ParseStr(uSQLsource, 3, " ")
    strOwner = objDBE.dbparmOwner
    If InStr(TableName, ".") Then
        strOwner = ParseStr(TableName, 1, ".")
        If strOwner <> objDBE.dbparmOwner And strOwner = UCase(objDBE.dbparmOwner) Then strOwner = objDBE.dbparmOwner
        TableName = ParseStr(TableName, 2, ".")
    End If
    
    'if this is a Temp Table skip all this happy horseshit
    If strOwner <> objDBE.dbparmOwner Then GoTo JustDropIt
    
    'Delete the PRIMARY KEY...
    If Not doSQL("ALTER TABLE " & strOwner & "." & TableName & " DROP PRIMARY KEY") Then GoTo ExitSub
    'Delete all FOREIGN KEYs...
    adoRS.Open "Select * From XSYSFKCONSTRAINTS Where TBNAME='" & TableName & "' And TBOWNER='" & strOwner & "'", objDBE.ActiveConnection, adOpenStatic, adLockReadOnly
    If err.Number = 0 Then
        While Not adoRS.EOF
            If Not doSQL("ALTER TABLE " & strOwner & "." & TableName & " DROP FOREIGN KEY " & adoRS("CONSTRAINTNAME")) Then GoTo ExitSub
            adoRS.MoveNext
        Wend
    End If
    err.Clear
    Call CloseRecordset(adoRS, False)
    'Delete any referencing FOREIGN KEYs from other tables...
    adoRS.Open "Select * From XSYSFKCONSTRAINTS Where REFDTBNAME='" & TableName & "' And REFDTBOWNER='" & strOwner & "'", objDBE.ActiveConnection, adOpenStatic, adLockReadOnly
    If err.Number = 0 Then
        While Not adoRS.EOF
            If Not doSQL("ALTER TABLE " & strOwner & "." & adoRS("TBNAME") & " DROP FOREIGN KEY " & adoRS("CONSTRAINTNAME")) Then GoTo ExitSub
            adoRS.MoveNext
        Wend
    End If
    err.Clear
    Call CloseRecordset(adoRS, False)
    'Delete all INDEXes...
    adoRS.Open "Select NAME From XSYSINDEXES Where TBNAME='" & TableName & "' And TBOWNER='" & strOwner & "'", objDBE.ActiveConnection, adOpenStatic, adLockReadOnly
    If err.Number = 0 Then
        While Not adoRS.EOF
            If Not doSQL("DROP INDEX " & strOwner & "." & TableName & "." & adoRS("NAME")) Then GoTo ExitSub
            adoRS.MoveNext
        Wend
        If Not doSQL("Delete From XSYSINDEXES Where TBNAME='" & TableName & "' And TBOWNER='" & strOwner & "'") Then GoTo ExitSub
    End If
    err.Clear
    Call CloseRecordset(adoRS, False)
    'Delete all column definitions...
    adoRS.Open "Select NAME From XSYSCOLUMNS Where TBNAME='" & TableName & "' And TBOWNER='" & strOwner & "'", objDBE.ActiveConnection, adOpenKeyset, adLockReadOnly
    If err.Number = 0 Then
        If Not doSQL("Delete From XSYSCOLUMNS Where TBNAME='" & TableName & "' And TBOWNER='" & strOwner & "'") Then GoTo ExitSub
    End If
    err.Clear
    Call CloseRecordset(adoRS, False)
    'Now delete the table itself...
    adoRS.Open "Select NAME From XSYSTABLES Where NAME='" & TableName & "' And OWNER='" & strOwner & "'", objDBE.ActiveConnection, adOpenKeyset, adLockReadOnly
    If err.Number = 0 Then
        If Not doSQL("Delete From XSYSTABLES Where NAME='" & TableName & "' And OWNER='" & strOwner & "'") Then GoTo ExitSub
    End If
    err.Clear
    Call CloseRecordset(adoRS, False)

JustDropIt:
    Select Case objDBE.DBEngine
        Case dbeSQLBase
            NewSQLSource = SQLsource
        Case dbeSQLServer
            NewSQLSource = SQLsource
        Case Else
    End Select
    
    'Execute the change to the database...
    fSuccess = ExecuteSQL(NewSQLSource)
    SQLsource = NewSQLSource
            
ExitSub:
    Call CloseRecordset(adoRS, True)
    doDropTable = fSuccess
    DoEvents
    Call Trace(trcExit, "doDropTable", trcSQL)
    Exit Function
End Function
'Public Function doSQL(ByVal SQLsource As String) As Boolean
Public Function doSQL(SQLsource As String) As Boolean
    Dim fSuccess As Boolean
    Dim iPos As Integer
    Dim ErrorString As String
    Dim ErrorCode As Long
    Dim Operation As String
    Dim strOwner As String
    Dim TableName As String
    Dim ColumnName As String
    Dim Remarks As String
    Dim uSQLsource As String
    
    Call Trace(trcEnter, "doSQL: " & SQLsource, trcSQL)
    Call objDBE.initDBEngineParms(objDBE.DBEngine)
    
    uSQLsource = UCase(SQLsource)
    Operation = ParseStr(uSQLsource, 1, " ")
    Select Case Operation
        Case "ALTER"    'TABLE
            fSuccess = doAlterTable(SQLsource)
        Case "COMMENT"  'ON
            Remarks = ParseStr(SQLsource, 4, " ")
            strOwner = ParseStr(Remarks, 1, ".")
            TableName = ParseStr(Remarks, 2, ".")
            ColumnName = ParseStr(Remarks, 3, ".")
            If ColumnName = vbNullString Then
                ColumnName = TableName
                TableName = strOwner
                strOwner = objDBE.dbparmOwner
            End If
            If strOwner <> objDBE.dbparmOwner And strOwner = UCase(objDBE.dbparmOwner) Then strOwner = objDBE.dbparmOwner
            iPos = InStr(SQLsource, "'")
            Remarks = Mid(SQLsource, iPos + 1, Len(SQLsource) - iPos - 1)
            If ParseStr(uSQLsource, 3, " ") = "COLUMN" Then
                If Not doSQL("Update XSYSCOLUMNS Set REMARKS='" & FixQuotes(Remarks) & "' Where TBNAME='" & TableName & "' And TBOWNER='" & strOwner & "' And NAME='" & ColumnName & "'") Then GoTo ExitSub
            ElseIf ParseStr(uSQLsource, 3, " ") = "TABLE" Then
                If Not doSQL("Update XSYSTABLES Set REMARKS='" & FixQuotes(Remarks) & "' Where NAME='" & TableName & "'And OWNER='" & strOwner & "' ") Then GoTo ExitSub
            End If
            If objDBE.DBEngine = dbeSQLBase Then
                fSuccess = ExecuteSQL(SQLsource)
            Else
                fSuccess = True
            End If
        Case "CREATE"
            Select Case ParseStr(uSQLsource, 2, " ")
                Case "TABLE"
                    fSuccess = doCreateTable(SQLsource, " ", " ", " ", " ")
                Case "INDEX", "UNIQUE"
                    fSuccess = doCreateIndex(SQLsource)
                Case "PUBLIC", "SYNONYM"
                    fSuccess = ExecuteSQL(SQLsource)
                Case Else
                    'Error...
                    SQLsource = "Unrecognized CREATE command:" & vbCrLf & SQLsource
                    GoTo ExitSub
            End Select
        Case "DROP"
            Select Case ParseStr(uSQLsource, 2, " ")
                Case "TABLE"
                    fSuccess = doDropTable(SQLsource)
                Case "INDEX"
                    fSuccess = doDropIndex(SQLsource)
                Case "PUBLIC", "SYNONYM"
                    fSuccess = ExecuteSQL(SQLsource)
                Case Else
                    'Error...
                    SQLsource = "Unrecognized DROP command:" & vbCrLf & SQLsource
                    GoTo ExitSub
            End Select
        Case Else
            fSuccess = ExecuteSQL(SQLsource)
    End Select
    
    If Not fSuccess Then
        'Note: Error is returned in SQLSource...
        Call objUI.Display("SQL Command Failed...", deBothStatusBarAndTextBox)
        Call objUI.Display(SQLsource, deTextBoxOnly)
    End If
    
ExitSub:
    Call Trace(trcExit, "doSQL", trcSQL)
    doSQL = fSuccess
    DoEvents
End Function
Public Function ExecuteSQL(SQLsource As String) As Boolean
    Dim ErrorCode As Long
    Dim ErrorString As String
    Dim saveErrNumber As Long
    Dim saveErrString As String
    
    Call Trace(trcEnter, "ExecuteSQL: " & SQLsource, trcSQL)
    Call objDBE.initDBEngineParms(objDBE.DBEngine)
    
    On Error GoTo ErrorHandler
    ExecuteSQL = False
    objDBE.ClearError
    Call objDBE.ExecuteSQL(SQLsource)
    ExecuteSQL = True
    Call Trace(trcExit, "ExecuteSQL", trcSQL)
    Exit Function
    
ErrorHandler:
    ErrorCode = objDBE.LastErrorCode
    ErrorString = objDBE.LastErrorString
    If ErrorCode = 0 And ErrorString = vbNullString Then
        ErrorString = BuildADOerror(objDBE.ActiveConnection, ErrorCode)
        If ErrorCode = 0 Then
            ErrorCode = err.Number
            ErrorString = err.Description
        End If
    End If
    
    Select Case ErrorCode
        Case 0
            ErrorString = "Unexpected error encountered executing SQL Statement: " & vbCr & vbCr & _
                SQLsource & vbCr & vbCr & _
                saveErrString & " (Error #" & saveErrNumber & ")"
            Call MsgBox(ErrorString, vbCritical)
        Case objDBE.dberrSynonymDoesNotExist    '00605 DIC SYN Synonym does not exist
            ExecuteSQL = True
        Case objDBE.dberrTableDoesNotExist, objDBE.dberrCannotDropTable, objDBE.dberrInvalidColumn, _
            objDBE.dberrIndexDoesNotExist, objDBE.dberrNotValidObject, objDBE.dberrPrimaryKeyNotFound, _
            objDBE.dberrForeignKeyNotFound, objDBE.dberrInvalidObject
            ExecuteSQL = True
        Case objDBE.dberrAlreadyDBA             '01105 SEC UAD User already has DBA authority
            ExecuteSQL = True
        Case Else
            ErrorString = "Unexpected error encountered executing SQL Statement: " & vbCr & vbCr & _
                SQLsource & vbCr & vbCr & _
                ErrorString
            Call MsgBox(ErrorString, vbCritical)
    End Select
    SQLsource = ErrorString
    Call Trace(trcBody, "Error: " & ErrorString, trcSQL)
    Call Trace(trcExit, "ExecuteSQL", trcSQL)
    Exit Function
End Function
Public Sub InitDisplay(DisplayControl As Object, StatusBarControl As Object)
    Set objUI.drtfDisplay = DisplayControl
    Set objUI.dsbStatus = StatusBarControl
End Sub
Public Function ReassembleDataType(ByVal strDataType As String, strPrecision As String, strScale As String) As String
    Call objDBE.initDBEngineParms(objDBE.DBEngine)
    If Val(strPrecision) > 0 And Val(strScale) > 0 Then
        strDataType = strDataType & "(" & strPrecision & "," & strScale & ")"
    ElseIf Val(strPrecision) > 0 Then
        Select Case UCase(strDataType)
            Case "INTEGER", "INT", "SMALLINT", "REAL", "FLOAT", "DOUBLE PRECISION", "TIMESTAMP", "DATETIME"
            Case Else
                strDataType = strDataType & "(" & strPrecision & ")"
        End Select
    End If
    ReassembleDataType = strDataType
End Function

