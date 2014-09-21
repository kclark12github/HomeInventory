VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDBEngine"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'clsDBEngine - clsDBEngine.cls
'   Database Engine-specific Module ...
'   Copyright © 2002, SunGard Investor Accounting Systems
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Problem:    Programmer:     Description:
'   02/21/02    63002       Ken Clark       Added logic in AbortTrans() to detect 'naturally occurring' dberrNoActiveTX errors;
'   01/29/02    None        Ken Clark       Changed  dbparmTimeStamp to "SYSDATETIME" from "SYSDATE" for SQLBase;
'   01/10/02    58817       Ken Clark       Defined dberrCannotDeleteRow for SQL Server;
'   01/07/02    None        Ken Clark       Effectively replaced FixROWID with the former FixCharROWID (renaming the old FixROWID
'                                           FixBinaryROWID and commenting it out, but retaining it for reference);
'   01/04/02    60893       JAD             Changed FixCharROWID to deal with Len(strROWID) instead of 8
'   12/28/01    60893       JAD             In SQLApplyAsClause, if AS is specified in original Token and it is not already in
'                                           single or double quotes, surround it in quotes
'   12/10/01    None        Ken Clark       Changed Public Functions returning a Boolean value to Subs since they all already used
'                                           error handling (err.Raise) to communicate failure back to calling routine, the return
'                                           value was simply redundant and confusing to the programmer and caused more coding than
'                                           was necessary;
'   12/05/01    None        Ken Clark       Added capability to use UID and PWD fields from file DSN to determine LoginID and
'                                           password for the system administrator user (typically SYSADM/sa);
'   11/13/01    None        Ken Clark       Replaced occurrences of fmtFormatDateTime with fmtShortDateTime
'                                           and fmtLongDate with fmtShortDateYYYY;
'   10/30/01    None        Ken Clark       Added dberrInvalidObject;
'   10/16/01    None        JAD             Changed adOpenKeyset to adOpenForwardOnly in SQLApplyAsClause.  Limited experience
'                                           indicates that the "Cannot create new transaction because capacity was exceeded.
'                                           (8004D01D)" error only affects recordsets that are opened and not closed before
'                                           the transaction starts.  In such a performance critical routine, we need to go with
'                                           speed until errors dictate otherwise.
'   10/12/01    None        Ken Clark       Moved FixDate from modDates;
'   10/01/01    58219       Ken Clark       Added dberrDuplicatePrimaryKey;
'   09/28/01    58219       Ken Clark       Added dberrStatementTerminated;
'                                           Changed SQLApplyAsClause not to treat quoted constants any differently than other
'                                           fields;
'                                           Introduced FixCharROWID;
'                                           Introduced FindROWID to address the multi-platform issues associated with adoRS.Find
'                                           functions on ROWID fields;
'                                           Removed dbpDeadlockSleepTime;
'                                           Replaced initDBErrors with initDBEngineParms;
'   09/27/01    58219       Ken Clark       Corrected parsing problems in SQLApplyAsClause and ParseSQLSelect regarding the use of
'                                           "Distinct";
'   09/26/01    58219       Ken Clark       Corrected connection problems in IsInDB for SQL Server databases. The way the "sp_who"
'                                           stored procedure was being executed was causing another connection to be opened on the
'                                           databases, causing the very situation we were trying to avoid. The code has been
'                                           changed to open a special connection to the "master" database long enough to execute
'                                           the "sp_who" stored procedure, and the resulting logic was altered to consider which
'                                           login name and connected database before returning that the user is indeed in the
'                                           connected database more than once.
'   09/25/01    58219       Ken Clark       SQL Server owner corrections ('SYSADM' vs 'dbo' case sensitivity issues);
'   09/24/01    None        Ken Clark       Added owner stuff to SQLApplyAsClause;
'   09/24/01    None        JAD             Added Trace messages on SQLApplyAsClause
'   09/23/01    None        Ken Clark       Assigned good value to dberrIndexDoesNotExist for SQL Server;
'   09/04/01    None        Ken Clark       Assigned good value to dberrInvalidPassword for SQL Server;
'                                           Removed Format statement in FixROWID as it was not deailing with Hex values properly;
'   08/31/01    None        Ken Clark       Added dberrTableDoesNotExist support;
'   08/15/01    None        Ken Clark       Added dberrInvalidColumn = 1401;
'                                           Added dberrIndexDoesNotExist = -320;
'   07/26/01    None        Ken Clark       Added dberrDBNotConfigured = 9279;
'                                           [MERANT][ODBC SQLBase driver][SQLBase]
'                                           09279 GCI EOR Cannot find specified protocol entry(80004005)
'   07/24/01    None        Ken Clark       Corrected SQL Server ROWID conversion problem in FixROWID;
'   06/29/01    None        Ken Clark       Added dbpDeadlockSleepTime;
'   06/07/01    None        Ken Clark       Added dberrDBShutdown;
'   06/06/01    None        Ken Clark       Replaced ErrorHandler's GoTo statements with Resume all routines with error handlers;
'   06/01/01    None        Ken Clark       Added dberrNotExecuted;
'   03/12/01    None        Ken Clark       Removed OldPW argument from GrantUser/DoGrantUser as it is not truly required as
'                                           initially thought;
'   03/09/01    None        Ken Clark       Split-out DoGrantUser from GrantUser so it can be called from DBUtility where it is
'                                           needed to operate on ID fields as they are read from a recordset without reconnecting
'                                           between each one);
'                                           Added ALTER PASSWORD functionality so as to package all such administrative functions
'                                           in one routine (else we'd be making special SYSADM connections in several places);
'   03/08/01    None        Ken Clark       Implemented SQLApplyAnyClause for SQL Server;
'   02/05/01    None        Ken Clark       Added dbeROWID();
'   02/04/01    None        Ken Clark       Moved GrantUser from libCenturaSQLAPI and refitted for use with SQL Server;
'   02/03/01    None        Ken Clark       Implemented IsInDB();
'   02/02/01    None        Ken Clark       Added FixROWID();
'   01/27/01    None        Ken Clark       Created;
'=================================================================================================================================
Option Explicit
Private mConstants As SIASDB.clsSIASDBConstants

Private adoConn As ADODB.Connection
Private dbeDBEngine As SIASDB.dbeDBEngineEnum
Private fAborting As Boolean
Private fDebugTXN As Boolean
Private intMaxDeadlockRetry As Integer
Private longActiveTXLevel As Long
Private longLastError As Long
Private objTrace As SIASUTL.clsTrace
Private saveErrNumber As Long
Private saveErrString As String
Private strDatabaseName As String
Private strDBUpperCaseFunctionName As String
Private strDBCurrentTimestamp As String
Private strLastError As String
Private strFileDSN As String
Private strFileDSNdir As String
Private strODBCDriverName As String
Private strServerName As String
Private strUserID As String

Public dberrAddForeignUnmatched As Long
Public dberrAlreadyDBA As Long
Public dberrCannotDeleteRow As Long
Public dberrCannotDropTable As Long
Public dberrCannotUpdateRow As Long
Public dberrDBLicenseExceeded As Long
Public dberrDBNotConfigured As Long
Public dberrDBShutdown As Long
Public dberrDeadLock As Long
Public dberrDuplicateKey As Long
Public dberrDuplicatePrimaryKey As Long
Public dberrForeignKeyNotFound As Long
Public dberrInconsistentView As Long
Public dberrIndexDoesNotExist As Long
Public dberrInvalidColumn As Long
Public dberrInvalidObject As Long
Public dberrInvalidPassword As Long
Public dberrInvalidUserID As Long
Public dberrLockTimeOut As Long
Public dberrNoNestedInsert As Long
Public dberrNotExecuted As Long
Public dberrNotValidObject As Long
Public dberrPrimaryKeyNotFound As Long
Public dberrSessionEndedAbnormally As Long
Public dberrSessionTerminated As Long
Public dberrStatementTerminated As Long
Public dberrSynonymDoesNotExist As Long
Public dberrTableDoesNotExist As Long
Public dberrUnmatchedForeignKey As Long

Public dbparmAsc As String
Public dbparmComment As String
Public dbparmConcatenator As String
Public dbparmDefault As String
Public dbparmDesc As String
Public dbparmGo As String
Public dbparmOwner As String
Public dbparmTerm As String
Public dbparmTimeStamp As String

Public Enum DBEConnectionStringEnum
    dbecseDatabase = 0
    dbecseServer = 1
    dbecseDriver = 2
    dbecseUserID = 3
    dbecsePassword = 4
End Enum
Private Sub Class_Initialize()
    Set mConstants = New SIASDB.clsSIASDBConstants
    
    Set adoConn = Nothing
    fAborting = False
    fDebugTXN = False
    intMaxDeadlockRetry = 0
    longActiveTXLevel = 0
    longLastError = 0
    Set objTrace = Nothing
    saveErrNumber = 0
    saveErrString = vbNullString
    strDatabaseName = vbNullString
    dbeDBEngine = dbeUnknown
    strDBUpperCaseFunctionName = vbNullString
    strDBCurrentTimestamp = vbNullString
    strLastError = vbNullString
    strFileDSN = vbNullString
    strFileDSNdir = vbNullString
    strODBCDriverName = vbNullString
    strServerName = vbNullString
    
    dberrAddForeignUnmatched = 0
    dberrAlreadyDBA = 0
    dberrCannotDeleteRow = 0
    dberrCannotDropTable = 0
    dberrCannotUpdateRow = 0
    dberrDBLicenseExceeded = 0
    dberrDBNotConfigured = 0
    dberrDBShutdown = 0
    dberrDeadLock = 0
    dberrDuplicateKey = 0
    dberrDuplicatePrimaryKey = 0
    dberrForeignKeyNotFound = 0
    dberrInconsistentView = 0
    dberrIndexDoesNotExist = 0
    dberrInvalidColumn = 0
    dberrInvalidObject = 0
    dberrInvalidPassword = 0
    dberrInvalidUserID = 0
    dberrLockTimeOut = 0
    dberrNoNestedInsert = 0
    dberrNotExecuted = 0
    dberrNotValidObject = 0
    dberrPrimaryKeyNotFound = 0
    dberrSessionEndedAbnormally = 0
    dberrSessionTerminated = 0
    dberrStatementTerminated = 0
    dberrSynonymDoesNotExist = 0
    dberrTableDoesNotExist = 0
    dberrUnmatchedForeignKey = 0
    
    dbparmAsc = vbNullString
    dbparmComment = vbNullString
    dbparmConcatenator = vbNullString
    dbparmDefault = vbNullString
    dbparmDesc = vbNullString
    dbparmGo = vbNullString
    dbparmOwner = vbNullString
    dbparmTerm = vbNullString
    dbparmTimeStamp = vbNullString
End Sub
Private Sub Class_Terminate()
    Set mConstants = Nothing
    Set objTrace = Nothing
    Call CloseConnection(adoConn, True)
End Sub
Public Sub AbortTrans()
    Dim ErrorCode As Long
    
    On Error GoTo ErrorHandler
    Call Trace(trcEnter, "AbortTrans()", trcApplicationDetail)
    If fAborting Then GoTo ExitSub
    
    If fDebugTXN Then Call MsgBox(Format(Now, "General Date") & " About to .RollbackTrans...", vbInformation)
    Call Trace(trcBody, "AbortTrans() - .RollbackTrans", trcSQL)
    fAborting = True
    adoConn.RollbackTrans
    Call Trace(trcBody, "AbortTrans() - .RollbackTrans Complete", trcSQL)
    If fDebugTXN Then Call MsgBox(Format(Now, "General Date") & " .RollbackTrans complete...", vbInformation)
    If longActiveTXLevel = 0 Then   'JAD1105'
        err.Raise dberrDBENoActiveTXN, mConstants.App.EXEName & " (clsDBEngine.AbortTrans)", Replace(mConstants.dberrDBENoActiveTXNStr, "%1", "Abort")
    Else
        longActiveTXLevel = longActiveTXLevel - 1
    End If
    fAborting = False
    ClearError
    
ExitSub:
    'Theoretically, we never leave here unsuccessfully...
    Call Trace(trcExit, "AbortTrans()", trcApplicationDetail)
    Exit Sub

ErrorHandler:
    Select Case err.Number
        'Skip any errors we know we raised ourselves...
        Case dberrDBENoActiveTXN
        Case Else
            saveErrNumber = err.Number
            saveErrString = err.Description
            strLastError = BuildADOerror(adoConn, longLastError)
            If longLastError = 0 Then
                'This is weird, but ADO doesn't seem to return an error code for this condition,
                'so we search the error string from the provider(?) to see if we have a real
                'database-level dberrDBENoActiveTXN condition...
                If InStr(UCase(strLastError), "NO TRANSACTION IS ACTIVE") > 0 Then
                    Call Trace(trcBody, "AbortTrans() - No transaction detected, .RollbackTrans failed.", trcSQL)
                    If longActiveTXLevel > 0 Then longActiveTXLevel = longActiveTXLevel - 1
                    err.Raise dberrDBENoActiveTXN, mConstants.App.EXEName & " (clsDBEngine.AbortTrans)", Replace(mConstants.dberrDBENoActiveTXNStr, "%1", "Abort")
                    Resume ExitSub
                End If
                longLastError = saveErrNumber
                strLastError = saveErrString
            End If
            '?? we need to go someplace here ???
            Select Case longLastError
                Case dberrSessionTerminated, dberrSessionEndedAbnormally, dberrDBShutdown
                    longActiveTXLevel = longActiveTXLevel - 1
                Case Else
            End Select
    End Select
    Call Trace(trcExit, "AbortTrans() - ErrorHandler", trcApplicationDetail)
    err.Raise longLastError, mConstants.App.EXEName & " (clsDBEngine.AbortTrans)", strLastError
    Resume ExitSub
End Sub
Public Property Get ActiveConnection() As ADODB.Connection
    Set ActiveConnection = adoConn
End Property
Public Property Set ActiveConnection(Value As ADODB.Connection)
    Set adoConn = Value
End Property
Public Property Get ActiveTXLevel() As Long
    ActiveTXLevel = longActiveTXLevel
End Property
Public Property Let ActiveTXLevel(Value As Long)
    longActiveTXLevel = Value
End Property
Public Sub BeginTrans()
    Dim ErrorCode As Long
    
    On Error GoTo ErrorHandler
    Call Trace(trcEnter, "BeginTrans()", trcApplicationDetail)
    
    If longActiveTXLevel > 0 Then
        err.Raise dberrDBETooManyActiveTXNs, mConstants.App.EXEName & " (clsDBEngine.BeginTrans)", mConstants.dberrDBETooManyActiveTXNsStr
    End If
    
    If fDebugTXN Then Call MsgBox(Format(Now, "General Date") & " About to .BeginTrans...", vbInformation + vbMsgBoxRight)
    Call Trace(trcBody, "BeginTrans() - .BeginTrans", trcSQL)
    longActiveTXLevel = adoConn.BeginTrans
    Call Trace(trcBody, "BeginTrans() - .BeginTrans Complete", trcSQL)
    If fDebugTXN Then Call MsgBox(Format(Now, "General Date") & " .BeginTrans complete...")
    ClearError
    
ExitSub:
    Call Trace(trcExit, "BeginTrans()", trcApplicationDetail)
    Exit Sub

ErrorHandler:
    Select Case err.Number
        'Skip any errors we know we raised ourselves...
        Case dberrDBETooManyActiveTXNs
        Case Else
            saveErrNumber = err.Number
            saveErrString = err.Description
            strLastError = BuildADOerror(adoConn, longLastError)
            If longLastError = 0 Then
                longLastError = saveErrNumber
                strLastError = saveErrString
            End If
    End Select
    Call AbortTrans
    Call Trace(trcExit, "BeginTrans() - ErrorHandler", trcApplicationDetail)
    err.Raise longLastError, mConstants.App.EXEName & " (clsDBEngine.BeginTrans)", strLastError
    Exit Sub
End Sub
Public Function BuildConnectionString(DSN As String, ByVal LoginID As String, ByVal Password As String) As String
    Dim objCentura As SIASDB.clsCentura
    If DSN <> vbNullString Then
        strODBCDriverName = GetINIKey(DSN, "ODBC", "Driver", vbNullString)
        Select Case Trim(Left(strODBCDriverName, 16))
            Case "Microsoft Access"
                BuildConnectionString = "FileDSN=" & DSN & ";"
                DBEngine = dbeAccess
                strServerName = GetINIKey(DSN, "ODBC", "DBQ", vbNullString)
                strDatabaseName = ParsePath(strServerName, FileNameBase)
                If LoginID = vbNullString Then
                    'Get LoginID from DSN file...
                    LoginID = GetINIKey(DSN, "ODBC", "UserID", "s" & "a")
                    Password = DecryptPassword(GetINIKey(DSN, "ODBC", "PWD", vbNullString))
                End If
            Case "Centura SQLBase"
                BuildConnectionString = "FileDSN=" & DSN & ";"
                DBEngine = dbeSQLBase
                strServerName = GetINIKey(DSN, "ODBC", "SRVR", vbNullString)
                strDatabaseName = GetINIKey(DSN, "ODBC", "DB", vbNullString)
                If LoginID = vbNullString Then
                    'Get LoginID from DSN file...
                    LoginID = GetINIKey(DSN, "ODBC", "UID", "S" & "Y" & "S" & "A" & "D" & "M")
                    Password = DecryptPassword(GetINIKey(DSN, "ODBC", "PWD", EncryptPassword("S" & "Y" & "S" & "A" & "D" & "M")))
                End If
                'Try to add the DB to the SQL.INI file... if it fails, continue...
                Set objCentura = New SIASDB.clsCentura
                Call objCentura.AddDBtoINI(DBEngine, strServerName, strDatabaseName)
            Case "SQL Server"
                DBEngine = dbeSQLServer
                strServerName = GetINIKey(DSN, "ODBC", "SERVER", vbNullString)
                strDatabaseName = GetINIKey(DSN, "ODBC", "DATABASE", vbNullString)
                If LoginID = vbNullString Then
                    'Get LoginID from DSN file...
                    LoginID = GetINIKey(DSN, "ODBC", "UID", "s" & "a")
                    Password = DecryptPassword(GetINIKey(DSN, "ODBC", "PWD", vbNullString))
                End If
                BuildConnectionString = "Provider=SQLOLEDB;Data Source=" & strServerName & ";Database=" & strDatabaseName & ";User ID=" & LoginID & ";Password=" & Password & ";"
        End Select
    Else
        'Don't have a DSN so wing it with some defaults...
        Select Case DBEngine
            Case dbeAccess
            Case dbeSQLBase
                strODBCDriverName = "Centura SQLBase 3.5 32-bit Driver -NT & Win95"
                If LoginID = vbNullString Then
                    LoginID = "S" & "Y" & "S" & "A" & "D" & "M"
                    Password = "S" & "Y" & "S" & "A" & "D" & "M"
                End If
                BuildConnectionString = "DRIVER=" & strODBCDriverName & ";UID=SYSADM;PWD=SYSADM;SRVR=" & strServerName & ";DB=" & strDatabaseName
            Case dbeSQLServer
                strODBCDriverName = DBEngine
                If LoginID = vbNullString Then
                    LoginID = "s" & "a"
                    Password = vbNullString
                End If
                BuildConnectionString = "Provider=SQLOLEDB;Data Source=" & strServerName & ";Database=" & strDatabaseName & ";User ID=" & LoginID & ";Password=" & Password & ";"
        End Select
    End If
    
ExitSub:
    Set objCentura = Nothing
    Exit Function
End Function
Public Sub ClearError()
    If Not adoConn Is Nothing Then adoConn.Errors.Clear
    strLastError = vbNullString
    longLastError = 0
End Sub
Public Sub Connect(ConnectionString As String)
    Dim i As Integer
    Dim SCRDriver As String
    
    On Error GoTo ErrorHandler
    Call Trace(trcEnter, "Connect()", trcApplicationDetail)
    
    If adoConn Is Nothing Then Set adoConn = New ADODB.Connection
    If (adoConn.State And adStateOpen) = adStateOpen Then Call Disconnect(False)

    adoConn.CommandTimeout = 60
    adoConn.ConnectionTimeout = 60
    adoConn.CursorLocation = adUseServer        'ADO Default...
    'adoConn.IsolationLevel = adXactIsolated     'ADO Default: adXactCursorStability
    adoConn.Mode = adModeReadWrite              'ADO Default: adModeUnknown
    adoConn.Open ConnectionString
    Select Case Trim(Left(ParseConnectionString(ConnectionString, dbecseDriver), 16))
        Case "Microsoft Access"
            DBEngine = dbeAccess
        Case "Centura SQLBase"
            DBEngine = dbeSQLBase
        Case "SQL Server", "SQLOLEDB"
            DBEngine = dbeSQLServer
        Case Else
            DBEngine = dbeUnknown
    End Select
    ServerName = ParseConnectionString(ConnectionString, dbecseServer)
    DatabaseName = ParseConnectionString(ConnectionString, dbecseDatabase)
    UserID = ParseConnectionString(ConnectionString, dbecseUserID)
    ClearError
    
ExitSub:
    Call Trace(trcExit, "Connect()", trcApplicationDetail)
    Exit Sub
    
ErrorHandler:
    Select Case err.Number
        'Skip any errors we know we raised ourselves...
        Case Else
            saveErrNumber = err.Number
            saveErrString = err.Description
            strLastError = BuildADOerror(adoConn, longLastError)
            If longLastError = 0 Then
                longLastError = saveErrNumber
                strLastError = saveErrString
            End If
    End Select
    Call Trace(trcExit, "Connect() - ErrorHandler", trcApplicationDetail)
    err.Raise longLastError, mConstants.App.EXEName & " (clsDBEngine.Connect)", strLastError
    Exit Sub
End Sub
Public Sub ConnectByFileDSN(DSN As String, ByVal LoginID As String, ByVal Password As String)
    Dim i As Integer
    Dim ConnectionString As String
    Dim SCRDriver As String
    
    On Error GoTo ErrorHandler
    Call Trace(trcEnter, "ConnectByFileDSN()", trcApplicationDetail)
    
    ConnectionString = BuildConnectionString(DSN, LoginID, Password)
    Call VerifyDSN(DBEngine, strODBCDriverName, strServerName, strDatabaseName)
    Call Connect(ConnectionString)
    strFileDSN = DSN
    Call VerifySCRoj
    ClearError
    
ExitSub:
    Call Trace(trcExit, "ConnectByFileDSN()", trcApplicationDetail)
    Exit Sub
    
ErrorHandler:
    Select Case err.Number
        'Skip any errors we know we raised ourselves...
        Case dberrDBEUnsupportedODBCDriver
        Case Else
            saveErrNumber = err.Number
            saveErrString = err.Description
            strLastError = BuildADOerror(adoConn, longLastError)
            If longLastError = 0 Then
                longLastError = saveErrNumber
                strLastError = saveErrString
            End If
    End Select
    Call Trace(trcExit, "ConnectByFileDSN() - ErrorHandler", trcApplicationDetail)
    err.Raise longLastError, mConstants.App.EXEName & " (clsDBEngine.ConnectByFileDSN)", strLastError
    Exit Sub
End Sub
Private Sub CreateRegistryDSN(RootKey As Long, KeyPath As String, ServerNameKey As String, ODBCDriver As String, DLLname As String, ServerName As String, DatabaseName As String)
    Dim ODBCListKey As String
    
    'CreateRegistryKey and SetRegistryKeyValue do their own error handling (i.e. call MsgBox themselves on errors)...
    CreateRegistryKey RootKey, KeyPath & "\" & mConstants.App.EXEName
    Select Case ODBCDriver
        Case "Microsoft Access Driver (*.mdb)"
            SetRegistryKeyValue RootKey, KeyPath & "\" & mConstants.App.EXEName, "DriverId", REG_DWORD, &H19
            SetRegistryKeyValue RootKey, KeyPath & "\" & mConstants.App.EXEName, "FIL", REG_SZ, "MS Access;"
            SetRegistryKeyValue RootKey, KeyPath & "\" & mConstants.App.EXEName, "PWD", REG_SZ, ""
            SetRegistryKeyValue RootKey, KeyPath & "\" & mConstants.App.EXEName, "SafeTransactions", REG_DWORD, &H0
            SetRegistryKeyValue RootKey, KeyPath & "\" & mConstants.App.EXEName, "UID", REG_SZ, ""
            CreateRegistryKey RootKey, KeyPath & "\" & mConstants.App.EXEName & "\Engines\Jet"
            SetRegistryKeyValue RootKey, KeyPath & "\" & mConstants.App.EXEName & "\Engines\Jet", "ImplicitCommitSync", REG_SZ, ""
            SetRegistryKeyValue RootKey, KeyPath & "\" & mConstants.App.EXEName & "\Engines\Jet", "MaxBufferSize", REG_DWORD, &H800
            SetRegistryKeyValue RootKey, KeyPath & "\" & mConstants.App.EXEName & "\Engines\Jet", "PageTimeout", REG_DWORD, &H5
            SetRegistryKeyValue RootKey, KeyPath & "\" & mConstants.App.EXEName & "\Engines\Jet", "Threads", REG_DWORD, &H3
            SetRegistryKeyValue RootKey, KeyPath & "\" & mConstants.App.EXEName & "\Engines\Jet", "UserCommitSync", REG_SZ, "Yes"
        Case "SQL Server"
            SetRegistryKeyValue RootKey, KeyPath & "\" & mConstants.App.EXEName, "LastUser", REG_SZ, "sa"
            SetRegistryKeyValue RootKey, KeyPath & "\" & mConstants.App.EXEName, "Regional", REG_SZ, "Yes"
        Case "CR SQLBase"
            SetRegistryKeyValue RootKey, KeyPath & "\" & mConstants.App.EXEName, "ApplicationUsingThreads", REG_SZ, "1"
            'SetRegistryKeyValue RootKey, KeyPath & "\" & mConstants.App.EXEName, "ColumnExpressions", REG_SZ, "1"
            SetRegistryKeyValue RootKey, KeyPath & "\" & mConstants.App.EXEName, "CursorCacheSize", REG_SZ, "6"
            SetRegistryKeyValue RootKey, KeyPath & "\" & mConstants.App.EXEName, "InputMessageSize", REG_SZ, "0"
            SetRegistryKeyValue RootKey, KeyPath & "\" & mConstants.App.EXEName, "LockTimeOut", REG_SZ, "300"
            SetRegistryKeyValue RootKey, KeyPath & "\" & mConstants.App.EXEName, "LogonID", REG_SZ, "SYSADM"
            SetRegistryKeyValue RootKey, KeyPath & "\" & mConstants.App.EXEName, "NoRecovery", REG_SZ, "0"
            SetRegistryKeyValue RootKey, KeyPath & "\" & mConstants.App.EXEName, "ReleasePlan", REG_SZ, "0"
            SetRegistryKeyValue RootKey, KeyPath & "\" & mConstants.App.EXEName, "Servers", REG_SZ, ""
        Case "Centura SQLBase 3.01 32-bit Driver -NT & Win95", _
             "Centura SQLBase 3.5 32-bit Driver -NT & Win95"
            SetRegistryKeyValue RootKey, KeyPath & "\" & mConstants.App.EXEName, "ApplicationUsingThreads", REG_SZ, "1"
            'SetRegistryKeyValue RootKey, KeyPath & "\" & mConstants.App.EXEName, "ColumnExpressions", REG_SZ, "1"
            SetRegistryKeyValue RootKey, KeyPath & "\" & mConstants.App.EXEName, "CursorCacheSize", REG_SZ, "6"
            SetRegistryKeyValue RootKey, KeyPath & "\" & mConstants.App.EXEName, "InputMessageSize", REG_SZ, "0"
            SetRegistryKeyValue RootKey, KeyPath & "\" & mConstants.App.EXEName, "LockTimeOut", REG_SZ, "300"
            SetRegistryKeyValue RootKey, KeyPath & "\" & mConstants.App.EXEName, "LogonID", REG_SZ, "SYSADM"
            SetRegistryKeyValue RootKey, KeyPath & "\" & mConstants.App.EXEName, "NoRecovery", REG_SZ, "0"
            SetRegistryKeyValue RootKey, KeyPath & "\" & mConstants.App.EXEName, "ReleasePlan", REG_SZ, "0"
            SetRegistryKeyValue RootKey, KeyPath & "\" & mConstants.App.EXEName, "Servers", REG_SZ, ""
        Case Else
            err.Raise dberrDBEUnsupportedODBCDriver, mConstants.App.EXEName & " (clsDBEngine.CreateRegistryDSN)", Replace(mConstants.dberrDBEUnsupportedODBCDriverStr, "%1", ODBCDriver)
    End Select
    SetRegistryKeyValue RootKey, KeyPath & "\" & mConstants.App.EXEName, ServerNameKey, REG_SZ, ServerName
    SetRegistryKeyValue RootKey, KeyPath & "\" & mConstants.App.EXEName, "Driver", REG_SZ, DLLname
    SetRegistryKeyValue RootKey, KeyPath & "\" & mConstants.App.EXEName, "Description", REG_SZ, "FiRRe User DSN created by FiRRe - " & Format(Now, "mmmm d, yyyy hh:nn AMPM")
    SetRegistryKeyValue RootKey, KeyPath & "\" & mConstants.App.EXEName, "Database", REG_SZ, DatabaseName
    
    'Finally, update the list of ODBC Data Sources to include the one we just added...
    If GetRegistryKeyValue(RootKey, KeyPath, "ODBC Data Sources") = vbNullString Then
        CreateRegistryKey RootKey, KeyPath & "\ODBC Data Sources"
    End If
    SetRegistryKeyValue RootKey, KeyPath & "\ODBC Data Sources", mConstants.App.EXEName, REG_SZ, ODBCDriver
End Sub
Public Property Get DatabaseName() As String
    DatabaseName = strDatabaseName
End Property
Public Property Let DatabaseName(Value As String)
    strDatabaseName = Value
End Property
Public Property Get DBCurrentTimestamp() As String
    DBCurrentTimestamp = strDBCurrentTimestamp
End Property
Public Property Get DBEngine() As SIASDB.dbeDBEngineEnum
    DBEngine = dbeDBEngine
End Property
Public Property Let DBEngine(NewValue As SIASDB.dbeDBEngineEnum)
    dbeDBEngine = NewValue
End Property
Public Property Get DBUpperCaseFunctionName() As String
    DBUpperCaseFunctionName = strDBUpperCaseFunctionName
End Property
Public Property Get DebugTXN() As Boolean
    DebugTXN = fDebugTXN
End Property
Public Property Let DebugTXN(NewValue As Boolean)
    fDebugTXN = NewValue
End Property
Public Sub Disconnect(Destroy As Boolean)
    If Not adoConn Is Nothing Then Call CloseConnection(adoConn, Destroy)
    ClearError
End Sub
Public Sub DoGrantUser(cn As ADODB.Connection, ID As String, PW As String, DBA As Boolean)
    Dim adoRS As ADODB.Recordset
    
    On Error Resume Next
    Select Case DBEngine
        Case dbeSQLBase
            cn.Execute "GRANT CONNECT TO " & ID & " IDENTIFIED BY " & PW
            If DBA Then cn.Execute "GRANT DBA TO " & ID
            If cn.Errors.Count > 0 Then
                If cn.Errors(0).NativeError <> dberrAlreadyDBA Then GoTo ErrorExit  'Already DBA (warning)...
                cn.Errors.Clear
            End If
            
'The following code has been removed as the "GRANT CONNECT TO <id> IDENTIFIED BY <pw>" seems to effectively
'change the user's password...
'            'Update the SQLBase password...
'            If PW <> OldPW Then
'                cn.Execute "ALTER PASSWORD " & OldPW & " TO " & PW
'                If err.Number <> 0 Then GoTo ErrorExit
'                cn.Errors.Clear
'            End If
        Case dbeSQLServer
            'See if a Login exists representing our Operator...
            Set adoRS = cn.Execute("sp_helplogins '" & ID & "'")
            If Not adoRS.EOF And err.Number = 0 And cn.Errors.Count = 0 Then
                'Login exists, so reset the password to what we have on file...
                Call CloseRecordset(adoRS, True)
                Set adoRS = cn.Execute("sp_password NULL, '" & PW & "', '" & ID & "'")
                If cn.Errors.Count > 0 Then
                    If cn.Errors(0).NativeError <> 15478 Then GoTo ErrorExit    'Password changed...
                    cn.Errors.Clear
                End If
            Else
                'Login doesn't exist, so add it...
                err.Clear
                Call CloseRecordset(adoRS, True)
                Set adoRS = cn.Execute("sp_addlogin '" & ID & "', '" & PW & "'")
                If cn.Errors.Count > 0 Then
                    If cn.Errors(0).NativeError <> 15298 Then GoTo ErrorExit     'New login created....
                    cn.Errors.Clear
                End If
            End If
            Call CloseRecordset(adoRS, True)
            
            'See of a User exists for our Operator/Login...
            Set adoRS = cn.Execute("sp_helpuser '" & ID & "'")
            If err.Number <> 0 Or cn.Errors.Count <> 0 Then
                'User does not exist, so add one... (let UserName and GroupName arguments default to LoginID)...
                err.Clear
                Call CloseRecordset(adoRS, True)
                Set adoRS = cn.Execute("sp_adduser '" & ID & "'")
                If cn.Errors.Count > 0 Then
                    If cn.Errors(0).NativeError <> 15341 Then GoTo ErrorExit    'Granted database access to '<ID>'...
                    cn.Errors.Clear
                End If
            End If
            Call CloseRecordset(adoRS, True)
        Case Else
    End Select
    ClearError

ExitSub:
    Call CloseRecordset(adoRS, True)
    Exit Sub
    
ErrorExit:
    strLastError = BuildADOerror(cn, longLastError)
    Call CloseRecordset(adoRS, True)
    On Error GoTo 0
    err.Raise longLastError, mConstants.App.EXEName & " (clsDBEngine.DoGrantUser)", strLastError
End Sub
Public Sub EndTrans()
    Dim ErrorCode As Long
    
    On Error GoTo ErrorHandler
    Call Trace(trcEnter, "EndTrans()", trcApplicationDetail)
    
    If fDebugTXN Then Call MsgBox(Format(Now, "General Date") & " About to .CommitTrans...", vbInformation)
    Call Trace(trcBody, "EndTrans() - .CommitTrans", trcSQL)
    adoConn.CommitTrans
    Call Trace(trcBody, "EndTrans() - .CommitTrans Complete", trcSQL)
    If fDebugTXN Then Call MsgBox(Format(Now, "General Date") & " .CommitTrans complete...", vbInformation)
    If longActiveTXLevel = 0 Then
        err.Raise dberrDBENoActiveTXN, mConstants.App.EXEName & " (clsDBEngine.EndTrans)", Replace(mConstants.dberrDBENoActiveTXNStr, "%1", "Commit")
    Else
        longActiveTXLevel = longActiveTXLevel - 1
    End If
    ClearError
    
ExitSub:
    Call Trace(trcExit, "EndTrans()", trcApplicationDetail)
    Exit Sub

ErrorHandler:
    Select Case err.Number
        'Skip any errors we know we raised ourselves...
        Case dberrDBENoActiveTXN
        Case Else
            saveErrNumber = err.Number
            saveErrString = err.Description
            strLastError = BuildADOerror(adoConn, longLastError)
            If longLastError = 0 Then
                longLastError = saveErrNumber
                strLastError = saveErrString
            End If
            '?? we need to go someplace here ???
            Select Case longLastError
                Case dberrSessionTerminated, dberrSessionEndedAbnormally, dberrDBShutdown
                    longActiveTXLevel = longActiveTXLevel - 1
                Case Else
            End Select
    End Select
    Call AbortTrans
    Call Trace(trcExit, "EndTrans() - ErrorHandler", trcApplicationDetail)
    err.Raise longLastError, mConstants.App.EXEName & " (clsDBEngine.EndTrans)", strLastError
    Exit Sub
End Sub
Public Sub ExecuteSQL(SQLsource As String)
    Dim RecordsAffected As Long
    
    On Error GoTo ErrorHandler
    Call Trace(trcBody, "ExecuteSQL() - .Execute """ & SQLsource & """, RecordsAffected", trcSQL)
    adoConn.Execute SQLsource, RecordsAffected
    Call Trace(trcBody, "ExecuteSQL() - Execute Complete; " & RecordsAffected & " Records Affected...", trcSQL)
    If InconsistentViewDetected(SQLsource, RecordsAffected) Then
        err.Raise dberrInconsistentView, mConstants.App.EXEName & " (clsDBEngine.ExecuteSQL)", "Someone else has modified this record." & vbCr & _
                    "Please re-start your transaction to view the latest record values."
    End If
    ClearError
    
ExitSub:
    Exit Sub
    
ErrorHandler:
    Select Case err.Number
        'Skip any errors we know we raised ourselves...
        Case dberrInconsistentView
            Call Trace(trcBody, "ExecuteSQL() - Inconsistent View Detected, raising error..." & vbCrLf & vbCrLf & _
                "SQL Statement: " & vbCrLf & SQLsource, trcSQL)
        Case Else
            saveErrNumber = err.Number
            saveErrString = err.Description
            strLastError = BuildADOerror(adoConn, longLastError)
            If longLastError = 0 Then
                longLastError = saveErrNumber
                strLastError = "Unexpected error encountered executing SQL Statement: " & vbCr & vbCr & _
                    SQLsource & vbCr & vbCr & _
                    saveErrString & " (Error #" & saveErrNumber & ")"
            End If
            Call Trace(trcBody, "ExecuteSQL() - ErrorHandler..." & vbCrLf & vbCrLf & _
                "Error: " & vbCrLf & strLastError & vbCrLf & _
                "SQL Statement: " & vbCrLf & SQLsource & vbCrLf, trcSQL)
    End Select
    err.Raise longLastError, mConstants.App.EXEName & " (clsDBEngine.ExecuteSQL)", strLastError
    Exit Sub
End Sub
Public Property Get FileDSN() As String
    FileDSN = strFileDSN
End Property
Public Property Let FileDSN(Value As String)
    strFileDSN = Value
End Property
Public Property Get FileDSNdir() As String
    FileDSNdir = strFileDSNdir
End Property
Public Property Let FileDSNdir(Value As String)
    strFileDSNdir = Value
End Property
Public Property Get LastErrorCode() As Long
    LastErrorCode = longLastError
End Property
Public Property Get LastErrorString() As String
    LastErrorString = strLastError
End Property
Public Function InconsistentViewDetected(ByVal SQLsource As String, ByVal RecordsAffected As Long) As Boolean
    Dim iPos As Integer

    InconsistentViewDetected = False
    If RecordsAffected > 0 Then Exit Function
    iPos = InStr(UCase(SQLsource), " WHERE ")
    If iPos > 0 Then
        If InStr(UCase(Mid(SQLsource, iPos)), "ROWID") > 0 Then InconsistentViewDetected = True
    End If
End Function
Public Sub MakeVirtualRecordset(SQLsource As String, ByRef vRS As ADODB.Recordset, Optional ByVal MaxRecords As Long = 0, Optional pRS As Variant)
    Dim adoRS As New ADODB.Recordset
    Dim fld As ADODB.Field
    Dim dupField As ADODB.Field
    Dim strTemp As String
    Dim FoundDup As Boolean
    Dim i As Integer
    Dim FieldCount As Integer
    Dim FieldList() As String
    Dim FieldType() As DataTypeEnum
    Dim FieldSize() As Long
    Dim FieldAttributes() As FieldAttributeEnum
    Dim FieldValues() As Variant
    
    On Error GoTo ErrorHandler
    Call Trace(trcEnter, "MakeVirtualRecordset()", trcApplicationDetail)
    
    'kfc - adoRS.MaxRecords = 1
    If SQLsource = vbNullString Then
        'Assume the programmer intends to copy pRS into vRS instead of opening a new Recordset...
        Set adoRS = Nothing
        Set adoRS = pRS
    Else
        adoRS.MaxRecords = MaxRecords
        'adoRS.Open sqlSource, adOpenStatic, adLockReadOnly
        'JAD This is really a boon to performance if it works ...
        'adOpenForwardOnly instead of adOpenStatic
        'frmACReview counts on a Static recordset
        If Not IsMissing(pRS) Then
            OpenRecordset adoRS, SQLsource, adOpenStatic, adLockReadOnly, False
        Else
            OpenRecordset adoRS, SQLsource, adOpenForwardOnly, adLockReadOnly, False
        End If
    End If
    FieldCount = adoRS.Fields.Count
    ReDim FieldList(0 To FieldCount - 1)
    ReDim FieldType(0 To FieldCount - 1)
    ReDim FieldSize(0 To FieldCount - 1)
    ReDim FieldAttributes(0 To FieldCount - 1)
    ReDim FieldValues(0 To FieldCount - 1)
    
    i = 0
    For Each fld In adoRS.Fields
        FieldList(i) = fld.Name
        FieldType(i) = fld.Type
        FieldSize(i) = fld.DefinedSize
        FieldAttributes(i) = fld.Attributes
        i = i + 1
    Next fld
    'kfc - CloseRecordset adoRS, False
    
    If Not vRS Is Nothing Then
        CloseRecordset vRS, False
        'kfc - Don't do this anymore 'cause vRS is being declared
        '"WithEvents" and setting him to Nothing and recreating him seems
        'to disrupt the event handling in the calling routine...
        'Set vRS = Nothing
    Else
        Set vRS = New ADODB.Recordset 'See comment above...
    End If
    vRS.CursorLocation = adUseClient    'May not matter, but adUseServer is default...
    vRS.CursorType = adOpenStatic       'Updatable snapshot
    vRS.LockType = adLockOptimistic     'Allow updates
    'This was causing errors in the vRS.Open
    'vRS.Source = sqlSource              '...for reference
    For i = 0 To FieldCount - 1
        'Before we append this new field, we need to ensure we won't get a duplicate error...
        FoundDup = False
        For Each dupField In vRS.Fields
            If FieldList(i) = dupField.Name Then
                FoundDup = True
                Exit For
            End If
        Next
        If FoundDup Then FieldList(i) = "Field" & i
        
        vRS.Fields.Append FieldList(i), FieldType(i), FieldSize(i), FieldAttributes(i)
        'Double-check the .Attribute property to be sure it's going to handle Nulls...
        If FieldAttributes(i) = 0 Then
            vRS.Fields(i).Attributes = adFldUpdatable + adFldIsNullable + adFldMayBeNull
        Else
            If Not CBool(vRS.Fields(i).Attributes And adFldUpdatable) Then vRS.Fields(i).Attributes = vRS.Fields(i).Attributes + adFldUpdatable
            If Not CBool(vRS.Fields(i).Attributes And adFldIsNullable) Then vRS.Fields(i).Attributes = vRS.Fields(i).Attributes + adFldIsNullable
            If Not CBool(vRS.Fields(i).Attributes And adFldMayBeNull) Then vRS.Fields(i).Attributes = vRS.Fields(i).Attributes + adFldMayBeNull
        End If
    Next i
    'Call adoDumpField(vRS.Fields(13), "", "NOPROP")
    vRS.Open
    'Call adoDumpField(vRS.Fields(13), "", "NOPROP")
    
    'Copy the data from the real recordset to the virtual one...
    'kfc - adoRS.MaxRecords = 0
    'kfc - adoRS.Open sqlSource, adOpenKeyset, adLockReadOnly
    If Not (adoRS.BOF And adoRS.EOF) Then
        adoRS.MoveFirst
        While Not adoRS.EOF
            'Populate the grid with the recordset data...
            
            vRS.AddNew
            For i = 0 To FieldCount - 1
                'FieldList(i) = adoRS(i).Name
                'Values(i) = adoRS(i).Value
                If Not CBool(vRS.Fields(i).Attributes And adFldIsNullable) And IsNull(adoRS(i).Value) Then
                    vRS(i).Value = Val(adoNullError)
                Else
                    vRS(i).Value = adoRS(i).Value
                End If
            Next i
            'vRS.AddNew FieldList, Values
            vRS.Update
            adoRS.MoveNext
        Wend
        vRS.MoveFirst
    End If
    If Not IsMissing(pRS) Then
        Set pRS = adoRS
        Set adoRS = Nothing
    Else
        CloseRecordset adoRS, True
    End If
    ClearError
    
ExitSub:
    Call Trace(trcExit, "MakeVirtualRecordset()", trcApplicationDetail)
    Exit Sub
    
ErrorHandler:
    Select Case err.Number
        'Skip any errors we know we raised ourselves...
        Case Else
            saveErrNumber = err.Number
            saveErrString = err.Description
            strLastError = BuildADOerror(adoConn, longLastError)
            If longLastError = 0 Then
                longLastError = saveErrNumber
                strLastError = saveErrString
            End If
            If longLastError = adoNullError Then strLastError = strLastError & vbCr & "Current Column: '" & vRS.Fields(i).Name & "'"
            strLastError = strLastError & vbCr & vbCr & "SQL Source: " & vbCr & SQLsource
    End Select
    Call Trace(trcExit, "MakeVirtualRecordset() - ErrorHandler", trcApplicationDetail)
    err.Raise longLastError, mConstants.App.EXEName & " (clsDBEngine.MakeVirtualRecordset)", strLastError
    Exit Sub
    Resume Next 'for debugging purposes...
End Sub
Public Property Get ODBCDriverName() As String
    ODBCDriverName = strODBCDriverName
End Property
Public Sub OpenRecordset(rs As ADODB.Recordset, ByVal SQLsource As String, CursorType As ADODB.CursorTypeEnum, LockType As ADODB.LockTypeEnum, Optional Create As Boolean = False)
    Dim i As Integer
    
    Call Trace(trcEnter, "OpenRecordset()", trcSQL)
    On Error GoTo ErrorHandler
    If Create Then Set rs = New ADODB.Recordset
    Select Case DBEngine
        Case dbeSQLServer
            'Parse-up the SQL statement in order to maintain the column names as SQLBase would have returned them...
            Call SQLApplyAsClause(SQLsource)
            rs.ActiveConnection = adoConn
            rs.Properties("Preserve On Commit") = True
            rs.Properties("Preserve On Abort") = True
            Call Trace(trcBody, "OpenRecordset() - .Open(""" & SQLsource & """, " & adoCursorType(CursorType) & "," & adoLockType(LockType) & ")", trcApplicationDetail + trcSQL)
            rs.Open SQLsource, , CursorType, LockType
        Case Else
            Call Trace(trcBody, "OpenRecordset() - .Open(""" & SQLsource & """, " & adoCursorType(CursorType) & "," & adoLockType(LockType) & ")", trcApplicationDetail + trcSQL)
            rs.Open SQLsource, adoConn, CursorType, LockType
    End Select
    Call Trace(trcBody, "OpenRecordset() - Complete; " & rs.RecordCount & " Records...", trcSQL)
    ClearError
    
ExitSub:
    Call Trace(trcExit, "OpenRecordset()", trcSQL)
    Exit Sub

ErrorHandler:
    Select Case err.Number
        'Skip any errors we know we raised ourselves...
        Case Else
            saveErrNumber = err.Number
            saveErrString = err.Description
            strLastError = BuildADOerror(adoConn, longLastError)
            If longLastError = 0 Then
                longLastError = saveErrNumber
                strLastError = saveErrString
                
                Call Trace(trcBody, "OpenRecordset() - ErrorHandler..." & vbCrLf & _
                    "Error: " & longLastError & "  " & strLastError & vbCrLf & vbCrLf & _
                    "SQL Statement: " & vbCrLf & SQLsource, trcApplicationDetail + trcSQL)
            Else
                Select Case longLastError
                    Case dberrSessionTerminated, dberrSessionEndedAbnormally, dberrDBShutdown
                        Call Trace(trcBody, "OpenRecordset() - Session Terminated..." & vbCrLf & _
                            "Error: " & longLastError & "  " & strLastError & vbCrLf & vbCrLf & _
                            "SQL Statement: " & vbCrLf & SQLsource, trcApplicationDetail + trcSQL)
                        'The connection is lost so we can't Abort or add records or anything!!!!
                    Case dberrDBEMysteryDisconnect  '-2147217887
                        'This seems to be the only error that is being returned here
                        'It is not an error from SQLBase (obviously, since it never gets there)
                        'It is an ADO / ODBC error ... we must assume that it is a lost connection
                        Call Trace(trcBody, "OpenRecordset() - Session Terminated (dberrDBEMysteryDisconnect)..." & vbCrLf & vbCrLf & _
                            "SQL Statement: " & vbCrLf & SQLsource & vbCrLf & strLastError, trcApplicationDetail + trcSQL)
                    Case Else
                        Call Trace(trcBody, "OpenRecordset() - Error: " & longLastError & "  " & strLastError & _
                            vbCrLf & vbCrLf & "SQL Statement: " & vbCrLf & SQLsource, _
                            trcApplicationDetail + trcSQL)
                End Select
            End If
            strLastError = strLastError & vbCr & vbCr & "SQL Source: " & vbCr & SQLsource
    End Select
    Call Trace(trcExit, "OpenRecordset() - ErrorHandler", trcApplicationDetail)
    err.Raise longLastError, mConstants.App.EXEName & " (clsDBEngine.OpenRecordset)", strLastError
    Exit Sub
    Resume Next 'for debugging purposes...
End Sub
Public Function ParseConnectionString(strConnection As String, ItemType As DBEConnectionStringEnum) As String
    Dim i As Integer
    Dim Token As String
    Dim ExtendedProperties As String
    Dim Quoted As Boolean
    
    If UCase(ParseStr(strConnection, 1, ";", """")) = "PROVIDER=SQLOLEDB" Then
        'Provider=SQLOLEDB;Data Source=WISDEV;Database=WildFiRRe;User ID=<UserID>;Password=<Password>;
        ExtendedProperties = strConnection
        Token = strConnection
    Else
        ExtendedProperties = vbNullString
        For i = 1 To TokenCount(strConnection, ";")
            Token = ParseStr(strConnection, i, ";", """")
            'Debug.Print "Token #" & i & ": " & Token
            If Mid(UCase(Token), 1, Len("EXTENDED PROPERTIES")) = "EXTENDED PROPERTIES" Then Exit For
        Next
        If Mid(UCase(Token), 1, Len("EXTENDED PROPERTIES")) <> "EXTENDED PROPERTIES" Then
            ParseConnectionString = vbNullString
            Exit Function
        End If
        ExtendedProperties = Mid(Token, Len("EXTENDED PROPERTIES=") + 1)
    End If
    
    i = 1
    While Token <> vbNullString
        Token = ParseStr(ExtendedProperties, i, ";", """")
        If Token <> vbNullString Then
            Select Case UCase(Mid(Token, 1, InStr(Token, "=") - 1))
                Case "DRIVER", "PROVIDER"
                    If ItemType = dbecseDriver Then
                        ParseConnectionString = Mid(Token, InStr(Token, "=") + 1)
                        Exit Function
                    End If
                Case "UID"
                    If ItemType = dbecseUserID Then
                        ParseConnectionString = Mid(Token, InStr(Token, "=") + 1)
                        Exit Function
                    End If
                Case "PWD"
                    If ItemType = dbecsePassword Then
                        ParseConnectionString = Mid(Token, InStr(Token, "=") + 1)
                        Exit Function
                    End If
                Case "SRVR"         'SQLBase
                    If ItemType = dbecseServer Then
                        ParseConnectionString = Mid(Token, InStr(Token, "=") + 1)
                        Exit Function
                    End If
                Case "SERVER"       'SQL Server
                    If ItemType = dbecseServer Then
                        ParseConnectionString = Mid(Token, InStr(Token, "=") + 1)
                        Exit Function
                    End If
                Case "DB"           'SQLBase
                    If ItemType = dbecseDatabase Then
                        ParseConnectionString = Mid(Token, InStr(Token, "=") + 1)
                        Exit Function
                    End If
                Case "DATABASE"     'SQL Server
                    If ItemType = dbecseDatabase Then
                        ParseConnectionString = Mid(Token, InStr(Token, "=") + 1)
                        Exit Function
                    End If
                Case "DATA SOURCE"  'SQL Server
                    If ItemType = dbecseServer Then
                        ParseConnectionString = Mid(Token, InStr(Token, "=") + 1)
                        Exit Function
                    End If
            End Select
        End If
        i = i + 1
    Wend
    
ExitSub:
    Exit Function
End Function
'Public Property Get dbeAccess() As String
'    dbeAccess = "Access"
'End Property
'Public Property Get dbeSQLBase() As String
'    dbeSQLBase = "SQLBase"
'End Property
'Public Property Get dbeSQLServer() As String
'    dbeSQLServer = "SQL Server"
'End Property
Public Property Get MaxDeadlockRetry() As Integer
    MaxDeadlockRetry = intMaxDeadlockRetry
End Property
Public Property Let MaxDeadlockRetry(NewValue As Integer)
    intMaxDeadlockRetry = MaxDeadlockRetry
End Property
Public Function dbeROWID(Optional TableName As String = vbNullString) As String
    Select Case DBEngine
        Case dbeSQLBase
            If TableName <> vbNullString Then dbeROWID = TableName & "."
            dbeROWID = dbeROWID & "ROWID"
        Case dbeSQLServer
            If TableName <> vbNullString Then
                dbeROWID = "CONVERT(CHAR(8)," & TableName & ".ROWID) as '" & TableName & ".ROWID'"
            Else
                dbeROWID = "CONVERT(CHAR(8),ROWID) as 'ROWID'"
            End If
        Case Else
    End Select
End Function
Public Sub FindROWID(rs As ADODB.Recordset, SearchField As String, strConvertedROWID As String, fMoveFirst As Boolean)
    If fMoveFirst Then rs.MoveFirst
    Do While Not rs.EOF
        If FixROWID(rs(SearchField)) = strConvertedROWID Then Exit Do
        rs.MoveNext
    Loop
End Sub
'Public Function FixBinaryROWID(vROWID As Variant) As String
'    Dim buffer As String
'    Dim fld As ADODB.Field
'    Dim i As Integer
'    Dim strByte As String
'
'    Select Case DBEngine
'        Case dbeSQLBase
'            If Left(vROWID, 1) <> "'" Then
'                FixROWID = "'" & vROWID & "'"
'            Else
'                FixROWID = vROWID
'            End If
'        Case dbeSQLServer
'            If Len(vROWID) = 18 And Left(vROWID, 2) = "0x" Then
'                'Assume the value's already been converted...
'                FixROWID = vROWID
'                Exit Function
'            End If
'            Set fld = vROWID
'            buffer = "0x"
'            For i = 1 To fld.ActualSize
'                'For some reason certain values are disregarded by the "00" format string, and a single
'                'hex digit is returned, screwing up the ROWID value... So do this in two steps...
'                strByte = Hex(Asc(Mid(fld.Value, i, 1)))
'                If Len(strByte) = 1 Then strByte = "0" & strByte
'                buffer = buffer & strByte
'            Next i
'            FixROWID = buffer
'        Case Else
'    End Select
'End Function
Public Function FixDate(ByVal strDate As Variant) As String
    'Assume that if the argument contains a space, it must be a timestamp, so format it
    'as fmtShortDateTime, otherwise use fmtShortDateYYYY... then deal with the fact that SQLBase
    'needs this in mm/dd/yyyy format...
    If InStr(strDate, " ") > 0 Then
        Select Case DBEngine
            Case dbeSQLBase
                strDate = "'" & FormatMDY(CDate(strDate)) & " " & Format(strDate, "HH:mm:ss") & "'"
            Case dbeSQLServer
                'Note: Untested in "English (United Kingdom)"
                strDate = "'" & Format(strDate, fmtShortDateTime) & "'"
            Case Else
        End Select
    Else
        Select Case DBEngine
            Case dbeSQLBase
                strDate = "'" & FormatMDY(CDate(strDate)) & "'"
            Case dbeSQLServer
                'Note: Untested in "English (United Kingdom)"
                strDate = "'" & Format(strDate, fmtShortDateYYYY) & "'"
            Case Else
        End Select
    End If
    
    FixDate = strDate
End Function
Public Function FixROWID(vROWID As Variant) As String
    Dim i As Integer
    Dim buffer As String
    Dim strByte As String
    Dim strROWID As String
    
    Select Case DBEngine
        Case dbeSQLBase
            If Left(vROWID, 1) <> "'" Then
                FixROWID = "'" & vROWID & "'"
            Else
                FixROWID = vROWID
            End If
        Case dbeSQLServer
            If Len(vROWID) = 18 And Left(vROWID, 2) = "0x" Then
                'Assume the value's already been converted...
                FixROWID = vROWID
                Exit Function
            End If
            'When we have a length of 10 bytes, the first and last appear to be ASC(27) or ESC...
            'So strip them out before processing...
            If Len(vROWID) = 10 Then strROWID = Mid(vROWID, 2, 8) Else strROWID = vROWID
            buffer = "0x"
            For i = 1 To Len(strROWID)  'was 8
                'For some reason certain values are disregarded by the "00" format string, and a single
                'hex digit is returned, screwing up the ROWID value... So do this in two steps...
                strByte = Hex(Asc(Mid(strROWID, i, 1)))
                If Len(strByte) = 1 Then strByte = "0" & strByte
                buffer = buffer & strByte
            Next i
            FixROWID = buffer
        Case Else
    End Select
End Function
Public Sub GrantUser(ID As String, PW As String, DBA As Boolean)
    Dim SYSADMconn As ADODB.Connection
    Dim strConnection As String
    Dim SAlogin As String
    Dim SApassword As String
    
    On Error GoTo ErrorHandler
    Call Trace(trcEnter, "GrantUser()", trcApplicationDetail)
    
    Set SYSADMconn = New ADODB.Connection
    SYSADMconn.CommandTimeout = 60
    SYSADMconn.ConnectionTimeout = 60
    SYSADMconn.CursorLocation = adUseServer        'ADO Default...
    'SYSADMconn.IsolationLevel = adXactIsolated     'ADO Default: adXactCursorStability
    SYSADMconn.Mode = adModeReadWrite              'ADO Default: adModeUnknown
    
    Select Case DBEngine
        Case dbeAccess
            'Get LoginID from DSN file...
            SAlogin = GetINIKey(strFileDSN, "ODBC", "UserID", "s" & "a")
            SApassword = DecryptPassword(GetINIKey(strFileDSN, "ODBC", "PWD", vbNullString))
            strConnection = "DRIVER=" & strODBCDriverName & ";UID=" & SAlogin & ";PWD=" & SApassword & ";SERVER=" & strServerName & ";DATABASE=" & strDatabaseName
        Case dbeSQLBase
            'Get LoginID from DSN file...
            SAlogin = GetINIKey(strFileDSN, "ODBC", "UID", "S" & "Y" & "S" & "A" & "D" & "M")
            SApassword = DecryptPassword(GetINIKey(strFileDSN, "ODBC", "PWD", EncryptPassword("S" & "Y" & "S" & "A" & "D" & "M")))
            strConnection = "DRIVER=" & strODBCDriverName & ";UID=" & SAlogin & ";PWD=" & SApassword & ";SRVR=" & strServerName & ";DB=" & strDatabaseName
        Case dbeSQLServer
            'Get LoginID from DSN file...
            SAlogin = GetINIKey(strFileDSN, "ODBC", "UID", "s" & "a")
            SApassword = DecryptPassword(GetINIKey(strFileDSN, "ODBC", "PWD", vbNullString))
            strConnection = "DRIVER=" & strODBCDriverName & ";UID=" & SAlogin & ";PWD=" & SApassword & ";SERVER=" & strServerName & ";DATABASE=" & strDatabaseName
    End Select
    
    SYSADMconn.Open strConnection, "", ""
    If DBEngine = dbeSQLBase Then SYSADMconn.BeginTrans  'SQL Server Login/User Operations cannot take place in a transaction...

    Call DoGrantUser(SYSADMconn, ID, PW, DBA)
    ClearError

ExitSub:
    On Error Resume Next
    'SQL Server Login/User Operations cannot take place in a transaction, but SQLBase ones do...
    If DBEngine = dbeSQLBase Then SYSADMconn.CommitTrans
    Call CloseConnection(SYSADMconn, True)
    Call Trace(trcExit, "GrantUser()", trcApplicationDetail)
    Exit Sub
    
ErrorHandler:
    Select Case err.Number
        'Skip any errors we know we raised ourselves...
        Case Else
            saveErrNumber = err.Number
            saveErrString = err.Description
            strLastError = BuildADOerror(adoConn, longLastError)
            If longLastError = 0 Then
                longLastError = saveErrNumber
                strLastError = saveErrString
            End If
    End Select
    
    'SQL Server Login/User Operations cannot take place in a transaction, but SQLBase ones do...
    If DBEngine = dbeSQLBase Then SYSADMconn.RollbackTrans
    Call CloseConnection(SYSADMconn, True)
    Call Trace(trcExit, "GrantUser() - ErrorHandler", trcApplicationDetail)
    err.Raise longLastError, mConstants.App.EXEName & " (clsDBEngine.GrantUser)", strLastError
    Exit Sub
    Resume Next 'for debugging purposes...
End Sub
Public Sub initDBEngineParms(DBEngine As SIASDB.dbeDBEngineEnum)
    Select Case DBEngine
        Case dbeAccess
            dbparmAsc = vbNullString
            dbparmComment = "! "
            dbparmConcatenator = "&"
            dbparmDefault = vbNullString
            dbparmDesc = vbNullString
            dbparmGo = vbNullString
            dbparmOwner = vbNullString
            dbparmTerm = ";"
            dbparmTimeStamp = "#" & Now & "#"
            
            'Assign Access Database Error Codes... (negative numbers represent unknown values)...
            dberrAddForeignUnmatched = -388
            dberrAlreadyDBA = -1105
            dberrCannotDeleteRow = -383
            dberrCannotDropTable = -601
            dberrCannotUpdateRow = -386
            dberrDBLicenseExceeded = -2522
            dberrDBNotConfigured = -9279
            dberrDBShutdown = -1315
            dberrDeadLock = -1801
            dberrDuplicateKey = -805
            dberrDuplicatePrimaryKey = dberrDuplicateKey
            dberrForeignKeyNotFound = -1421
            dberrInconsistentView = -806
            dberrIndexDoesNotExist = -320
            dberrInvalidColumn = -1401
            dberrInvalidObject = -208
            dberrInvalidPassword = -404
            dberrInvalidUserID = -405
            dberrLockTimeOut = -1805
            dberrNoNestedInsert = -384
            dberrNotExecuted = -203
            dberrNotValidObject = &H80040E37         'Table not found...
            dberrPrimaryKeyNotFound = -1422
            dberrSessionEndedAbnormally = -9024
            dberrSessionTerminated = -3806
            dberrStatementTerminated = -3621
            dberrSynonymDoesNotExist = -605
            dberrTableDoesNotExist = -4902
            dberrUnmatchedForeignKey = -382
        Case dbeSQLBase
            dbparmAsc = " ASC"
            dbparmComment = "remark "
            dbparmConcatenator = "||"
            dbparmDefault = "WITH DEFAULT"
            dbparmDesc = " DESC"
            dbparmGo = vbNullString
            dbparmOwner = "SYSADM"
            dbparmTerm = ";"
            dbparmTimeStamp = "SYSDATETIME"     '"SYSDATE"
        
            'Assign SQLBase Database Error Codes... (negative numbers represent unknown values)...
            dberrAddForeignUnmatched = 388
            dberrAlreadyDBA = 1105
            dberrCannotDeleteRow = 383
            dberrCannotDropTable = 601
            dberrCannotUpdateRow = 386
            dberrDBLicenseExceeded = 2522
            dberrDBNotConfigured = 9279         '[MERANT][ODBC SQLBase driver][SQLBase]09279 GCI EOR Cannot find specified protocol entry(80004005)
            dberrDBShutdown = 1315
            dberrDeadLock = 1801
            dberrDuplicateKey = 805
            dberrDuplicatePrimaryKey = dberrDuplicateKey
            dberrForeignKeyNotFound = 1421
            dberrInconsistentView = 806
            dberrIndexDoesNotExist = 320
            dberrInvalidColumn = 1401
            dberrInvalidObject = -208
            dberrInvalidPassword = 404
            dberrInvalidUserID = 405
            dberrLockTimeOut = 1805
            dberrNotExecuted = 203
            dberrNoNestedInsert = 384
            dberrNotValidObject = -3727         'No equivalence...
            dberrPrimaryKeyNotFound = 1422
            dberrSessionEndedAbnormally = 9024
            dberrSessionTerminated = 3806
            dberrStatementTerminated = -3621
            dberrSynonymDoesNotExist = 605
            dberrTableDoesNotExist = -4902
            dberrUnmatchedForeignKey = 382
        Case dbeSQLServer
            dbparmAsc = vbNullString
            dbparmComment = "-- "   'or /* xxx */
            dbparmConcatenator = "+"
            dbparmDefault = "DEFAULT "
            dbparmDesc = vbNullString
            dbparmGo = "GO" & dbparmTerm & vbCrLf
            dbparmOwner = "dbo"
            dbparmTerm = vbNullString
            dbparmTimeStamp = "CURRENT_TIMESTAMP"
    
            'Assign SQL Server Error Codes... (negative numbers represent unknown values)...
            dberrAddForeignUnmatched = 1770     '?
            dberrAlreadyDBA = -1105             '?
            dberrCannotDeleteRow = 547
            dberrCannotDropTable = 3701
            dberrCannotUpdateRow = 3725         '?
            dberrDBLicenseExceeded = 18458      '?
            dberrDBNotConfigured = -9279        '?
            dberrDBShutdown = -1315             '?
            dberrDeadLock = 16964               '?Maybe 7112 or 1219
            dberrDuplicateKey = 2601
            dberrDuplicatePrimaryKey = 2627
            dberrForeignKeyNotFound = -1421     '?
            dberrIndexDoesNotExist = 3703
            dberrInconsistentView = -806        '?
            dberrInvalidColumn = -1401          '?
            dberrInvalidObject = 208
            dberrInvalidPassword = 18456        '
            dberrInvalidUserID = 4002           '?
            dberrLockTimeOut = -1805            '?
            dberrNotExecuted = -203             '?
            dberrNoNestedInsert = 8164          '?
            dberrNotValidObject = 3727
            dberrPrimaryKeyNotFound = -1422     '?
            dberrSessionEndedAbnormally = -9024
            dberrSessionTerminated = 3911       '?
            dberrStatementTerminated = 3621
            dberrSynonymDoesNotExist = -605     'No equivalence...
            dberrTableDoesNotExist = 4902
            dberrUnmatchedForeignKey = 1769     '? or 1770
    End Select
    'Set the following code(s) for use in libADO... (which cannot access these "dberr" codes)...
    Call adoSetStatementTerminatedCode(dberrStatementTerminated)
End Sub
Public Function IsInDB(UserID As String) As Boolean
    Dim adoRS As ADODB.Recordset
    Dim adoCMD As ADODB.Command
    Dim vRS As ADODB.Recordset
    Dim SPID As Variant
    Dim masterConn As ADODB.Connection
    
    On Error Resume Next
    IsInDB = False
    Select Case DBEngine
        Case dbeSQLBase
            'This function isn't reliable, so just assume False until it is fixed (if ever)...
            'IsInDB = IsInSQLBase(strServerName, strDatabaseName, UserID, vbNullString)
            IsInDB = False
        Case dbeSQLServer
            'First get my SPID (i.e. Process ID)...
            Set adoRS = New ADODB.Recordset
            adoRS.Open "SELECT @@SPID AS 'SPID'", adoConn, adOpenStatic, adLockReadOnly
            SPID = adoRS("SPID")
            Call CloseRecordset(adoRS, True)
            
            'Now see how other processes are connected under this LoginID...
            Set masterConn = New ADODB.Connection
            masterConn.Open Replace(adoConn.ConnectionString, strDatabaseName, "master")
            
            
            Set adoCMD = New ADODB.Command
            adoCMD.CommandText = "sp_who '" & Trim(UserID) & "'"
            adoCMD.CommandType = adCmdText
            Set adoCMD.ActiveConnection = masterConn
            Set adoRS = adoCMD.Execute()
            'adoRS.Open adoCMD, , adOpenKeyset, adLockOptimistic
            'adoRS.Open "sp_who '" & Trim(UserID) & "'", masterConn, adOpenStatic, adLockReadOnly
            While (Not adoRS.EOF) And (Not IsInDB)
                If adoRS("spid") <> SPID And adoRS("dbname") = strDatabaseName Then IsInDB = True
                adoRS.MoveNext
            Wend
            Call CloseRecordset(adoRS, True)
            'Don't use CloseConnection() here because the connection doesn't seem to go away according
            'to the SQL Server Enterprise Manager... The following code does seem to work...
            '...then again, maybe it worked by coincidence...
            masterConn.Close
            Set masterConn = Nothing
        Case Else
    End Select

ExitSub:
    If Not adoRS Is Nothing Then Call CloseRecordset(adoRS, True)
    Exit Function
End Function
Public Sub ParseSQLSelect(ByVal SQLstatement As String, FieldList As String, TableList As String, WhereClause As String, OrderByClause As String)
    Dim iFrom As Integer
    Dim iWhere As Integer
    Dim iOrderBy As Integer
    
    On Error GoTo ErrorHandler
    
    FieldList = vbNullString
    TableList = vbNullString
    WhereClause = vbNullString
    OrderByClause = vbNullString
    
    'Gotta have a SELECT and FROM...
    SQLstatement = Trim(Replace(SQLstatement, vbCrLf, " "))
    If UCase(Mid(SQLstatement, 1, 6)) <> "SELECT" Then Exit Sub
    If UCase(Left(SQLstatement, Len("SELECT DISTINCT"))) = "SELECT DISTINCT" Then
        SQLstatement = Trim(Right(SQLstatement, Len(SQLstatement) - Len("SELECT DISTINCT")))
    Else
        SQLstatement = Trim(Right(SQLstatement, Len(SQLstatement) - Len("SELECT")))
    End If
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
            GoTo ExitSub
        Else
            TableList = Trim(Left(SQLstatement, iOrderBy))
            SQLstatement = Trim(Mid(SQLstatement, iOrderBy + Len(" ORDER BY ")))
            OrderByClause = Trim(SQLstatement)
            GoTo ExitSub
        End If
    Else
        TableList = Trim(Left(SQLstatement, iWhere))
        SQLstatement = Trim(Mid(SQLstatement, iWhere + Len(" WHERE ")))
        iOrderBy = InStr(UCase(SQLstatement), " ORDER BY ")
        
        If iOrderBy = 0 Then
            WhereClause = Trim(SQLstatement)
            GoTo ExitSub
        Else
            WhereClause = Trim(Left(SQLstatement, iOrderBy))
            SQLstatement = Trim(Mid(SQLstatement, iOrderBy + Len(" ORDER BY ")))
            OrderByClause = Trim(SQLstatement)
            GoTo ExitSub
        End If
    End If
    ClearError
    
ExitSub:
    Exit Sub

ErrorHandler:
    Select Case err.Number
        'Skip any errors we know we raised ourselves...
        Case Else
            saveErrNumber = err.Number
            saveErrString = err.Description
            strLastError = BuildADOerror(adoConn, longLastError)
            If longLastError = 0 Then
                longLastError = saveErrNumber
                strLastError = saveErrString
            End If
    End Select
    err.Raise longLastError, mConstants.App.EXEName & " (clsDBEngine.ParseSQLSelect)", strLastError
    Exit Sub
    Resume Next 'for debugging purposes...
End Sub
Public Property Get ServerName() As String
    ServerName = strServerName
End Property
Public Property Let ServerName(Value As String)
    strServerName = Value
End Property
Public Sub SQLApplyAsClause(ByRef SQLsource As String)
    Dim i As Integer
    Dim iOnStart As Long
    Dim iOnEnd As Long
    Dim FieldList As String
    Dim TableList As String
    Dim SaveTableList As String
    Dim WhereClause As String
    Dim OrderByClause As String
    Dim Token As String
    Dim rsColumns As ADODB.Recordset
    Dim TableName As String
    Dim TBNAME As String
    Dim strOwner As String
    Dim whereAS As Integer
            
    Call Trace(trcEnter, "SQLApplyAsClause", trcApplicationDetail)
    Call initDBEngineParms(DBEngine)
    Call ParseSQLSelect(SQLsource, FieldList, TableList, WhereClause, OrderByClause)
    SaveTableList = TableList       'TableList will be chopped-up if ANSI SQL-92 Join syntax is found, so save it here to rebuild SQL statement later...
    If UCase(Left(SQLsource, Len("SELECT DISTINCT"))) = "SELECT DISTINCT" Then
        SQLsource = "Select Distinct"
    Else
        SQLsource = "Select"
    End If
    While Len(FieldList) > 0
        If UCase(Left(FieldList, 8)) = "CONVERT(" Then
            'CONVERT takes two arguments, so find the correct comma position...
            i = InStr(9, FieldList, ",")        's/b the first one...
            i = InStr(i + 1, FieldList, ",")    's/b the one we want...
        Else
            i = InStr(FieldList, ",")
        End If
        If i > 0 Then
            Token = Trim(Mid(FieldList, 1, i - 1))
            FieldList = Trim(Mid(FieldList, i))
        Else
            Token = Trim(FieldList)
            FieldList = vbNullString
        End If
        
        If Token = "*" Then
            SQLsource = SQLsource & " " & Token
        ElseIf InStr(Token, ".*") > 0 Then
            'Need to handle Table Aliases here...
            TableName = Trim(Left(Token, InStr(Token, ".*") - 1))
            TBNAME = TableName
            
            'Before getting started, double check for ANSI SQL-92 Join syntax (i.e. "LEFT OUTER JOIN" stuff)...
            If InStr(UCase(TableList), " JOIN ") > 0 Then
                'Debug.Assert "Found a ANSI SQL-92 Join Syntax..."
                TableList = Replace(UCase(TableList), " LEFT OUTER JOIN ", ",")
                TableList = Replace(UCase(TableList), " RIGHT OUTER JOIN ", ",")
                TableList = Replace(UCase(TableList), " FULL OUTER JOIN ", ",")
                TableList = Replace(UCase(TableList), " MERGE JOIN ", ",")
                TableList = Replace(UCase(TableList), " INNER JOIN ", ",")
                'Now remove the "ON expression1 operator expression2" portion of the syntax...
                iOnStart = InStr(UCase(TableList), " ON ")
                While iOnStart > 0
                    'Need to find the end of the "ON" clause...
                    iOnEnd = InStrMX(1, TableList, "<>=")               'This should get us past the first expression...
                    iOnEnd = InStrMX(iOnEnd, TableList, " ")            'This should get us the value we want...
                    TableList = Mid(TableList, 1, iOnStart - 1) & Trim(Mid(TableList, iOnEnd + 1))
                    iOnStart = InStr(UCase(TableList), " ON ")
                Wend
            End If
            
            'OK, Now TableList should be in the ole format we like to work with... (i.e. "Table1, Table2" or "Table1 T1, Table2 T2"...)
            For i = 1 To TokenCount(TableList, ",")
                Token = Trim(ParseStr(TableList, i, ",", """"))
                If Token = TableName Then
                    TBNAME = TableName
                    Exit For
                ElseIf TokenCount(Token, " ") = 2 Then
                    If Trim(ParseStr(Token, 2, " ")) = TableName Then
                        TBNAME = Trim(ParseStr(Token, 1, " "))
                        Exit For
                    End If
                End If
            Next
            strOwner = dbparmOwner
            If InStr(TBNAME, ".") Then
                strOwner = ParseStr(TBNAME, 1, ".")
                If strOwner <> dbparmOwner And strOwner = UCase(dbparmOwner) Then strOwner = dbparmOwner
                TBNAME = ParseStr(TBNAME, 2, ".")
            End If
            
            Set rsColumns = New ADODB.Recordset
            'Note that this is an adOpenForwardOnly cursor for performance reasons
            Call OpenRecordset(rsColumns, "Select NAME from XSYSCOLUMNS Where TBNAME='" & TBNAME & "' And TBOWNER='" & strOwner & "' Order By COLNO", adOpenForwardOnly, adLockReadOnly)
            While Not rsColumns.EOF
                If rsColumns("NAME") <> "ROWID" Then SQLsource = SQLsource & " " & TableName & "." & rsColumns("NAME") & " As """ & TableName & "." & rsColumns("NAME") & ""","
                rsColumns.MoveNext
            Wend
            Call CloseRecordset(rsColumns, True)
            SQLsource = Left(SQLsource, Len(SQLsource) - 1)     'Remove the final comma...
        ElseIf InStr(UCase(Token), " AS ") <> 0 Then
            whereAS = InStr(UCase(Token), " AS ")
            If Mid(Token, whereAS + 4, 1) = """" Or Mid(Token, whereAS + 4, 1) = "'" Then
                SQLsource = SQLsource & " " & Token
            Else
                SQLsource = SQLsource & " " & Left(Token, whereAS - 1) & " As """ & Mid(Token, whereAS + 4) & """"
            End If
'Don't treat quoted constants any different from other fields...
'        ElseIf Left(Token, 1) = "'" Then                        'Assume a quoted constant...
'            SQLsource = SQLsource & " " & Token & " As """ & Mid(Token, 2, Len(Token) - 2) & """"
        Else
            SQLsource = SQLsource & " " & Token & " As """ & Token & """"
        End If
        
        If Left(FieldList, 1) = "," Then
            SQLsource = SQLsource & ","
            FieldList = Trim(Mid(FieldList, 2))
        End If
        
        'Debug.Print "Token #" & i & ": " & Token
    Wend
    SQLsource = SQLsource & " From " & SaveTableList
    If WhereClause <> vbNullString Then SQLsource = SQLsource & " Where " & WhereClause
    If OrderByClause <> vbNullString Then SQLsource = SQLsource & " Order By " & OrderByClause
    
ExitSub:
    Call CloseRecordset(rsColumns, True)
    Call Trace(trcExit, "SQLApplyAsClause", trcApplicationDetail)
    Exit Sub

ErrorHandler:
    Select Case err.Number
        'Skip any errors we know we raised ourselves...
        Case Else
            saveErrNumber = err.Number
            saveErrString = err.Description
            strLastError = BuildADOerror(adoConn, longLastError)
            If longLastError = 0 Then
                longLastError = saveErrNumber
                strLastError = saveErrString
            End If
    End Select
    Call Trace(trcExit, "SQLApplyAsClause() - ErrorHandler", trcApplicationDetail)
    err.Raise longLastError, mConstants.App.EXEName & " (clsDBEngine.SQLApplyAsClause)", strLastError
    Exit Sub
    Resume Next 'for debugging purposes...
End Sub
Public Sub SQLexecute(Source As String)
    Dim adoError As Long
    Dim RecordsAffected As Long
    Dim Operation As String
    Dim TableName As String
    Dim iPos As Integer
    Dim PossibleDups As String
    
    On Error GoTo ErrorHandler
    Call Trace(trcEnter, "SQLexecute()", trcApplicationDetail)
    
    Operation = UCase(Left(Trim(Source), 6))
    Select Case Operation
        Case "UPDATE"           'UPDATE <TableName> SET ...
            TableName = ParseStr(Source, 2, " ")
        Case "INSERT"           'INSERT INTO <TableName> ...
            TableName = ParseStr(Source, 3, " ")
        Case "DELETE"           'DELETE FROM <TableName> ...
            TableName = ParseStr(Source, 3, " ")
        Case Else
            err.Raise dberrDBEUnsupportedSQL, mConstants.App.EXEName & " (clsDBEngine.SQLexecute)", Replace(mConstants.dberrDBEUnsupportedSQLStr, "%1", Source)
    End Select
    
    If longActiveTXLevel <= 0 Then
        err.Raise dberrDBENoActiveTXN, mConstants.App.EXEName & " (clsDBEngine.SQLexecute)", Replace(mConstants.dberrDBENoActiveTXNStr, "%1", Operation)
    End If
    
    If fDebugTXN Then MsgBox Format(Now, "General Date") & " About to .Execute " & vbCr & Source, vbInformation
    Call Trace(trcBody, "SQLexecute() - .Execute """ & Source & """, RecordsAffected, adCmdText", trcSQL)
    adoConn.Execute Source, RecordsAffected, adCmdText
    Call Trace(trcBody, "SQLexecute() - Execute Complete; " & RecordsAffected & " Records Affected...", trcSQL)
    If fDebugTXN Then MsgBox Format(Now, "General Date") & " .Execute complete...", vbInformation
    If InconsistentViewDetected(Source, RecordsAffected) Then
        Call Trace(trcBody, "SQLexecute() - Inconsistent View Detected, raising error..." & vbCrLf & vbCrLf & _
            "SQL Statement: " & vbCrLf & Source, trcSQL)
        err.Raise dberrInconsistentView, "FiRReApplication.SQLexecute", "Someone else has modified this record." & vbCr & _
                    "Please re-start your transaction to view the latest record values." & _
                    vbCr & vbCr & "(" & TableName & ")" & vbCrLf & vbCrLf & _
                "SQL Source: " & vbCrLf & Source
        GoTo ExitSub
    End If
    ClearError
    
ExitSub:
    Call Trace(trcExit, "SQLexecute()", trcApplicationDetail)
    Exit Sub

ErrorHandler:
    Select Case err.Number
        'Skip any errors we know we raised ourselves...
        Case dberrInconsistentView
        Case dberrDBEUnsupportedSQL
        Case Else
            saveErrNumber = err.Number
            saveErrString = err.Description
            strLastError = BuildADOerror(adoConn, longLastError)
            If longLastError = 0 Then
                longLastError = saveErrNumber
                strLastError = saveErrString
            Else
                Select Case longLastError
                    Case dberrCannotDeleteRow
                        Call Trace(trcBody, "SQLexecute() - Cannot Delete Row..." & vbCrLf & vbCrLf & _
                            "SQL Statement: " & vbCrLf & Source, trcSQL)
                        strLastError = strLastError & vbCrLf & _
                            TableName & " is attached to another record." & vbCr & _
                            "You must first change all of the records it is attached to, " & vbCr & _
                            "before you can delete it."
                    Case dberrCannotUpdateRow
                        Call Trace(trcBody, "SQLexecute() - Cannot Update Row..." & vbCrLf & vbCrLf & _
                            "SQL Statement: " & vbCrLf & Source, trcSQL)
                        strLastError = strLastError & vbCrLf & _
                            "This " & TableName & " is used as a link to another record." & vbCr & _
                            "It cannot be changed."
                    Case dberrDuplicateKey
                        Call Trace(trcBody, "SQLexecute() - Duplicate Key Detected..." & vbCrLf & vbCrLf & _
                            "SQL Statement: " & vbCrLf & Source, trcSQL)
                        Select Case DBEngine
                            Case dbeSQLServer
                                'The error from SQL Server should identify the index causing the problem...
                                'No need to elaborate any further here...
                                strLastError = strLastError & vbCrLf & _
                                    "Duplicate key exists." & vbCr & _
                                    "Please choose a different value for this " & TableName & "." & vbCr & vbCr & _
                                    Mid(strLastError, 1, InStr(strLastError, vbCr))
                            Case dbeSQLBase
                                strLastError = strLastError & vbCrLf & _
                                    "Duplicate key exists." & vbCr & _
                                    "Please choose a different value for this " & TableName & "."
                            Case Else
                                strLastError = strLastError & vbCrLf & _
                                    "Duplicate key exists." & vbCr & _
                                    "Please choose a different value for this " & TableName & "."
                        End Select
                    Case dberrDuplicatePrimaryKey
                        Call Trace(trcBody, "SQLexecute() - Duplicate Primary Key Detected..." & vbCrLf & vbCrLf & _
                            "SQL Statement: " & vbCrLf & Source, trcSQL)
                            'The error from SQL Server should identify the index causing the problem...
                            'No need to elaborate any further here...
                            strLastError = strLastError & vbCrLf & _
                                "Duplicate key exists on primary key field." & vbCr & _
                                "Please choose a different value for this " & TableName & "." & vbCr & vbCr & _
                                Mid(strLastError, 1, InStr(strLastError, vbCr))
                        'Mask this error as a normal dberrDuplicateKey error so we can react to them both the same way...
                        longLastError = dberrDuplicateKey
                    Case dberrInconsistentView
                        Call Trace(trcBody, "SQLexecute() - Inconsistent View Detected..." & vbCrLf & vbCrLf & _
                            "SQL Statement: " & vbCrLf & Source, trcSQL)
                        strLastError = strLastError & vbCrLf & _
                            "Someone else has modified this record." & vbCr & _
                            "Please re-start your transaction to view the latest record values." & _
                            vbCr & vbCr & "(" & TableName & ")" & vbCrLf & vbCrLf & _
                        "SQL Source: " & vbCrLf & Source
                    Case dberrLockTimeOut
                        Call Trace(trcBody, "SQLexecute() - Lock TimeOut Detected..." & vbCrLf & vbCrLf & _
                            "SQL Statement: " & vbCrLf & Source, trcSQL)
                        strLastError = strLastError & vbCrLf & _
                            "Someone else has one of the records you require locked, and you have timed-out waiting..." & vbCr & _
                            "Please try again later." & _
                            vbCr & vbCr & "(" & TableName & ")" & vbCrLf & vbCrLf & _
                        "SQL Source: " & vbCrLf & Source
                    Case dberrDeadLock
                        Call Trace(trcBody, "SQLexecute() - Deadlock Detected..." & vbCrLf & vbCrLf & _
                            "SQL Statement: " & vbCrLf & Source, trcSQL)
                        'Since the deadlock error causes an implicit rollback, update our indicator...
                        longActiveTXLevel = longActiveTXLevel - 1
                        strLastError = strLastError & vbCrLf & _
                            "Deadlock error." & vbCrLf & vbCrLf & _
                            "SQL Source: " & vbCrLf & Source
                    Case Else
                        strLastError = strLastError & vbCrLf & _
                            "SQL Source: " & vbCrLf & Source
                        Call Trace(trcBody, "SQLexecute() - " & strLastError, trcSQL)
                End Select
            End If
    End Select
    Call Trace(trcExit, "SQLexecute() - ErrorHandler", trcApplicationDetail)
    err.Raise longLastError, mConstants.App.EXEName & " (clsDBEngine.SQLexecute)", strLastError
    Exit Sub
    Resume Next 'for debugging purposes...
End Sub
Public Property Get TraceObject() As SIASUTL.clsTrace
    Set TraceObject = objTrace
End Property
Public Property Set TraceObject(NewValue As SIASUTL.clsTrace)
    Set objTrace = NewValue
End Property
Public Property Get UserID() As String
    UserID = strUserID
End Property
Public Property Let UserID(Value As String)
    strUserID = Value
End Property
Private Sub VerifyDSN(DBEngine As SIASDB.dbeDBEngineEnum, ODBCDriver As String, ServerName As String, DatabaseName As String)
    Dim defaultRegistryDriver As String
    Dim RegistryDriver As String
    Dim RegistryODBCDriver As String
    Dim RegistryDB As String
    Dim RegistryServer As String
    Dim ServerNameKey As String
    Dim RootKey As Long
    Dim strRootKey As String
    Dim KeyPath As String
    Dim LocalODBCDriver As String
    Dim ErrString As String
    Dim saveTraceMode As Boolean
    
'Note: libRegistry routines do their own error handling (i.e. call MsgBox themselves on errors)...
    
    Call Trace(trcEnter, "VerifyDSN()", trcApplicationDetail)
    saveTraceMode = TraceMode
    Select Case DBEngine
        Case dbeAccess
            ServerNameKey = "DBQ"
            strDBUpperCaseFunctionName = "UCASE"
            strDBCurrentTimestamp = "CURRENT_TIMESTAMP"
            LocalODBCDriver = ODBCDriver
            defaultRegistryDriver = SystemDir & "\odbcjt32.dll"
            Call initDBEngineParms(DBEngine)
        Case dbeSQLBase
            ServerNameKey = "ServerName"
            strDBUpperCaseFunctionName = "@upper"
            strDBCurrentTimestamp = "SYSDATE"
            
            'Apparently, SCR version 7 (latest P2SODBC.DLL) SCR no longer supports the ODBC driver
            'distributed by Centura but does support its own driver "CR SQLBase", so we us it here...
            LocalODBCDriver = "CR SQLBase"
            defaultRegistryDriver = SystemDir & "\Crgup14.dll"
            'LocalODBCDriver = "Centura SQLBase 3.5 32-bit Driver -NT & Win95"
            'defaultRegistryDriver = s.SystemDir & "\C2gup14.dll"
            Call initDBEngineParms(DBEngine)
        Case dbeSQLServer
            ServerNameKey = "Server"
            strDBUpperCaseFunctionName = "UPPER"
            strDBCurrentTimestamp = "CURRENT_TIMESTAMP"
            LocalODBCDriver = ODBCDriver
            defaultRegistryDriver = SystemDir & "\sqlsrv32.dll"
            Call initDBEngineParms(DBEngine)
    End Select
    
    'First, get the filename of the driver we're using... Since we're already connected to get to this point,
    'we have to assume the appropriate ODBCDriver is installed. If we don't get a value back, blow up.
    RegistryDriver = GetRegistryKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBCINST.INI\" & LocalODBCDriver, "Driver")
    If RegistryDriver = vbNullString Then
        'ErrString = "Cannot identify ODBC Driver in the system registry." & vbCrLf & _
            vbTab & "Registry key not found (or not set):" & vbCrLf & _
            vbTab & "[HKEY_LOCAL_MACHINE\SOFTWARE\ODBC\ODBCINST.INI\" & LocalODBCDriver & "]" & vbCrLf & _
            vbTab & "Driver" & vbTab & """" & """"
        'Call MsgBox(ErrString, vbCritical, App.EXEName)
        'Call MsgBox(ErrString, vbExclamation, App.EXEName)
        'Call Trace(trcBody, ErrString, trcApplicationDetail)
        'Call Trace(trcExit, "VerifyDSN()", trcApplicationDetail)
        'End
        RegistryDriver = defaultRegistryDriver  'Fault tolerance...
        ErrString = "Cannot identify ODBC Driver in the system registry." & vbCrLf & _
            vbTab & "Registry key not found (or not set):" & vbCrLf & _
            vbTab & "[HKEY_LOCAL_MACHINE\SOFTWARE\ODBC\ODBCINST.INI\" & LocalODBCDriver & "]" & vbCrLf & _
            vbTab & "Driver" & vbTab & """" & """" & vbCrLf & vbCrLf & _
            "This may affect " & App.EXEName & "'s ability to run reports."
        Call MsgBox(ErrString, vbExclamation, App.EXEName)
        TraceMode = True
        Call Trace(trcBody, ErrString, trcApplicationDetail)
    Else
        ErrString = "Verified ODBC Driver in the system registry..." & vbCrLf & _
            vbTab & "Registry Key:" & vbCrLf & _
            vbTab & "[HKEY_LOCAL_MACHINE\SOFTWARE\ODBC\ODBCINST.INI\" & LocalODBCDriver & "]" & vbCrLf & _
            vbTab & "Driver" & vbTab & """" & RegistryDriver & """"
        Call Trace(trcBody, ErrString, trcApplicationDetail)
    End If
    
    'Next, see if there's a System DSN named App.EXEName...
    RootKey = HKEY_LOCAL_MACHINE
    strRootKey = "HKEY_LOCAL_MACHINE"
    KeyPath = "SOFTWARE\ODBC\ODBC.INI"
    RegistryODBCDriver = GetRegistryKeyValue(RootKey, KeyPath & "\ODBC Data Sources", mConstants.App.EXEName)
    If RegistryODBCDriver = vbNullString Then
        'OK, no System DSN, check for a User DSN named mConstants.App.EXEName...
        RootKey = HKEY_CURRENT_USER
        strRootKey = "HKEY_CURRENT_USER"
        KeyPath = "Software\ODBC\ODBC.INI"
        RegistryODBCDriver = GetRegistryKeyValue(RootKey, KeyPath & "\ODBC Data Sources", mConstants.App.EXEName)
        
        If RegistryODBCDriver = vbNullString Then
            'OK, there's no User DSN either... I guess we need to add one (God, I hate Crystal Reports for making us do this)...
            'Now create the mConstants.App.EXEName key...
            TraceMode = True
            Call Trace(trcBody, mConstants.App.EXEName & " DSN not found... Creating DSN..." & vbCrLf & _
                vbTab & "Registry Key:" & vbCrLf & _
                vbTab & "[" & strRootKey & "\" & KeyPath & "]" & vbCrLf & _
                vbTab & "ODBC Driver" & vbTab & """" & LocalODBCDriver & """" & vbCrLf & _
                vbTab & ServerNameKey & vbTab & """" & ServerName & """" & vbCrLf & _
                vbTab & "Database" & vbTab & """" & DatabaseName & """", trcApplicationDetail)
            'CreateRegistryKey and SetRegistryKeyValue do their own error handling (i.e. call MsgBox themselves on errors)...
            Call CreateRegistryDSN(RootKey, KeyPath, ServerNameKey, LocalODBCDriver, RegistryDriver, ServerName, DatabaseName)
            GoTo ExitSub
        Else
            Call Trace(trcBody, "Verified " & mConstants.App.EXEName & " User DSN ODBC Driver..." & vbCrLf & _
                vbTab & "Registry Key:" & vbCrLf & _
                vbTab & "[" & strRootKey & "\" & KeyPath & "\ODBC Data Sources]" & vbCrLf & _
                vbTab & mConstants.App.EXEName & vbTab & """" & RegistryODBCDriver & """", trcApplicationDetail)
        End If
    Else
        Call Trace(trcBody, "Verified " & mConstants.App.EXEName & " System DSN ODBC Driver..." & vbCrLf & _
            "Registry Key:" & vbCrLf & _
            "[" & strRootKey & "\" & KeyPath & "\ODBC Data Sources]" & vbCrLf & _
            mConstants.App.EXEName & vbTab & """" & RegistryODBCDriver & """", trcApplicationDetail)
    End If
            
    'OK, we found a mConstants.App.EXEName DSN (System or User)... Double check to make sure the ODBC Driver, Server and Database have not changed...
    If RegistryODBCDriver <> LocalODBCDriver Then
        'Switched SQLBase to/from SQL Server; Delete the existing mConstants.App.EXEName key, and create a new one for the
        'current driver...
        TraceMode = True
        Call Trace(trcBody, mConstants.App.EXEName & " DSN ODBC Driver (""" & RegistryDriver & """) does not match that of connected database... Replacing existing DSN..." & vbCrLf & _
            vbTab & "Registry Key:" & vbCrLf & _
            vbTab & "[" & strRootKey & "\" & KeyPath & "]" & vbCrLf & _
            vbTab & "ODBC Driver" & vbTab & """" & LocalODBCDriver & """" & vbCrLf & _
            vbTab & ServerNameKey & vbTab & """" & ServerName & """" & vbCrLf & _
            vbTab & "Database" & vbTab & """" & DatabaseName & """", trcApplicationDetail)
        Call DeleteRegistryKey(RootKey, KeyPath & "\" & mConstants.App.EXEName, "")
        Call DeleteRegistryKeyValue(RootKey, KeyPath & "\ODBC Data Sources", mConstants.App.EXEName)
        Call CreateRegistryDSN(RootKey, KeyPath, ServerNameKey, LocalODBCDriver, RegistryDriver, ServerName, DatabaseName)
        GoTo ExitSub
    End If
    
    RegistryServer = GetRegistryKeyValue(RootKey, KeyPath & "\" & mConstants.App.EXEName, ServerNameKey)
    If RegistryServer <> ServerName Then
        'Update the Key value with the current server name...
        TraceMode = True
        Call Trace(trcBody, mConstants.App.EXEName & " DSN Database Server (""" & RegistryServer & """) does not match that of connected database... Updating DSN..." & vbCrLf & _
            vbTab & "Registry Key:" & vbCrLf & _
            vbTab & "[" & strRootKey & "\" & KeyPath & "\" & mConstants.App.EXEName & "]" & vbCrLf & _
            vbTab & ServerNameKey & vbTab & """" & ServerName & """", trcApplicationDetail)
        Call SetRegistryKeyValue(RootKey, KeyPath & "\" & mConstants.App.EXEName, ServerNameKey, REG_SZ, ServerName)
    End If
    RegistryDB = GetRegistryKeyValue(RootKey, KeyPath & "\" & mConstants.App.EXEName, "Database")
    If RegistryDB <> DatabaseName Then
        'Update the Key value with the current database name...
        TraceMode = True
        Call Trace(trcBody, mConstants.App.EXEName & " DSN Database Name (""" & RegistryDB & """) does not match that of connected database... Updating DSN..." & vbCrLf & _
            vbTab & "Registry Key:" & vbCrLf & _
            vbTab & "[" & strRootKey & "\" & KeyPath & "\" & mConstants.App.EXEName & "]" & vbCrLf & _
            vbTab & "Database" & vbTab & """" & DatabaseName & """", trcApplicationDetail)
        Call SetRegistryKeyValue(RootKey, KeyPath & "\" & mConstants.App.EXEName, "Database", REG_SZ, DatabaseName)
    End If

ExitSub:
    TraceMode = saveTraceMode
    Call Trace(trcExit, "VerifyDSN()", trcApplicationDetail)
End Sub
Private Sub VerifySCRoj()
    Dim RootKey As Long
    Dim KeyPath As String
    Dim Temp As String
    Dim saveTraceMode As Boolean
    
    Call Trace(trcEnter, "VerifySCRoj()", trcApplicationDetail)
    saveTraceMode = TraceMode
    RootKey = HKEY_CURRENT_USER
    KeyPath = "Software\Seagate Software\Crystal Reports\DatabaseOptions\OuterJoin"
    Call CreateRegistryKey(RootKey, KeyPath)       'Duplicate errors are ignored...
    Select Case DBEngine
        Case dbeSQLBase
            Temp = GetRegistryKeyValue(RootKey, KeyPath, "PlusEqual")
            If Temp = vbNullString Then
                TraceMode = True
                Call Trace(trcBody, "Updating Outer-Join syntax in the system registry..." & vbCrLf & _
                    vbTab & "Registry Key:" & vbCrLf & _
                    vbTab & "[HKEY_CURRENT_USER\" & KeyPath & "]" & vbCrLf & _
                    vbTab & "PlusEqual" & vbTab & """crgup11, C2GUP11, crgup12, C2GUP12, crgup13, C2GUP13, crgup14, C2GUP14""", trcApplicationDetail)
                Call SetRegistryKeyValue(RootKey, KeyPath, "PlusEqual", REG_SZ, "crgup11, C2GUP11, crgup12, C2GUP12, crgup13, C2GUP13, crgup14, C2GUP14")
            Else
                If InStr(LCase(Temp), "crgup14, c2gup14") = 0 Then
                    TraceMode = True
                    Call Trace(trcBody, "Updating Outer-Join syntax in the system registry..." & vbCrLf & _
                        vbTab & "Registry Key:" & vbCrLf & _
                        vbTab & "[HKEY_CURRENT_USER\" & KeyPath & "]" & vbCrLf & _
                        vbTab & "PlusEqual" & vbTab & """" & Temp & ", crgup11, C2GUP11, crgup12, C2GUP12, crgup13, C2GUP13, crgup14, C2GUP14""", trcApplicationDetail)
                    Call SetRegistryKeyValue(RootKey, KeyPath, "PlusEqual", REG_SZ, Temp & ", crgup11, C2GUP11, crgup12, C2GUP12, crgup13, C2GUP13, crgup14, C2GUP14")
                Else
                    Call Trace(trcBody, "Verified Outer-Join syntax in the system registry..." & vbCrLf & _
                        vbTab & "Registry Key:" & vbCrLf & _
                        vbTab & "[HKEY_CURRENT_USER\" & KeyPath & "]" & vbCrLf & _
                        vbTab & "PlusEqual" & vbTab & """" & Temp & """", trcApplicationDetail)
                End If
            End If
        Case dbeSQLServer
            Temp = GetRegistryKeyValue(RootKey, KeyPath, "StarEqual")
            If Temp = vbNullString Then
                TraceMode = True
                Call Trace(trcBody, "Updating Outer-Join syntax in the system registry..." & vbCrLf & _
                    vbTab & "Registry Key:" & vbCrLf & _
                    vbTab & "[HKEY_CURRENT_USER\" & KeyPath & "]" & vbCrLf & _
                    vbTab & "StarEqual" & vbTab & """sqlsrv32, sqlsrvr""", trcApplicationDetail)
                Call SetRegistryKeyValue(RootKey, KeyPath, "StarEqual", REG_SZ, "sqlsrv32, sqlsrvr")
            Else
                If InStr(LCase(Temp), "sqlsrv32") = 0 Then
                    TraceMode = True
                    Call Trace(trcBody, "Updating Outer-Join syntax in the system registry..." & vbCrLf & _
                        vbTab & "Registry Key:" & vbCrLf & _
                        vbTab & "[HKEY_CURRENT_USER\" & KeyPath & "]" & vbCrLf & _
                        vbTab & "StarEqual" & vbTab & """" & Temp & ", sqlsrv32, sqlsrvr""", trcApplicationDetail)
                    Call SetRegistryKeyValue(RootKey, KeyPath, "StarEqual", REG_SZ, Temp & ", sqlsrv32, sqlsrvr")
                Else
                    Call Trace(trcBody, "Verified Outer-Join syntax in the system registry..." & vbCrLf & _
                        vbTab & "Registry Key:" & vbCrLf & _
                        vbTab & "[HKEY_CURRENT_USER\" & KeyPath & "]" & vbCrLf & _
                        vbTab & "StarEqual" & vbTab & """" & Temp & """", trcApplicationDetail)
                End If
            End If
        Case Else
    End Select

ExitSub:
    TraceMode = saveTraceMode
    Call Trace(trcExit, "VerifySCRoj()", trcApplicationDetail)
End Sub


