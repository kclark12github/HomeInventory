VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSIASDBConstants"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'clsSIASDBConstants - clsSIASDBConstants.cls
'   Constant Definitions for the DBEngine Object Library...
'   Copyright © 2002, SunGard Investor Accounting Systems
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Problem:    Programmer:     Description:
'   12/06/01    None        Ken Clark       Created;
'=================================================================================================================================
Option Explicit
Private Const BASE_ERROR_NUMBER = &HF003
Public Enum General
    'Errors logged in clsCentura...
    dberrDBENotSQLBase = vbObjectError + BASE_ERROR_NUMBER + 0
    dberrDBESQLININTF = vbObjectError + BASE_ERROR_NUMBER + 1
    dberrDBEServerPathNTF = vbObjectError + BASE_ERROR_NUMBER + 2
    dberrDBENewServerPathNTF = vbObjectError + BASE_ERROR_NUMBER + 3
    'Errors logged in clsDBEngine...
    dberrDBENoActiveTXN = vbObjectError + BASE_ERROR_NUMBER + 4
    dberrDBETooManyActiveTXNs = vbObjectError + BASE_ERROR_NUMBER + 5
    dberrDBEUnsupportedODBCDriver = vbObjectError + BASE_ERROR_NUMBER + 6
    dberrDBEUnsupportedSQL = vbObjectError + BASE_ERROR_NUMBER + 7
    dberrDBEMysteryDisconnect = -2147217887
End Enum
Public Enum dbeDBEngineEnum
    dbeUnknown = 0
    dbeAccess = 1
    dbeSQLBase = 2
    dbeSQLServer = 3
End Enum
Private myApp As VB.App
Public Property Get dberrDBENotSQLBaseStr() As String
    dberrDBENotSQLBaseStr = "DBEngine is not SQLBase."
End Property
Public Property Get dberrDBESQLININTFStr() As String
    dberrDBESQLININTFStr = "Unable to determine location of SQL.INI file."
End Property
Public Property Get dberrDBEServerPathNTFStr() As String
    dberrDBEServerPathNTFStr = "Unable to determine %1 key under [%2] in %3 file."
End Property
Public Property Get dberrDBENewServerPathNTFStr() As String
    dberrDBENewServerPathNTFStr = "Unable to locate %1 under [%2] in %3 file."
End Property
Public Property Get dberrDBENoActiveTXNStr() As String
    dberrDBENoActiveTXNStr = "%1 operation requires an active transaction."
End Property
Public Property Get dberrDBETooManyActiveTXNsStr() As String
    dberrDBETooManyActiveTXNsStr = "Transaction is already active. Multiple or nested transactions are not permitted."
End Property
Public Property Get dberrDBEUnsupportedODBCDriverStr() As String
    dberrDBEUnsupportedODBCDriverStr = "Unsupported ODBC driver specified "" %1""."
End Property
Public Property Get dberrDBEUnsupportedSQLStr() As String
    dberrDBEUnsupportedSQLStr = "Unsupported SQL statement encountered." & vbCrLf & "SQL Statement: " & vbCrLf & "%1"
End Property
Public Property Get dberrDBEMysteryDisconnectStr() As String
    dberrDBEMysteryDisconnectStr = ""
End Property
Public Property Get App() As Object
    Set App = mApp
End Property
Public Property Set App(ByVal objApp As Object)
    Set mApp = objApp
End Property
Friend Property Get mApp() As Object
    Set mApp = myApp
End Property
Friend Property Set mApp(ByVal objApp As Object)
    Set myApp = objApp
End Property
Private Sub Class_Initialize()
    Set myApp = VB.App
End Sub
Private Sub Class_Terminate()
    Set myApp = Nothing
End Sub

