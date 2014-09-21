VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataBaseInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
Private mlClassDebugID As Long  'local variable to hold the serialized class ID that was created in Class_Initialize
Public Key As String
Private mvarRecordset As ADODB.Recordset
Public Property Set Recordset(vData As ADODB.Recordset)
    On Error GoTo SetErr

    Call Trace(trcEnter, "Class: " & TypeName(Me) & " instance " & mlClassDebugID & " - Property Set Recordset()")
    Set mvarRecordset = vData
    Call Trace(trcExit, "Class: " & TypeName(Me) & " instance " & mlClassDebugID & " - Property Set Recordset()")
    Exit Property

SetErr:
    Call RaiseError(MyUnhandledError, "DataBaseInfo:Recordset Property Set")
End Property
Public Property Get Recordset() As ADODB.Recordset
    Call Trace(trcEnter, "Class: " & TypeName(Me) & " instance " & mlClassDebugID & " - Property Get Recordset()")
    On Error GoTo GetErr

    Set Recordset = mvarRecordset
    Call Trace(trcExit, "Class: " & TypeName(Me) & " instance " & mlClassDebugID & " - Property Get Recordset")
    Exit Property

GetErr:
    Call RaiseError(MyUnhandledError, "DataBaseInfo:Recordset Property Get")
End Property
Private Sub Class_Initialize()
    mlClassDebugID = GetNextClassDebugID()  'get the next available class ID, and print out that the class was created successfully
    Call Trace(trcEnter, "Class: " & TypeName(Me) & " instance " & mlClassDebugID & " - Initialize()")
    Call Trace(trcExit, "Class: " & TypeName(Me) & " instance " & mlClassDebugID & " - Initialize()")
End Sub
Private Sub Class_Terminate()
    'the class is being destroyed
    Call Trace(trcEnter, "Class: " & TypeName(Me) & " instance " & mlClassDebugID & " - Terminate()")
    Call Trace(trcExit, "Class: " & TypeName(Me) & " instance " & mlClassDebugID & " - Terminate()")
End Sub
Public Property Get ClassDebugID()
    ClassDebugID = mlClassDebugID
End Property

