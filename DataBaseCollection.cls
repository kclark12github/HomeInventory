VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataBaseCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"DataBaseInfo"
Attribute VB_Ext_KEY = "Member0" ,"DataBaseInfo"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mlClassDebugID As Long  'local variable to hold the serialized class ID that was created in Class_Initialize
Private mCol As Collection
Public Function Add(Key As String, Recordset As ADODB.Recordset) As DataBaseInfo
    'create a new object
    Dim objNewMember As DataBaseInfo
    Set objNewMember = New DataBaseInfo

    Call Trace(trcEnter, "Class: " & TypeName(Me) & " instance " & mlClassDebugID & " - Add(""" & Key & """, Recordset)")
    'set the properties passed into the method
    objNewMember.Key = Key
    Set objNewMember.Recordset = Recordset
    If Len(Key) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, Key
    End If

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
    Call Trace(trcExit, "Class: " & TypeName(Me) & " instance " & mlClassDebugID & " - Add()")
End Function
Public Property Get Item(vntIndexKey As Variant) As DataBaseInfo
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
    Call Trace(trcEnter, "Class: " & TypeName(Me) & " instance " & mlClassDebugID & " - Property Get Item()")
    Set Item = mCol(vntIndexKey)
    Call Trace(trcExit, "Class: " & TypeName(Me) & " instance " & mlClassDebugID & " - Property Get Item()")
End Property
Public Sub Clear()
    Dim i As Integer
    Dim iCol As DataBaseInfo
    
    Call Trace(trcEnter, "Class: " & TypeName(Me) & " instance " & mlClassDebugID & " - Clear()")
    For i = 1 To mCol.Count
        Set iCol = mCol(1)
        mCol.Remove 1
        Set iCol = Nothing
    Next i
    Call Trace(trcExit, "Class: " & TypeName(Me) & " instance " & mlClassDebugID & " - Clear")
End Sub
Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Call Trace(trcEnter, "Class: " & TypeName(Me) & " instance " & mlClassDebugID & " - Property Get Count()")
    Count = mCol.Count
    Call Trace(trcExit, "Class: " & TypeName(Me) & " instance " & mlClassDebugID & " - Property Get Count()")
End Property
Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)

    Call Trace(trcEnter, "Class: " & TypeName(Me) & " instance " & mlClassDebugID & " - Remove()")
    mCol.Remove vntIndexKey
    Call Trace(trcExit, "Class: " & TypeName(Me) & " instance " & mlClassDebugID & " - Remove()")
End Sub
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Call Trace(trcEnter, "Class: " & TypeName(Me) & " instance " & mlClassDebugID & " - Get NewEnum()")
    Set NewEnum = mCol.[_NewEnum]
    Call Trace(trcExit, "Class: " & TypeName(Me) & " instance " & mlClassDebugID & " - Property Get NewEnum()")
End Property
Private Sub Class_Initialize()
    mlClassDebugID = GetNextClassDebugID()  'get the next available class ID, and print out that the class was created successfully
    Call Trace(trcEnter, "Class: " & TypeName(Me) & " instance " & mlClassDebugID & " - Initialize()")
    Set mCol = New Collection
    Call Trace(trcExit, "Class: " & TypeName(Me) & " instance " & mlClassDebugID & " - Initialize()")
End Sub
Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Call Trace(trcEnter, "Class: " & TypeName(Me) & " instance " & mlClassDebugID & " - Terminate()")
    Clear
    Set mCol = Nothing
    Call Trace(trcExit, "Class: " & TypeName(Me) & " instance " & mlClassDebugID & " - Terminate()")
End Sub
Public Property Get ClassDebugID()
    ClassDebugID = mlClassDebugID
End Property

