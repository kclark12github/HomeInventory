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

'set this to 0 to disable debug code in this class
#Const DebugMode = 1


#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    Private mlClassDebugID As Long
#End If
Public Key As String
'local variable(s) to hold property value(s)
Private mvarDataBaseName As String 'local copy
Private mvarPathName As String 'local copy
Private mvarProvider As String 'local copy
Private mvarUserName As String 'local copy
Private mvarPassword As String 'local copy
Public Property Let Password(ByVal vData As String)
    On Error GoTo PasswordLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Password = 5
    mvarPassword = vData
    Exit Property


PasswordLetErr:


    Call RaiseError(MyUnhandledError, "DataBaseInfo:Password Property Let")
End Property


Public Property Get Password() As String
    On Error GoTo PasswordGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Password
    Set Password = mvarPassword
    Exit Property


PasswordGetErr:


    Call RaiseError(MyUnhandledError, "DataBaseInfo:Password Property Get")
End Property



Public Property Let UserName(ByVal vData As String)
    On Error GoTo UserNameLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UserName = 5
    mvarUserName = vData
    Exit Property


UserNameLetErr:


    Call RaiseError(MyUnhandledError, "DataBaseInfo:UserName Property Let")
End Property


Public Property Get UserName() As String
    On Error GoTo UserNameGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UserName
    Set UserName = mvarUserName
    Exit Property


UserNameGetErr:


    Call RaiseError(MyUnhandledError, "DataBaseInfo:UserName Property Get")
End Property



Public Property Let Provider(ByVal vData As String)
    On Error GoTo ProviderLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Provider = 5
    mvarProvider = vData
    Exit Property


ProviderLetErr:


    Call RaiseError(MyUnhandledError, "DataBaseInfo:Provider Property Let")
End Property


Public Property Get Provider() As String
    On Error GoTo ProviderGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Provider
    Set Provider = mvarProvider
    Exit Property


ProviderGetErr:


    Call RaiseError(MyUnhandledError, "DataBaseInfo:Provider Property Get")
End Property



Public Property Let PathName(ByVal vData As String)
    On Error GoTo PathNameLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PathName = 5
    mvarPathName = vData
    Exit Property


PathNameLetErr:


    Call RaiseError(MyUnhandledError, "DataBaseInfo:PathName Property Let")
End Property


Public Property Get PathName() As String
    On Error GoTo PathNameGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PathName
    PathName = mvarPathName
    Exit Property


PathNameGetErr:


    Call RaiseError(MyUnhandledError, "DataBaseInfo:PathName Property Get")
End Property



Public Property Let DataBaseName(ByVal vData As String)
    On Error GoTo DataBaseNameLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DataBaseName = 5
    mvarDataBaseName = vData
    Exit Property


DataBaseNameLetErr:


    Call RaiseError(MyUnhandledError, "DataBaseInfo:DataBaseName Property Let")
End Property


Public Property Get DataBaseName() As String
Attribute DataBaseName.VB_UserMemId = 0
    On Error GoTo DataBaseNameGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DataBaseName
    DataBaseName = mvarDataBaseName
    Exit Property


DataBaseNameGetErr:


    Call RaiseError(MyUnhandledError, "DataBaseInfo:DataBaseName Property Get")
End Property






Private Sub Class_Initialize()
    #If DebugMode Then
        'get the next available class ID, and print out
        'that the class was created successfully
        mlClassDebugID = GetNextClassDebugID()
        Debug.Print "'" & TypeName(Me) & "' instance " & mlClassDebugID & " created"
    #End If
End Sub


Private Sub Class_Terminate()
    'the class is being destroyed
    #If DebugMode Then
        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub


#If DebugMode Then
    Public Property Get ClassDebugID()
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If

