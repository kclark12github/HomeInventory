VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSQL 
   Caption         =   "SQL"
   ClientHeight    =   5256
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   7596
   LinkTopic       =   "Form1"
   ScaleHeight     =   5256
   ScaleWidth      =   7596
   StartUpPosition =   1  'CenterOwner
   Begin ComCtl3.CoolBar cbToolBar 
      Height          =   648
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7572
      _ExtentX        =   13356
      _ExtentY        =   1143
      FixedOrder      =   -1  'True
      _CBWidth        =   7572
      _CBHeight       =   648
      _Version        =   "6.0.8169"
      Caption1        =   "Database:"
      Child1          =   "txtDatabase"
      MinHeight1      =   288
      Width1          =   1752
      Key1            =   "DB"
      NewRow1         =   0   'False
      Caption2        =   "Tables:"
      Child2          =   "dbcTables"
      MinHeight2      =   288
      Width2          =   2952
      Key2            =   "Tables"
      NewRow2         =   -1  'True
      Caption3        =   "Fields:"
      Child3          =   "cboFields"
      MinHeight3      =   288
      Width3          =   1872
      Key3            =   "Fields"
      NewRow3         =   0   'False
      Begin VB.TextBox txtDatabase 
         Height          =   288
         Left            =   876
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   24
         Width           =   6624
      End
      Begin VB.ComboBox cboFields 
         Height          =   288
         Left            =   3636
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   336
         Width           =   3864
      End
      Begin MSDataListLib.DataCombo dbcTables 
         Height          =   288
         Left            =   732
         TabIndex        =   6
         Top             =   336
         Width           =   2196
         _ExtentX        =   3874
         _ExtentY        =   508
         _Version        =   393216
         Text            =   ""
      End
   End
   Begin VB.Frame frameResults 
      Caption         =   "Results"
      Height          =   1092
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   3492
      Begin VB.TextBox txtResults 
         BeginProperty Font 
            Name            =   "r_ansi"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   852
         Left            =   60
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   3
         Top             =   180
         Width           =   3372
      End
   End
   Begin VB.Frame frameSQL 
      Caption         =   "SQL Statement"
      Height          =   1092
      Left            =   1140
      TabIndex        =   0
      Top             =   720
      Width           =   3492
      Begin VB.TextBox txtSQL 
         Height          =   852
         Left            =   60
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   3372
      End
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   4
      Top             =   5004
      Width           =   7596
      _ExtentX        =   13399
      _ExtentY        =   445
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1270
            MinWidth        =   1270
            Key             =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10478
            Key             =   "Message"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   2
            Object.Width           =   1270
            MinWidth        =   1270
            TextSave        =   "8:08 PM"
            Key             =   "Time"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const MarginTwips As Integer = 60
Const KB As Double = 1024
Const MB As Double = (1024 * KB)
Dim BufferLimit As Double
Dim InitialWidth As Double
Dim InitialHeight As Double
Dim RecordsAffected As Long
Public cnSQL As ADODB.Connection
Dim rsTables As New ADODB.Recordset
Dim rsFields As New ADODB.Recordset
Private Sub ExecuteSQL()
    Dim adoRS As ADODB.Recordset
    Dim adoError As ADODB.Error
    Dim fActiveTrans As Boolean
    Dim fld As ADODB.Field
    Dim fResponse As Boolean
    Dim ErrorCount As Long
    Dim RecordsOutput As Long
    Dim strOutput As String
    
    On Error GoTo ErrorHandler
    txtResults.Text = vbNullString
    txtResults.SetFocus
    
    fActiveTrans = False
    Select Case UCase(Mid(txtSQL.Text, 1, 6))
        Case "UPDATE", "DELETE"
            cnSQL.BeginTrans
            fActiveTrans = True
            Set adoRS = cnSQL.Execute(txtSQL.Text, RecordsAffected)
        Case "SELECT"
            Set adoRS = New ADODB.Recordset
            adoRS.Open txtSQL.Text, cnSQL, adOpenKeyset, adLockReadOnly
    End Select
    
    For Each adoError In cnSQL.Errors
        If Trim(adoError.Description) <> vbNullString Then
            txtResults.Text = txtResults.Text & adoError.Description & "(" & Hex(adoError.Number) & ")" & vbCrLf
            txtResults.Text = txtResults.Text & vbTab & "Source: " & adoError.Source & vbCrLf & _
                vbTab & "SQL State: " & adoError.SQLState & vbCrLf & _
                vbTab & "Native Error: " & adoError.NativeError & vbCrLf
            ErrorCount = ErrorCount + 1
        End If
    Next
    
    If ErrorCount > 0 Then
        If fActiveTrans Then cnSQL.RollbackTrans
        Exit Sub
    End If
    
    Select Case UCase(Mid(txtSQL.Text, 1, 6))
        Case "DELETE"
            fResponse = MsgBox(RecordsAffected & " Records deleted... Commit transaction?", vbYesNo, Me.Caption) = vbYes
        Case "UPDATE"
            fResponse = MsgBox(RecordsAffected & " Records updated... Commit transaction?", vbYesNo, Me.Caption) = vbYes
        Case "SELECT"
            'Print Column Headers...
            strOutput = vbNullString
            For Each fld In adoRS.Fields
                strOutput = strOutput & fld.Name
                Select Case fld.Type
                    Case adLongVarChar
                        strOutput = strOutput & String(80 - Len(fld.Name) + 1, " ")
                    Case adVarChar, adChar
                        If fld.DefinedSize > 80 Then
                            strOutput = strOutput & String(80 - Len(fld.Name) + 1, " ")
                        Else
                            strOutput = strOutput & String(fld.DefinedSize - Len(fld.Name) + 1, " ")
                        End If
                    Case adInteger, adCurrency
                        strOutput = strOutput & String(10 - Len(fld.Name) + 1, " ")
                    Case adDate, adDBDate, adDBTimeStamp
                        strOutput = strOutput & String(20 - Len(fld.Name) + 1, " ")
                    Case Else
                        strOutput = strOutput & " "
                End Select
            Next
            txtResults.Text = strOutput & vbCrLf
            
            'Now a column header separator line...
            strOutput = vbNullString
            For Each fld In adoRS.Fields
                Select Case fld.Type
                    Case adLongVarChar
                        strOutput = strOutput & String(80, "=") & " "
                    Case adVarChar, adChar
                        If fld.DefinedSize > 80 Then
                            strOutput = strOutput & String(80, "=") & " "
                        Else
                            strOutput = strOutput & String(fld.DefinedSize, "=") & " "
                        End If
                    Case adInteger, adCurrency
                        strOutput = strOutput & String(10, "=") & " "
                    Case adDate, adDBDate, adDBTimeStamp
                        strOutput = strOutput & String(20, "=") & " "
                    Case Else
                        strOutput = strOutput & String(Len(fld.Name), "=") & " "
                End Select
            Next
            txtResults.Text = txtResults.Text & strOutput & vbCrLf
            
            RecordsOutput = 0
            If Not (adoRS.EOF And adoRS.BOF) Then
                'Now print a row for each record...
                adoRS.MoveFirst
                While Not adoRS.EOF And Len(txtResults.Text) < BufferLimit
                    strOutput = vbNullString
                    For Each fld In adoRS.Fields
                        If IsNull(fld.Value) Then
                            Select Case fld.Type
                                Case adVarChar, adChar
                                    If Len(fld.Name) > fld.DefinedSize Then
                                        strOutput = strOutput & "<Null>" & String(Len(fld.Name) - Len("<Null>") + 1, " ")
                                    ElseIf fld.DefinedSize > 80 Then
                                        strOutput = strOutput & "<Null>" & String(80 - Len("<Null>") + 1, " ")
                                    Else
                                        strOutput = strOutput & "<Null>" & String(fld.DefinedSize - Len("<Null>") + 1, " ")
                                    End If
                                Case adCurrency
                                    If Len(fld.Name) > Len("<Null>") Then
                                        strOutput = strOutput & "<Null>" & String(Len(fld.Name) - Len("<Null>") + 1, " ")
                                    Else
                                        strOutput = strOutput & "<Null>" & String(10 - Len("<Null>") + 1, " ")
                                    End If
                                Case adInteger
                                    If Len(fld.Name) > 10 Then
                                        strOutput = strOutput & "<Null>" & String(Len(fld.Name) - Len("<Null>") + 1, " ")
                                    Else
                                        strOutput = strOutput & "<Null>" & String(10 - Len("<Null>") + 1, " ")
                                    End If
                                Case adDate, adDBDate, adDBTimeStamp
                                    If Len(fld.Name) > 20 Then
                                        strOutput = strOutput & "<Null>" & String(Len(fld.Name) - Len("<Null>") + 1, " ")
                                    Else
                                        strOutput = strOutput & "<Null>" & String(20 - Len("<Null>") + 1, " ")
                                    End If
                                Case adBoolean
                                    If Len(fld.Name) > Len("<Null>") Then
                                        strOutput = strOutput & "<Null>" & String(Len(fld.Name) - Len("<Null>") + 1, " ")
                                    Else
                                        strOutput = strOutput & "<Null>" & String(Len("false") - Len("<Null>") + 1, " ")
                                    End If
                                Case Else
                                    If Len(fld.Name) > Len("<Null>") Then
                                        strOutput = strOutput & "<Null>" & String(Len(fld.Name) - Len("<Null>") + 1, " ")
                                    ElseIf fld.ActualSize > 80 Then
                                        strOutput = strOutput & "<Null>" & String(80 - Len("<Null>") + 1, " ")
                                    ElseIf fld.DefinedSize > 80 Then
                                        strOutput = strOutput & "<Null>" & String(80 - Len("<Null>") + 1, " ")
                                    Else
                                        strOutput = strOutput & "<Null>" & String(fld.DefinedSize - Len("<Null>") + 1, " ")
                                    End If
                            End Select
                        Else
                            'If fld.DefinedSize > 80 Then Debug.Print fld.Name & " (" & fld.DefinedSize & ")"
                            Select Case fld.Type
                                Case adVarChar, adChar
                                    If Len(fld.Name) > fld.DefinedSize Then
                                        strOutput = strOutput & fld.Value & String(Len(fld.Name) - Len(fld.Value) + 1, " ")
                                    ElseIf fld.DefinedSize > 80 Then
                                        strOutput = strOutput & Mid(fld.Value, 1, 80) & String(80 - Len(fld.Value) + 1, " ")
                                    Else
                                        strOutput = strOutput & fld.Value & String(fld.DefinedSize - Len(fld.Value) + 1, " ")
                                    End If
                                Case adCurrency
                                    If Len(fld.Name) > Len(Format(fld.Value, "Currency")) Then
                                        strOutput = strOutput & Format(fld.Value, "Currency") & String(Len(fld.Name) - Len(Format(fld.Value, "Currency")) + 1, " ")
                                    Else
                                        strOutput = strOutput & Format(fld.Value, "Currency") & String(10 - Len(Format(fld.Value, "Currency")) + 1, " ")
                                    End If
                                Case adInteger
                                    If Len(fld.Name) > 10 Then
                                        strOutput = strOutput & fld.Value & String(Len(fld.Name) - Len(fld.Value) + 1, " ")
                                    Else
                                        strOutput = strOutput & fld.Value & String(10 - Len(fld.Value) + 1, " ")
                                    End If
                                Case adDate, adDBDate, adDBTimeStamp
                                    If Len(fld.Name) > 20 Then
                                        strOutput = strOutput & Format(fld.Value, "dd-MMM-yyyy hh:mm AMPM") & String(Len(fld.Name) - 20 + 1, " ")
                                    Else
                                        strOutput = strOutput & Format(fld.Value, "dd-MMM-yyyy hh:mm AMPM") & " "
                                    End If
                                Case adBoolean
                                    If Len(fld.Name) > Len("false") Then
                                        strOutput = strOutput & fld.Value & String(Len(fld.Name) - Len(fld.Value) + 1, " ")
                                    Else
                                        strOutput = strOutput & fld.Value & String(Len("false") - Len(fld.Value) + 1, " ")
                                    End If
                                Case Else
                                    If Len(fld.Name) > Len(fld.Value) Then
                                        strOutput = strOutput & fld.Value & String(Len(fld.Name) - Len(fld.Value) + 1, " ")
                                    ElseIf fld.ActualSize > 80 Then
                                        strOutput = strOutput & Mid(fld.Value, 1, 80) & " "
                                    ElseIf fld.DefinedSize > 80 Then
                                        strOutput = strOutput & fld.Value & String(80 - Len(fld.Value) + 1, " ")
                                    Else
                                        strOutput = strOutput & fld.Value & String(fld.DefinedSize - Len(fld.Value) + 1, " ")
                                    End If
                            End Select
                        End If
                    Next
                    txtResults.Text = txtResults.Text & strOutput & vbCrLf
                    RecordsOutput = RecordsOutput + 1
                    adoRS.MoveNext
                Wend
                RecordsAffected = adoRS.RecordCount
            End If
        Case Else
    End Select
    
    If fActiveTrans Then
        If fResponse Then
            cnSQL.CommitTrans
        Else
            cnSQL.RollbackTrans
        End If
    End If
    
    Select Case UCase(Mid(txtSQL.Text, 1, 6))
        Case "SELECT"
            If Len(txtResults.Text) >= BufferLimit Then
                strOutput = RecordsOutput & " of " & RecordsAffected & " record(s)"
            Else
                strOutput = RecordsOutput & " record(s) read"
            End If
        Case "DELETE"
            strOutput = RecordsOutput & " record(s) deleted"
        Case "UPDATE"
            strOutput = RecordsOutput & " record(s) updated"
    End Select
    sbStatus.Panels("Message").Text = strOutput
    txtSQL.SetFocus
    Exit Sub
    
ErrorHandler:
    Select Case Err.Number
        Case 7  'Out of memory
            BufferLimit = Len(txtResults.Text)
            Resume Next
        Case Else
            MsgBox Err.Description & " (" & Err.Number & ")", vbExclamation, Me.Caption
    End Select
    If fActiveTrans Then cnSQL.RollbackTrans
    Exit Sub
End Sub
Private Sub cbToolBar_HeightChanged(ByVal NewHeight As Single)
    Form_Resize
End Sub
Private Sub dbcTables_Click(Area As Integer)
    Dim SQLsource As String
    Dim fld As ADODB.Field
    
    If dbcTables.BoundText = vbNullString Then Exit Sub
    
    CloseRecordset rsFields, False
    SQLsource = "SELECT * " & "FROM [" & dbcTables.BoundText & "]"
    rsFields.MaxRecords = 1
    rsFields.Open SQLsource, cnSQL, adOpenKeyset, adLockReadOnly
    cboFields.Clear
    For Each fld In rsFields.Fields
        cboFields.AddItem fld.Name
    Next fld
    cboFields.ListIndex = 0
End Sub
Private Sub Form_Activate()
    Dim SQLsource As String
    
    sbStatus.Panels("Status").Text = "SQL"
    
    CloseRecordset rsTables, False
    SQLsource = _
        "SELECT UserTables.Name " & _
        "FROM   MSysObjects AS SysTables, MSysObjects AS UserTables " & _
        "WHERE  UserTables.ParentId=SysTables.Id AND " & _
        "       UserTables.Type=1 AND " & _
        "       UserTables.Flags=0 AND " & _
        "       SysTables.Name='Tables' " & _
        "ORDER BY UserTables.Name;"
    rsTables.Open SQLsource, cnSQL, adOpenKeyset, adLockReadOnly
    dbcTables.ListField = "Name"
    Set dbcTables.RowSource = rsTables
    dbcTables_Click dbcAreaButton
    
    BufferLimit = 50 * KB
    InitialWidth = Me.Width
    InitialHeight = Me.Height
End Sub
Private Sub Form_Resize()
    Dim NewFrameWidth As Double
    
    'NewFrameWidth = Me.ScaleWidth - cmdExecute.Width - (3 * MarginTwips)
    NewFrameWidth = Me.ScaleWidth - (2 * MarginTwips)
    
    If Me.Width < InitialWidth Or Me.Height < InitialHeight Then
        'Debug.Print "Initial Width x Height: " & InitialWidth & "x" & InitialHeight
        'Debug.Print "Me.Width x Height: " & Me.Width & "x" & Me.Height
        Me.Move Me.Left, Me.Top, InitialWidth, InitialHeight
        Exit Sub
    End If
    cbToolBar.Move 0, 0, Me.ScaleWidth, cbToolBar.Height
    'cmdExecute.Move Me.ScaleWidth - cmdExecute.Width - MarginTwips, MarginTwips
    'cmdClose.Move Me.ScaleWidth - cmdClose.Width - MarginTwips, cmdExecute.Top + cmdExecute.Height + MarginTwips
    'frameSQL.Move MarginTwips, MarginTwips, NewFrameWidth, Me.ScaleHeight / 3
    frameSQL.Move MarginTwips, MarginTwips + cbToolBar.Height, NewFrameWidth, Me.ScaleHeight / 3
    txtSQL.Move MarginTwips, 3 * MarginTwips, frameSQL.Width - (2 * MarginTwips), frameSQL.Height - (4 * MarginTwips)
    'frameResults.Move MarginTwips, frameSQL.Top + frameSQL.Height + MarginTwips, NewFrameWidth, Me.ScaleHeight - frameSQL.Height - (3 * MarginTwips) - sbStatus.Height
    frameResults.Move MarginTwips, frameSQL.Top + frameSQL.Height + MarginTwips, NewFrameWidth, Me.ScaleHeight - frameSQL.Height - (3 * MarginTwips) - sbStatus.Height - cbToolBar.Height
    txtResults.Move MarginTwips, 3 * MarginTwips, frameResults.Width - (2 * MarginTwips), frameResults.Height - (4 * MarginTwips)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    CloseRecordset rsTables, True
    CloseRecordset rsFields, True
    Set cnSQL = Nothing
End Sub
Private Sub txtSQL_GotFocus()
    While Right(txtSQL.Text, 2) = vbCrLf
        txtSQL.Text = Left(txtSQL.Text, Len(txtSQL.Text) - 2)
    Wend
    TextSelected
End Sub
Private Sub txtSQL_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Right(Trim(txtSQL.Text), 1) = ";" Then ExecuteSQL
    End If
End Sub
Private Sub txtSQL_LostFocus()
    While Right(txtSQL.Text, 2) = vbCrLf
        txtSQL.Text = Left(txtSQL.Text, Len(txtSQL.Text) - 2)
    Wend
End Sub
