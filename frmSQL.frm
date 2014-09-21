VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
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
      TabIndex        =   8
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
         TabIndex        =   0
         Top             =   24
         Width           =   6624
      End
      Begin VB.ComboBox cboFields 
         Height          =   288
         Left            =   3636
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   336
         Width           =   3864
      End
      Begin MSDataListLib.DataCombo dbcTables 
         Height          =   288
         Left            =   732
         TabIndex        =   1
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
      TabIndex        =   5
      Top             =   2160
      Width           =   3492
      Begin MSDataGridLib.DataGrid dgdList 
         Height          =   852
         Left            =   60
         TabIndex        =   9
         Top             =   180
         Width           =   3372
         _ExtentX        =   5948
         _ExtentY        =   1503
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
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
         TabIndex        =   6
         Top             =   180
         Width           =   3372
      End
   End
   Begin VB.Frame frameSQL 
      Caption         =   "SQL Statement"
      Height          =   1092
      Left            =   1140
      TabIndex        =   3
      Top             =   720
      Width           =   3492
      Begin VB.TextBox txtSQL 
         Height          =   852
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   4
         Top             =   180
         Width           =   3372
      End
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   7
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
            TextSave        =   "5:36 PM"
            Key             =   "Time"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblA 
      AutoSize        =   -1  'True
      Caption         =   "A"
      Height          =   192
      Left            =   6360
      TabIndex        =   10
      Top             =   1140
      Visible         =   0   'False
      Width           =   108
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
Dim rsList As ADODB.Recordset
Dim rsTables As New ADODB.Recordset
Dim rsFields As New ADODB.Recordset
Dim MouseY As Single
Dim MouseX As Single
Private SortDESC() As Boolean
Private Sub ExecuteSQL()
    Dim adoError As ADODB.Error
    Dim fActiveTrans As Boolean
    Dim fld As ADODB.Field
    Dim fResponse As Boolean
    Dim ErrorCount As Long
    Dim strOutput As String
    Dim eMessage As String
    Dim BooleanFormat As New StdDataFormat
    Dim CurrencyFormat As New StdDataFormat
    Dim DateFormat As New StdDataFormat
    Dim col As Column
    
    On Error GoTo ErrorHandler
    BooleanFormat.Format = "Yes/No"
    CurrencyFormat.Format = "Currency"
    DateFormat.Format = "dd-MMM-yyyy hh:nn AMPM"
    
    Set dgdList.DataSource = Nothing
    txtResults.Text = vbNullString
    txtResults.Visible = True
    txtResults.SetFocus
    dgdList.Visible = False
    
    fActiveTrans = False
    Select Case UCase(Mid(txtSQL.Text, 1, 6))
        Case "UPDATE", "DELETE"
            cnSQL.BeginTrans
            fActiveTrans = True
            Set rsList = cnSQL.Execute(txtSQL.Text, RecordsAffected)
            dgdList.Visible = False
            txtResults.Visible = True
        Case "SELECT"
            CloseRecordset rsList, True
            Set rsList = New ADODB.Recordset
            rsList.CursorLocation = adUseClient
            rsList.Open txtSQL.Text, cnSQL, adOpenKeyset, adLockReadOnly
            Set dgdList.DataSource = rsList
            ReDim SortDESC(0 To dgdList.Columns.Count - 1)
            
            For Each fld In rsList.Fields
                Set col = dgdList.Columns(fld.Name)
                Select Case fld.Type
                    Case adCurrency
                        Set col.DataFormat = CurrencyFormat
                        col.Alignment = dbgRight
                    Case adBoolean
                        Set col.DataFormat = BooleanFormat
                        col.Alignment = dbgCenter
                    Case adDate, adDBDate
                        Set col.DataFormat = DateFormat
                        col.Alignment = dbgCenter
                    Case Else
                        col.Alignment = dbgGeneral
                End Select
            Next
            dgdList.Visible = True
            txtResults.Visible = False
    End Select
    
    If cnSQL.Errors.Count > 0 Then
        For Each adoError In cnSQL.Errors
            If Trim(adoError.Description) <> vbNullString Then
                strOutput = adoError.Description & "(" & Hex(adoError.Number) & ")" & vbCrLf
                strOutput = strOutput & vbTab & "Source: " & adoError.Source & vbCrLf & _
                    vbTab & "SQL State: " & adoError.SQLState & vbCrLf & _
                    vbTab & "Native Error: " & adoError.NativeError & vbCrLf
                ErrorCount = ErrorCount + 1
                Select Case UCase(Mid(txtSQL.Text, 1, 6))
                    Case "DELETE", "UPDATE"
                        txtResults = strOutput
                        If ErrorCount > 0 Then
                            If fActiveTrans Then cnSQL.RollbackTrans
                            Exit Sub
                        End If
                    Case "SELECT"
                        If rsList.RecordCount = 0 Then MsgBox strOutput, vbExclamation, Me.Caption
                    Case Else
                End Select
            End If
        Next
    End If
    
    Select Case UCase(Mid(txtSQL.Text, 1, 6))
        Case "DELETE"
            fResponse = MsgBox(RecordsAffected & " Records deleted... Commit transaction?", vbYesNo, Me.Caption) = vbYes
        Case "UPDATE"
            fResponse = MsgBox(RecordsAffected & " Records updated... Commit transaction?", vbYesNo, Me.Caption) = vbYes
        Case "SELECT"
            RecordsAffected = rsList.RecordCount
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
            strOutput = RecordsAffected & " record(s) read"
        Case "DELETE"
            strOutput = RecordsAffected & " record(s) deleted"
        Case "UPDATE"
            strOutput = RecordsAffected & " record(s) updated"
    End Select
    sbStatus.Panels("Message").Text = strOutput
    txtSQL.SetFocus
    Exit Sub
    
ErrorHandler:
    Select Case Err.Number
        'Case 7, -2147024882  'Out of memory
        '    eMessage = " (truncated at " & Format(Len(txtResults.Text), "#,###") & " bytes)"
        Case Else
            MsgBox Err.Description & " (" & Err.Number & ")", vbExclamation, Me.Caption
    End Select
    'BufferLimit = Len(txtResults.Text)
    If fActiveTrans Then cnSQL.RollbackTrans
    Resume Next
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
Private Sub dgdList_DblClick()
    Dim col As Column
    Dim ColRight As Single
    Dim ColumnFormat As New StdDataFormat
    Dim DataWidth As Long
    Dim iCol As Integer
    Dim ResizeWindow As Single
    Dim rsTemp As ADODB.Recordset
    Dim WidestData As Long
    
    Me.MousePointer = vbHourglass
    
    ResizeWindow = 36
    For iCol = dgdList.LeftCol To dgdList.Columns.Count - 1
        Set col = dgdList.Columns(iCol)
        If col.Visible And col.Width > 0 Then
            ColRight = col.Left + col.Width
            If MouseY <= col.Top And MouseX >= (ColRight - ResizeWindow) And MouseX <= (ColRight + ResizeWindow) Then
                dgdList.ClearSelCols
                lblA.Caption = col.Caption
                WidestData = lblA.Width
                Set ColumnFormat = col.DataFormat
                Set rsTemp = rsList.Clone(adLockReadOnly)
                rsTemp.MoveFirst
                While Not rsTemp.EOF
                    If Not IsNull(rsTemp(col.Caption).Value) Then
                        If Not ColumnFormat Is Nothing Then
                            lblA.Caption = Format(rsTemp(col.Caption).Value, col.DataFormat.Format)
                        Else
                            lblA.Caption = CStr(rsTemp(col.Caption).Value)
                        End If
                        DataWidth = lblA.Width
                        If DataWidth > WidestData Then WidestData = DataWidth
                    End If
                    rsTemp.MoveNext
                Wend
                CloseRecordset rsTemp, True
                Set ColumnFormat = Nothing
                col.Width = WidestData + (4 * ResizeWindow)
                If col.Width > dgdList.Width Then col.Width = col.Width - ResizeWindow
                GoTo ExitSub
            End If
        End If
    Next iCol
    
ExitSub:
    Me.MousePointer = vbDefault
End Sub
Private Sub dgdList_HeadClick(ByVal ColIndex As Integer)
    If SortDESC(ColIndex) Then
        rsList.Sort = dgdList.Columns(ColIndex).Caption & " DESC"
    Else
        rsList.Sort = dgdList.Columns(ColIndex).Caption & " ASC"
    End If
    SortDESC(ColIndex) = Not SortDESC(ColIndex)
End Sub
Private Sub dgdList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X
    MouseY = Y
End Sub
Private Sub Form_Activate()
    Dim SQLsource As String
    
    sbStatus.Panels("Status").Text = "SQL"
    dgdList.Visible = False
    txtResults.Visible = True
    
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
    txtSQL.SetFocus
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
    dgdList.Move MarginTwips, 3 * MarginTwips, frameResults.Width - (2 * MarginTwips), frameResults.Height - (4 * MarginTwips)
    txtResults.Move MarginTwips, 3 * MarginTwips, frameResults.Width - (2 * MarginTwips), frameResults.Height - (4 * MarginTwips)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    CloseRecordset rsList, True
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
