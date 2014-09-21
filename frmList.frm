VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{B513F5C0-5FB8-11D3-BC8C-00104B9E0792}#1.0#0"; "SSULTRAGRID.OCX"
Begin VB.Form frmList 
   Caption         =   "List"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   ScaleHeight     =   2475
   ScaleWidth      =   5985
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picScaleObj 
      Height          =   495
      Left            =   4620
      ScaleHeight     =   435
      ScaleWidth      =   675
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   2160
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   794
            MinWidth        =   804
            Picture         =   "frmList.frx":0000
            Key             =   "Top"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   794
            MinWidth        =   804
            Picture         =   "frmList.frx":045C
            Key             =   "Bottom"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   794
            MinWidth        =   804
            Picture         =   "frmList.frx":08B8
            Key             =   "Filter"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1270
            MinWidth        =   1270
            Key             =   "Position"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "Status"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Key             =   "Message"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   2
            Object.Width           =   1376
            MinWidth        =   1270
            TextSave        =   "10:44 PM"
            Key             =   "Time"
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid dgdList 
      Height          =   612
      Left            =   3420
      TabIndex        =   0
      Top             =   600
      Width           =   1212
      _ExtentX        =   2143
      _ExtentY        =   1085
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   16
      TabAction       =   2
      WrapCellPointer =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
   Begin UltraGrid.SSUltraGrid ssugList 
      Height          =   1995
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   3519
      _Version        =   65536
      GridFlags       =   263168
      LayoutFlags     =   5505028
      BorderStyle     =   1
      RowConnectorStyle=   2
      TabNavigation   =   1
      ViewStyle       =   0
      ViewStyleBand   =   1
      InterBandSpacing=   0
      AlphaBlendEnabled=   0   'False
      Enabled         =   0   'False
      Bands           =   "frmList.frx":0D10
      Override        =   "frmList.frx":1040
      Appearance      =   "frmList.frx":1094
   End
   Begin VB.Label lblA 
      AutoSize        =   -1  'True
      Caption         =   "A"
      Height          =   195
      Left            =   5460
      TabIndex        =   2
      Top             =   1980
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Menu mnuList 
      Caption         =   "&List"
      Visible         =   0   'False
      Begin VB.Menu mnuListEdit 
         Caption         =   "&Edit Record"
      End
      Begin VB.Menu mnuListSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListNew 
         Caption         =   "&New Record"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuListCopy 
         Caption         =   "&Copy/Append Record"
      End
      Begin VB.Menu mnuListDelete 
         Caption         =   "&Delete Record"
      End
   End
End
Attribute VB_Name = "frmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const ssugWindowBackground As Long = &H80000005
Const ssugButtonFace As Long = &H8000000F
Const ssugBlack As Long = &H0
Const ssugWhite As Long = &HFFFFFF

Public WithEvents vrsList As ADODB.Recordset
Attribute vrsList.VB_VarHelpID = -1
Public rsList As ADODB.Recordset
Private RS As ADODB.Recordset
Private rsTemp As ADODB.Recordset
Public HiddenFields As String
Private ColLinkMap() As Boolean 'Used to denote fields based on linkage variables...
Private Key As String
'Private MouseY As Single
'Private MouseX As Single
'Private SortDESC() As Boolean
Private fAllowEditMode As Boolean
Private fEditMode As Boolean
Private EditRow As Long
'Private Sub dgdList_AfterColEdit(ByVal ColIndex As Integer)
'    Call Trace(trcBody, "dgdList_AfterColEdit(" & ColIndex & ")")
'End Sub
'Private Sub dgdList_AfterColUpdate(ByVal ColIndex As Integer)
'    Call Trace(trcBody, "dgdList_AfterColUpdate(" & ColIndex & ")")
'End Sub
'Private Sub dgdList_AfterDelete()
'    Call Trace(trcBody, "dgdList_AfterDelete")
'End Sub
'Private Sub dgdList_AfterInsert()
'    Call Trace(trcBody, "dgdList_AfterInsert")
'End Sub
'Private Sub dgdList_AfterUpdate()
'    Call Trace(trcBody, "dgdList_AfterUpdate")
'End Sub
'Private Sub dgdList_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
'    Call Trace(trcBody, "dgdList_BeforeColEdit(" & ColIndex & ", " & KeyAscii & ", Cancel)")
'End Sub
'Private Sub dgdList_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
'    Call Trace(trcBody, "dgdList_BeforeColUpdate(" & ColIndex & ", """ & OldValue & """, Cancel)")
'    If Not fEditMode Then Cancel = 1
'End Sub
'Private Sub dgdList_BeforeDelete(Cancel As Integer)
'    Call Trace(trcBody, "dgdList_BeforeDelete(Cancel)")
'End Sub
'Private Sub dgdList_BeforeInsert(Cancel As Integer)
'    Call Trace(trcBody, "dgdList_BeforeInsert(Cancel)")
'End Sub
'Private Sub dgdList_BeforeUpdate(Cancel As Integer)
'    Call Trace(trcBody, "dgdList_BeforeUpdate(Cancel)")
'    If Not fEditMode Then Cancel = 1
'End Sub
'Private Sub dgdList_Click()
'    UpdatePosition
'End Sub
'Private Sub dgdList_ColEdit(ByVal ColIndex As Integer)
'    Call Trace(trcBody, "dgdList_ColEdit(" & ColIndex & ")")
'    fEditMode = True
'    sbStatus.Panels("Status").Text = "Edit Mode"
'End Sub
'Private Sub dgdList_DblClick()
'    Dim col As Column
'    Dim ColRight As Single
'    Dim iCol As Integer
'    Dim ResizeWindow As Single
'
'    Me.MousePointer = vbHourglass
'
'    ResizeWindow = 36
'    For iCol = dgdList.LeftCol To dgdList.Columns.Count - 1
'        Set col = dgdList.Columns(iCol)
'        If col.Visible And col.Width > 0 Then
'            ColRight = col.Left + col.Width
'            If MouseY <= col.Top And MouseX >= (ColRight - ResizeWindow) And MouseX <= (ColRight + ResizeWindow) Then
'                dgdList.ClearSelCols
'                Call ResizeColumn(col)
'                GoTo ExitSub
'            End If
'        End If
'    Next iCol
'
'ExitSub:
'    Me.MousePointer = vbDefault
'End Sub
'Private Sub dgdList_HeadClick(ByVal ColIndex As Integer)
'    Dim ColName As String
'    Dim fld As ADODB.Field
'    Dim col As Column
'    Dim iCol As Integer
'    Dim ResizeWindow As Single
'
'    On Error Resume Next
'    ColName = dgdList.Columns(ColIndex).Caption
'    If RS.BOF And RS.EOF Then Exit Sub
'    'kfc - 05/27/00; This seems to be working now without reassignment of the DataSource (and the
'    '                residual code required to re-establish the look of the original data)...
'    '                Could be changing the JET driver back to 3.51 or reinstalling MDAC...(???)
'    'Set dgdList.DataSource = Nothing
'    RS.Sort = vbNullString
'    If SortDESC(ColIndex) Then
'        ColName = ColName & " DESC"
'    Else
'        ColName = ColName & " ASC"
'    End If
'    SortDESC(ColIndex) = Not SortDESC(ColIndex)
'    RS.Sort = ColName
'
''    'Working around bug Q230167...
''    If Not rsTemp Is Nothing Then
''        CloseRecordset rsTemp, False
''    Else
''        Set rsTemp = New ADODB.Recordset
''    End If
''
''    For Each fld In RS.Fields
''        rsTemp.Fields.Append fld.Name, fld.Type, fld.DefinedSize, fld.Attributes
''    Next fld
''    'Add the hidden field (assuming the value does not matter - usually used for Grids)...
''    'If Not IsMissing(HiddenFieldName) Then rsTemp.Fields.Append HiddenFieldName, adVarChar, 1
''    rsTemp.CursorType = adOpenStatic    'Updatable snapshot
''    rsTemp.LockType = adLockOptimistic  'Allow updates
''    rsTemp.Open
''
''    'Copy the data from the real recordset to the virtual one...
''    If Not (RS.BOF And RS.EOF) Then
''        RS.MoveFirst
''        While Not RS.EOF
''            'Populate the grid with the recordset data...
''            rsTemp.AddNew
''            For Each fld In RS.Fields
''                rsTemp(fld.Name).Value = RS(fld.Name).Value
''            Next fld
''            rsTemp.Update
''            RS.MoveNext
''        Wend
''        rsTemp.MoveFirst
''    End If
''    RS.Close
''    Set RS = Nothing
''
''    rsTemp.Sort = ColName
''    Set RS = rsTemp
''    Set rsTemp = Nothing

'kfc - 05/27/00; This seems to be working now without reassignment of the DataSource (and the
'                residual code required to re-establish the look of the original data)...
'                Could be changing the JET driver back to 3.51 or reinstalling MDAC...(???)
'    Set dgdList.DataSource = RS
'    ResetColumns
'
'    'Resize Columns based on content...
'    ResizeWindow = 36
'    For iCol = dgdList.LeftCol To dgdList.Columns.Count - 1
'        Set col = dgdList.Columns(iCol)
'        If col.Visible And col.Width > 0 Then
'            dgdList.ClearSelCols
'            Call ResizeColumn(col)
'        End If
'    Next iCol
'End Sub
'Private Sub dgdList_KeyUp(KeyCode As Integer, Shift As Integer)
'    Dim col As Column
'    Dim sCol As Column
'    Dim i As Integer
'
'    On Error Resume Next
'
'    Select Case KeyCode
'        Case vbKeyTab
'            Call Trace(trcBody, "dgdList_KeyUp(" & KeyCode & ", " & Shift & ") - Row: " & dgdList.BookMark & "; Column: " & dgdList.Columns(dgdList.col).Caption & "(" & dgdList.col & ")")
'        Case vbKeyEscape
'            If RS.EditMode = adEditInProgress Then RS.Update
'            fEditMode = False
'            EditRow = -1
'
'            If Not dgdList.EditActive Then dgdList.EditActive = False
'            For Each col In dgdList.Columns
'                If Not col.Locked Then col.Locked = True
'            Next col
'            dgdList.AllowUpdate = False
'            dgdList.TabAction = dbgControlNavigation
'            sbStatus.Panels("Status").Text = vbNullString
'        Case vbKeyF2
'            If Not fEditMode Then
'                fEditMode = True
'                EditRow = dgdList.Row
'                dgdList.AllowUpdate = True
'                dgdList.TabAction = dbgGridNavigation
'
'                'Position to the first non-hidden column...
'                Set sCol = Nothing
'                For Each col In dgdList.Columns
'                    col.Locked = False
'                    If sCol Is Nothing And col.Visible Then Set sCol = col
'                Next col
'                dgdList.col = sCol.ColIndex
'
'                dgdList.EditActive = True
'                sbStatus.Panels("Status").Text = "Edit Mode"
'                For i = 0 To dgdList.SelBookmarks.Count - 1
'                    dgdList.SelBookmarks.Remove 0
'                Next i
'                dgdList.SelBookmarks.Add dgdList.BookMark
'            ElseIf Not dgdList.EditActive Then
'                dgdList.EditActive = True
'                UpdatePosition
'            End If
'        Case vbKeyUp
'            If dgdList.EditActive Then
'                'Leave the current cell...
'                dgdList.EditActive = False
'                'Reposition to the previous row...
'                dgdList.Row = dgdList.Row - 1
'                'Force edit mode on that cell without requiring the user to type a key...
'                dgdList.EditActive = True
'                dgdList.SelStart = Len(dgdList.Columns(dgdList.col).Text)
'                UpdatePosition
'            End If
'        Case vbKeyDown
'            If dgdList.EditActive Then
'                'Leave the current cell...
'                dgdList.EditActive = False
'                'Reposition to the next row...
'                dgdList.Row = dgdList.Row + 1
'                'Force edit mode on that cell without requiring the user to type a key...
'                dgdList.EditActive = True
'                dgdList.SelStart = Len(dgdList.Columns(dgdList.col).Text)
'                UpdatePosition
'            End If
'    End Select
'    Set col = Nothing
'    Set sCol = Nothing
'End Sub
'Private Sub dgdList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    MouseX = X
'    MouseY = Y
'End Sub
'Private Sub dgdList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = vbKeyRButton And Not dgdList.EditActive And sbStatus.Panels("Status").Text <> "Edit Mode" Then
'        PopupMenu mnuList
'    End If
'End Sub
'Private Sub dgdList_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'    Dim col As Column
'    Dim i As Long
'
'    If Not IsEmpty(LastRow) And Not IsNull(LastRow) Then
'        'dgdList_KeyUp vbKeyEscape, 0
'        If RS.EditMode = adEditInProgress Then RS.Update
'    End If
'
'    If Not fEditMode Then
'        For i = 0 To dgdList.SelBookmarks.Count - 1
'            dgdList.SelBookmarks.Remove 0
'        Next i
'        If Not IsNull(dgdList.BookMark) Then dgdList.SelBookmarks.Add dgdList.BookMark
'
'        'Find a hidden column and use it to select the whole row...
'        For Each col In dgdList.Columns
'            If Not col.Visible Then
'                dgdList.col = col.ColIndex
'                Exit For
'            End If
'        Next col
'
'        'dgdList.col = dgdList.Columns("Junk").ColIndex
'    End If
'    UpdatePosition
'End Sub
'Private Sub dgdList_RowResize(Cancel As Integer)
'    Call Trace(trcBody, "dgdList_RowResize(Cancel) - RowHeight: " & dgdList.RowHeight)
'End Sub
Private Sub Form_Activate()
    Dim i As Integer
    Dim ssCol As UltraGrid.SSColumn
    Dim col As Column
    Dim fld As ADODB.Field
    Dim ResizeWindow As Single
    
    If vrsList Is Nothing Then
        Set RS = frmList.rsList
        fAllowEditMode = True
    Else
        Set RS = frmList.vrsList
        fAllowEditMode = False
    End If
'    Set dgdList.DataSource = RS
'    dgdList.AllowUpdate = fAllowEditMode
'    dgdList.AllowAddNew = fAllowEditMode
    
    Set ssugList.DataSource = Nothing
    Set ssugList.DataSource = RS
'    ssugMain.Bands(0).Columns("HOSTDB.DATABASE_NAME").Header.Caption = "Database"
'    ssugMain.Bands(0).Columns("ORIGIN.ALPHA_SORT").Header.Caption = "Origin"
'    ssugMain.Bands(0).Columns("COST_CENTER.CENTER_NAME").Header.Caption = "Cost Center"
'    ssugMain.Bands(0).Columns("OPTCODE.OPTCODE_NAME").Header.Caption = "Opt Code"
'    ssugMain.Bands(0).Columns("BL.HOSTDB_NUMBER").Hidden = True
'    ssugMain.Bands(0).Columns("BL.ORIGIN_NUMBER").Hidden = True
'    ssugMain.Bands(0).Columns("BL.OPTCODE_NUMBER").Hidden = True
'    ssugMain.Bands(0).Columns("BL.OPTCODE_CODE").Hidden = True
'    ssugMain.Bands(0).Columns("BL.CENTER_NUMBER").Hidden = True
'    ssugMain.Bands(0).Columns("BL.BOOKING_CODE").Header.Caption = "Booking #"
    
    For Each ssCol In ssugList.Bands(0).Columns
        ssCol.Header.Appearance.TextAlign = ssAlignLeft
        ssCol.Activation = ssActivationActivateOnly
        If Not ssCol.Hidden And ssCol.Width > 0 Then AutoSizeColumn ssugList, ssCol, True, Me.ScaleMode, picScaleObj
    Next ssCol
    
'    ReDim SortDESC(0 To RS.Fields.Count - 1)
'    ReDim ColLinkMap(0 To RS.Fields.Count - 1)
'    ResetColumns
'
'    'Get the column settings for the display...
'    For i = 0 To dgdList.Columns.Count - 1
'        dgdList.Columns(i).Width = GetSetting(App.ProductName, Me.Caption & " Settings", dgdList.Columns(i).Caption & " Width", dgdList.Columns(i).Width)
'    Next
    Me.Top = GetSetting(App.ProductName, Me.Caption & " Settings", "Form Top", Me.Top)
    Me.Left = GetSetting(App.ProductName, Me.Caption & " Settings", "Form Left", Me.Left)
    Me.Width = GetSetting(App.ProductName, Me.Caption & " Settings", "Form Width", Me.Width)
    Me.Height = GetSetting(App.ProductName, Me.Caption & " Settings", "Form Height", Me.Height)
        
    If SQLfilter <> vbNullString Then
        sbStatus.Panels("Message").Text = "Filter: " & SQLfilter
    Else
        sbStatus.Panels("Message").Text = vbNullString
    End If
    
'    'Resize Columns based on content...
'    ResizeWindow = 36
'    For i = dgdList.LeftCol To dgdList.Columns.Count - 1
'        Set col = dgdList.Columns(i)
'        If col.Visible And col.Width > 0 Then
'            dgdList.ClearSelCols
'            Call ResizeColumn(col)
'        End If
'    Next i
'
'    dgdList_Click
End Sub
Private Sub Form_Load()
    Me.Icon = Forms(Forms.Count - 2).Icon
'    dgdList.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
'    dgdList.RowHeight = 192.189     'so I don't forget (this is for MS Sans Serif 8 point font)
    dgdList.Visible = False
    ssugList.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - sbStatus.Height
    ssugList.Enabled = True
End Sub
Private Sub Form_Resize()
    If Me.Width < MinWidth Then Me.Width = MinWidth
    If Me.Height < MinHeight Then Me.Height = MinHeight
    dgdList.Width = Me.ScaleWidth
    dgdList.Height = Me.ScaleHeight - sbStatus.Height
    ssugList.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - sbStatus.Height
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    
    On Error Resume Next
    If Not RS Is Nothing Then
        If (RS.Status And adRecPendingChanges) = adRecPendingChanges Then RS.UpdateBatch
    End If
    
'    'Save the column settings for the next display...
'    For i = 0 To dgdList.Columns.Count - 1
'        SaveSetting App.ProductName, Me.Caption & " Settings", dgdList.Columns(i).Caption & " Width", dgdList.Columns(i).Width
'    Next
    SaveSetting App.ProductName, Me.Caption & " Settings", "Form Top", Me.Top
    SaveSetting App.ProductName, Me.Caption & " Settings", "Form Left", Me.Left
    SaveSetting App.ProductName, Me.Caption & " Settings", "Form Width", Me.Width
    SaveSetting App.ProductName, Me.Caption & " Settings", "Form Height", Me.Height
End Sub
Private Sub mnuListCopy_Click()
    Dim Table As String
    Dim FieldList As String
    Dim Values As String
    Dim fld As ADODB.Field
    Dim RecordsAffected As Long
    
    On Error GoTo ErrorHandler
        
    Table = RS.Fields(0).Properties("BASETABLENAME")
    For Each fld In RS.Fields
        If (RS(fld.Name).Attributes And adFldUpdatable) = adFldUpdatable Then
            FieldList = FieldList & "[" & fld.Name & "],"
            If IsNull(fld.Value) Then
                Values = Values & "Null,"
            Else
                Select Case fld.Type
                    Case adCurrency
                        Values = Values & fld.Value & ","
                    Case adBoolean
                        Values = Values & fld.Value & ","
                    Case adDate, adDBDate, adDBTimeStamp
                        Values = Values & "#" & fld.Value & "#,"
                    Case adBinary, adLongVarBinary, adLongVarChar, adChar, adVarChar
                        Values = Values & "'" & SQLQuote(fld.Value) & "',"
                    Case Else
                        Values = Values & "'" & SQLQuote(fld.Value) & "',"
                End Select
            End If
        End If
    Next fld
    FieldList = Mid(FieldList, 1, Len(FieldList) - 1)
    Values = Mid(Values, 1, Len(Values) - 1)
    
    adoConn.Execute "insert into [" & Table & "] (" & FieldList & ") values (" & Values & ")", RecordsAffected
    RefreshCommand RS, SQLkey
    
ExitSub:
    Exit Sub
    
ErrorHandler:
    Dim errorCode As Long
    MsgBox BuildADOerror(adoConn, errorCode), vbCritical, "frmList.mnuListCopy"
    GoTo ExitSub
    Resume Next
End Sub
Private Sub mnuListDelete_Click()
    Dim Table As String
    Dim RecordsAffected As Long
    
    On Error GoTo ErrorHandler
        
    If MsgBox("Are you sure you want to delete record #" & RS.BookMark & "...?", vbYesNo, Me.Caption) = vbYes Then
        'RS.Delete adAffectCurrent
        'RS.Update
        Table = RS.Fields("ID").Properties("BASETABLENAME")
        adoConn.Execute "delete from [" & Table & "] where ID=" & RS("ID"), RecordsAffected
        RefreshCommand RS, SQLkey
    End If

ExitSub:
    Exit Sub
    
ErrorHandler:
    Dim errorCode As Long
    MsgBox BuildADOerror(adoConn, errorCode), vbCritical, "frmList.mnuListDelete"
    GoTo ExitSub
    Resume Next
End Sub
Private Sub mnuListEdit_Click()
    Call UpdateMode(vbKeyF2, 0)
End Sub
Private Sub mnuListNew_Click()
    MsgBox "Sorry, New is not implemented yet...", vbExclamation, Me.Caption
End Sub
Private Sub ResetColumns()
    Dim ssCol As UltraGrid.SSColumn
    Dim col As Column
    Dim ctl As Control
    Dim fld As ADODB.Field
    Dim frm As Form
    Dim fFoundControl As Boolean

    Set frm = Forms(Forms.Count - 2)
    For Each fld In RS.Fields
        Set col = dgdList.Columns(fld.Name)

'    ssugMain.Bands(0).Columns("HOSTDB.DATABASE_NAME").Header.Caption = "Database"
        'Set ssCol = ssugList.Bands(0).Columns(fld.Name)
        col.Visible = True
        col.Locked = True
        Select Case fld.Type
            Case adCurrency
                col.NumberFormat = "Currency"
                col.Alignment = dbgRight
            Case adBoolean
                col.NumberFormat = "Yes/No"
                col.Alignment = dbgCenter
            Case adDate, adDBDate
                col.NumberFormat = fmtDate
                col.Alignment = dbgCenter
            Case adBinary, adLongVarBinary, adLongVarChar
                col.Visible = False
            Case Else
                col.Alignment = dbgGeneral
        End Select
        If fld.Name = "Junk" Then col.Visible = False

        'I'm not taking the time to do what Access does in displying Lookup fields...
        'I'm just going to live with making such linkage fields invisible...
        fFoundControl = False
        If Not (Len(col.Caption) > 2 And Right(col.Caption, 2) = "ID") Then
            For Each ctl In frm.Controls
                If ctl.Tag = vbNullString Then GoTo SkipControl
                Select Case TypeName(ctl)
                    Case "CheckBox", "DataCombo", "Label", "PictureBox", "RichTextBox", "TextBox"
                        If UCase(ctl.DataField) = UCase(fld.Name) Then
                            If ctl.Tag <> vbNullString Then fFoundControl = True
                            Exit For
                        End If
                End Select
SkipControl:
            Next ctl
        End If
        If fFoundControl Then
            ColLinkMap(col.ColIndex) = False
            col.Visible = True
        Else
            ColLinkMap(col.ColIndex) = True
            col.Visible = False
        End If
    Next fld

    Set fld = Nothing
    Set col = Nothing
    Set ctl = Nothing
    Set frm = Nothing
End Sub
'Private Sub ResizeColumn(col As Column)
'    Dim ColumnFormat As String
'    Dim DataWidth As Long
'    Dim ResizeWindow As Single
'    Dim rsTemp As ADODB.Recordset
'    Dim WidestData As Long
'
'    Me.MousePointer = vbHourglass
'
'    ResizeWindow = 36
'    lblA.Caption = col.Caption
'    WidestData = lblA.Width
'    ColumnFormat = col.NumberFormat
'    If Not RS.BOF And Not RS.EOF Then
'        Set rsTemp = RS.Clone(adLockReadOnly)
'        rsTemp.MoveFirst
'        While Not rsTemp.EOF
'            If Not IsNull(rsTemp(col.Caption).Value) Then
'                If ColumnFormat <> vbNullString Then
'                    lblA.Caption = Format(rsTemp(col.Caption).Value, col.NumberFormat)
'                Else
'                    lblA.Caption = CStr(rsTemp(col.Caption).Value)
'                End If
'                DataWidth = lblA.Width
'                If DataWidth > WidestData Then WidestData = DataWidth
'            End If
'            rsTemp.MoveNext
'        Wend
'        CloseRecordset rsTemp, True
'    End If
'    col.Width = WidestData + (4 * ResizeWindow)
'    If col.Width > dgdList.Width Then col.Width = col.Width - ResizeWindow
'
'ExitSub:
'    Me.MousePointer = vbDefault
'End Sub
Private Sub sbStatus_PanelClick(ByVal Panel As MSComCtlLib.Panel)
    Dim frm As Form
    
    Select Case UCase(Panel.Key)
        Case "TOP"
            RS.MoveFirst
            UpdatePosition
        Case "BOTTOM"
            RS.MoveLast
            UpdatePosition
        Case "FILTER"
            FilterCommand Me, RS, ""
        Case Else
    End Select
End Sub
Private Sub ssugList_AfterEnterEditMode()
    fEditMode = True
    UpdatePosition
End Sub
Private Sub ssugList_AfterExitEditMode()
    fEditMode = False
    UpdatePosition
End Sub
Private Sub ssugList_AfterRowActivate()
    'ssugList.ActiveRow.Appearance.ForeColor = ssugWindowBackground
End Sub
Private Sub ssugList_AfterSelectChange(ByVal SelectChange As UltraGrid.Constants_SelectChange)
    If SelectChange = ssSelectChangeRow Then fEditMode = False
    UpdatePosition
End Sub
Private Sub ssugList_BeforeRowActivate(ByVal Row As UltraGrid.SSRow)
    Row.Appearance.BackColor = ssugList.Override.SelectedRowAppearance.BackColor
    Row.Appearance.ForeColor = ssugList.Override.SelectedRowAppearance.ForeColor
End Sub
Private Sub ssugList_BeforeRowDeactivate(ByVal Cancel As UltraGrid.SSReturnBoolean)
    ssugList.ActiveRow.Appearance.BackColor = ssugList.Override.RowAppearance.BackColor
    ssugList.ActiveRow.Appearance.ForeColor = ssugList.Override.RowAppearance.ForeColor
End Sub
Private Sub ssugList_DblClick()
    Const fIncludeHeader As Boolean = True
    'try to autosize the column based on the text
    AutoSizeColFromMessagePos ssugList, fIncludeHeader, Me.ScaleMode, picScaleObj
End Sub
Private Sub ssugList_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Dim i As Integer
    ssugList.Layout.Override.AllowColSizing = ssAllowColSizingFree
    ssugList.Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    'ssugList.Layout.Override.AllowColSwapping = ssAllowColSwappingWithinGroup
    ssugList.Override.RowAlternateAppearance.BackColor = ssugButtonFace
    ssugList.Override.RowAppearance.BackColor = ssugButtonFace
    ssugList.Override.SelectedRowAppearance.BackColor = ssugBlack
    ssugList.Override.SelectedRowAppearance.ForeColor = ssugWhite
    For i = 0 To ssugList.Bands.Count - 1
        ssugList.Bands(i).Override.RowAlternateAppearance.BackColor = ssugButtonFace
        ssugList.Bands(i).Override.RowAppearance.BackColor = ssugButtonFace
    Next i
End Sub
Private Sub ssugList_KeyUp(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
    Call UpdateMode(KeyCode.Value, Shift)
End Sub
Private Sub ssugList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbKeyRButton And sbStatus.Panels("Status").Text <> "Edit Mode" Then
        PopupMenu mnuList
    End If
    ssugList.Refresh ssRefreshDisplay
End Sub
Private Sub ssugList_OnKillFocus(ByVal hWndGettingFocus As Stdole.OLE_HANDLE)
    Dim ssCancel As UltraGrid.SSReturnBoolean
    Call ssugList_BeforeRowDeactivate(ssCancel)
End Sub
Private Sub UpdateMode(KeyCode As Integer, Shift As Integer)
    Dim i As Integer

    On Error Resume Next

    Select Case KeyCode
        Case vbKeyTab
            Call Trace(trcBody, "UpdateMode(" & KeyCode & ", " & Shift & ") - Row: " & dgdList.BookMark & "; Column: " & dgdList.Columns(dgdList.col).Caption & "(" & dgdList.col & ")")
        Case vbKeyEscape
            If RS.EditMode = adEditInProgress Then RS.CancelUpdate
            fEditMode = False
            EditRow = -1
            UpdatePosition
        Case vbKeyF2
            fEditMode = True
            UpdatePosition
        Case vbKeyUp
'            If dgdList.EditActive Then
'                'Leave the current cell...
'                dgdList.EditActive = False
'                'Reposition to the previous row...
'                dgdList.Row = dgdList.Row - 1
'                'Force edit mode on that cell without requiring the user to type a key...
'                dgdList.EditActive = True
'                dgdList.SelStart = Len(dgdList.Columns(dgdList.col).Text)
'                UpdatePosition
'            End If
        Case vbKeyDown
'            If dgdList.EditActive Then
'                'Leave the current cell...
'                dgdList.EditActive = False
'                'Reposition to the next row...
'                dgdList.Row = dgdList.Row + 1
'                'Force edit mode on that cell without requiring the user to type a key...
'                dgdList.EditActive = True
'                dgdList.SelStart = Len(dgdList.Columns(dgdList.col).Text)
'                UpdatePosition
'            End If
    End Select
End Sub
Private Sub UpdatePosition()
    If (RS.State And adStateOpen) <> adStateOpen Then Exit Sub
    sbStatus.Panels("Position").Text = "Record " & ssugList.ActiveRow.BookMark & " of " & RS.RecordCount
    If fEditMode Then
        If ssugList.ActiveCell Is Nothing Then ssugList.ActiveCell = ssugList.ActiveRow.Cells(0)
        sbStatus.Panels("Status").Text = "Edit Mode (" & ssugList.ActiveCell.Column.BaseColumnName & ")"
        ssugList.Override.AllowUpdate = ssAllowUpdateYes
        ssugList.Override.CellClickAction = ssClickActionEdit
        ssugList.Override.EditCellAppearance.BackColor = vbWindowBackground
        ssugList.Override.EditCellAppearance.ForeColor = vbWindowText
        ssugList.TabNavigation = ssTabNavigationNextCell
        If ssugList.ActiveCell.DataChanged Then sbStatus.Panels("Status").Text = sbStatus.Panels("Status").Text & " *"
    Else
        sbStatus.Panels("Status").Text = vbNullString
        ssugList.Override.AllowUpdate = ssAllowUpdateNo
        ssugList.Override.CellClickAction = ssClickActionRowSelect
        ssugList.Override.EditCellAppearance.BackColor = vbButtonFace
        ssugList.Override.EditCellAppearance.ForeColor = vbWindowBackground
        ssugList.TabNavigation = ssTabNavigationNextControl
        sbStatus.Panels("Status").Text = vbNullString
    End If
End Sub
