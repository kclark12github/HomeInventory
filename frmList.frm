VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmList 
   Caption         =   "List"
   ClientHeight    =   2556
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   6312
   LinkTopic       =   "Form1"
   ScaleHeight     =   2556
   ScaleWidth      =   6312
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   312
      Left            =   0
      TabIndex        =   1
      Top             =   2244
      Width           =   6312
      _ExtentX        =   11134
      _ExtentY        =   550
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   804
            MinWidth        =   804
            Picture         =   "frmList.frx":0000
            Key             =   "Top"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   804
            MinWidth        =   804
            Picture         =   "frmList.frx":045C
            Key             =   "Bottom"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   804
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
            Object.Width           =   3027
            Key             =   "Message"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   2
            Object.Width           =   1270
            MinWidth        =   1270
            TextSave        =   "11:59 PM"
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
      _ExtentX        =   2138
      _ExtentY        =   1080
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   16
      TabAction       =   2
      WrapCellPointer =   -1  'True
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
   Begin VB.Label lblA 
      AutoSize        =   -1  'True
      Caption         =   "A"
      Height          =   192
      Left            =   2700
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   108
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
Public WithEvents vrsList As ADODB.Recordset
Attribute vrsList.VB_VarHelpID = -1
Public rsList As ADODB.Recordset
Private RS As ADODB.Recordset
Private rsTemp As ADODB.Recordset
Public HiddenFields As String
Private Key As String
Private MouseY As Single
Private MouseX As Single
Private SortDESC() As Boolean
Private fAllowEditMode As Boolean
Private fEditMode As Boolean
Private EditRow As Long
Private fDebug As Boolean
Private Sub dgdList_AfterColEdit(ByVal ColIndex As Integer)
    If fDebug Then sbStatus.Panels("Message").Text = "AfterColEdit"
End Sub
Private Sub dgdList_AfterColUpdate(ByVal ColIndex As Integer)
    If fDebug Then sbStatus.Panels("Message").Text = "AfterColUpdate"
End Sub
Private Sub dgdList_AfterDelete()
    If fDebug Then sbStatus.Panels("Message").Text = "AfterDelete"
End Sub
Private Sub dgdList_AfterInsert()
    If fDebug Then sbStatus.Panels("Message").Text = "AfterInsert"
End Sub
Private Sub dgdList_AfterUpdate()
    If fDebug Then sbStatus.Panels("Message").Text = "AfterUpdate"
End Sub
Private Sub dgdList_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    If fDebug Then sbStatus.Panels("Message").Text = "BeforeColEdit"
End Sub
Private Sub dgdList_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    If fDebug Then sbStatus.Panels("Message").Text = "BeforeColUpdate"
    If Not fEditMode Then Cancel = 1
End Sub
Private Sub dgdList_BeforeDelete(Cancel As Integer)
    If fDebug Then sbStatus.Panels("Message").Text = "BeforeDelete"
End Sub
Private Sub dgdList_BeforeInsert(Cancel As Integer)
    If fDebug Then sbStatus.Panels("Message").Text = "BeforeInsert"
End Sub
Private Sub dgdList_BeforeUpdate(Cancel As Integer)
    If fDebug Then sbStatus.Panels("Message").Text = "BeforeUpdate"
    If Not fEditMode Then Cancel = 1
End Sub
Private Sub dgdList_Click()
    UpdatePosition
End Sub
Private Sub dgdList_ColEdit(ByVal ColIndex As Integer)
    fEditMode = True
    sbStatus.Panels("Status").Text = "Edit Mode"
    If fDebug Then sbStatus.Panels("Message").Text = "ColEdit"
End Sub
Private Sub dgdList_DblClick()
    Dim col As Column
    Dim ColRight As Single
    Dim iCol As Integer
    Dim ResizeWindow As Single
    
    Me.MousePointer = vbHourglass
    
    ResizeWindow = 36
    For iCol = dgdList.LeftCol To dgdList.Columns.Count - 1
        Set col = dgdList.Columns(iCol)
        If col.Visible And col.Width > 0 Then
            ColRight = col.Left + col.Width
            If MouseY <= col.Top And MouseX >= (ColRight - ResizeWindow) And MouseX <= (ColRight + ResizeWindow) Then
                dgdList.ClearSelCols
                Call ResizeColumn(col)
                GoTo ExitSub
            End If
        End If
    Next iCol
    
ExitSub:
    Me.MousePointer = vbDefault
End Sub
Private Sub dgdList_HeadClick(ByVal ColIndex As Integer)
    Dim ColName As String
    Dim fld As ADODB.Field
    
    ColName = dgdList.Columns(ColIndex).Caption
    If RS.BOF And RS.EOF Then Exit Sub
    Set dgdList.DataSource = Nothing
    RS.Sort = vbNullString
    If SortDESC(ColIndex) Then
        ColName = ColName & " DESC"
    Else
        ColName = ColName & " ASC"
    End If
    SortDESC(ColIndex) = Not SortDESC(ColIndex)
    RS.Sort = ColName
    
'    'Working around bug Q230167...
'    If Not rsTemp Is Nothing Then
'        CloseRecordset rsTemp, False
'    Else
'        Set rsTemp = New ADODB.Recordset
'    End If
'
'    For Each fld In RS.Fields
'        rsTemp.Fields.Append fld.Name, fld.Type, fld.DefinedSize, fld.Attributes
'    Next fld
'    'Add the hidden field (assuming the value does not matter - usually used for Grids)...
'    'If Not IsMissing(HiddenFieldName) Then rsTemp.Fields.Append HiddenFieldName, adVarChar, 1
'    rsTemp.CursorType = adOpenStatic    'Updatable snapshot
'    rsTemp.LockType = adLockOptimistic  'Allow updates
'    rsTemp.Open
'
'    'Copy the data from the real recordset to the virtual one...
'    If Not (RS.BOF And RS.EOF) Then
'        RS.MoveFirst
'        While Not RS.EOF
'            'Populate the grid with the recordset data...
'            rsTemp.AddNew
'            For Each fld In RS.Fields
'                rsTemp(fld.Name).Value = RS(fld.Name).Value
'            Next fld
'            rsTemp.Update
'            RS.MoveNext
'        Wend
'        rsTemp.MoveFirst
'    End If
'    RS.Close
'    Set RS = Nothing
'
'    rsTemp.Sort = ColName
'    Set RS = rsTemp
'    Set rsTemp = Nothing
    Set dgdList.DataSource = RS
End Sub
Private Sub dgdList_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim col As Column
    Dim sCol As Column
    Dim i As Integer
    
    On Error Resume Next
        
    Select Case KeyCode
        Case vbKeyTab
            If fDebug Then Debug.Print "Row: " & dgdList.BookMark & "; Column: " & dgdList.Columns(dgdList.col).Caption & "(" & dgdList.col & ")"
        Case vbKeyEscape
            If RS.EditMode = adEditInProgress Then RS.Update
            fEditMode = False
            EditRow = -1
            
            If Not dgdList.EditActive Then dgdList.EditActive = False
            For Each col In dgdList.Columns
                If Not col.Locked Then col.Locked = True
            Next col
            dgdList.AllowUpdate = False
            dgdList.TabAction = dbgControlNavigation
            sbStatus.Panels("Status").Text = vbNullString
        Case vbKeyF2
            If Not fEditMode Then
                fEditMode = True
                EditRow = dgdList.Row
                dgdList.AllowUpdate = True
                dgdList.TabAction = dbgGridNavigation
                
                'Position to the first non-hidden column...
                Set sCol = Nothing
                For Each col In dgdList.Columns
                    col.Locked = False
                    If sCol Is Nothing And col.Visible Then Set sCol = col
                Next col
                dgdList.col = sCol.ColIndex
        
                dgdList.EditActive = True
                sbStatus.Panels("Status").Text = "Edit Mode"
                For i = 0 To dgdList.SelBookmarks.Count - 1
                    dgdList.SelBookmarks.Remove 0
                Next i
                dgdList.SelBookmarks.Add dgdList.BookMark
            ElseIf Not dgdList.EditActive Then
                dgdList.EditActive = True
                UpdatePosition
            End If
        Case vbKeyUp
            If dgdList.EditActive Then
                'Leave the current cell...
                dgdList.EditActive = False
                'Reposition to the previous row...
                dgdList.Row = dgdList.Row - 1
                'Force edit mode on that cell without requiring the user to type a key...
                dgdList.EditActive = True
                dgdList.SelStart = Len(dgdList.Columns(dgdList.col).Text)
                UpdatePosition
            End If
        Case vbKeyDown
            If dgdList.EditActive Then
                'Leave the current cell...
                dgdList.EditActive = False
                'Reposition to the next row...
                dgdList.Row = dgdList.Row + 1
                'Force edit mode on that cell without requiring the user to type a key...
                dgdList.EditActive = True
                dgdList.SelStart = Len(dgdList.Columns(dgdList.col).Text)
                UpdatePosition
            End If
    End Select
    Set col = Nothing
    Set sCol = Nothing
End Sub
Private Sub dgdList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X
    MouseY = Y
End Sub
Private Sub dgdList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbKeyRButton And Not dgdList.EditActive And sbStatus.Panels("Status").Text <> "Edit Mode" Then
        PopupMenu mnuList
    End If
End Sub
Private Sub dgdList_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim col As Column
    Dim i As Long
    
    If Not IsEmpty(LastRow) And Not IsNull(LastRow) Then
        'dgdList_KeyUp vbKeyEscape, 0
        If RS.EditMode = adEditInProgress Then RS.Update
    End If
    
    If Not fEditMode Then
        For i = 0 To dgdList.SelBookmarks.Count - 1
            dgdList.SelBookmarks.Remove 0
        Next i
        If Not IsNull(dgdList.BookMark) Then dgdList.SelBookmarks.Add dgdList.BookMark
    
        'Find a hidden column and use it to select the whole row...
        For Each col In dgdList.Columns
            If Not col.Visible Then
                dgdList.col = col.ColIndex
                Exit For
            End If
        Next col
    
        'dgdList.col = dgdList.Columns("Junk").ColIndex
    End If
    UpdatePosition
End Sub
Private Sub dgdList_RowResize(Cancel As Integer)
    If fDebug Then sbStatus.Panels("Message").Text = "RowHeight: " & dgdList.RowHeight
End Sub
Private Sub Form_Activate()
    Dim i As Integer
    Dim BooleanFormat As New StdDataFormat
    Dim CurrencyFormat As New StdDataFormat
    Dim DateFormat As New StdDataFormat
    Dim col As Column
    Dim fld As ADODB.Field
    Dim ResizeWindow As Single
    
    BooleanFormat.Format = "Yes/No"
    CurrencyFormat.Format = "Currency"
    DateFormat.Format = fmtDate
    
    If vrsList Is Nothing Then
        Set RS = frmList.rsList
        fAllowEditMode = True
    Else
        Set RS = frmList.vrsList
        fAllowEditMode = False
    End If
    Set dgdList.DataSource = RS
    dgdList.AllowUpdate = fAllowEditMode
    dgdList.AllowAddNew = fAllowEditMode
    ReDim SortDESC(0 To dgdList.Columns.Count - 1)
    
    For Each fld In RS.Fields
        Set col = dgdList.Columns(fld.Name)
        col.Visible = True
        col.Locked = True
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
            Case adBinary, adLongVarBinary, adLongVarChar
                col.Visible = False
            Case Else
                col.Alignment = dbgGeneral
        End Select
        If fld.Name = "Junk" Then col.Visible = False
    Next
    
    'Get the column settings for the display...
    For i = 0 To dgdList.Columns.Count - 1
        dgdList.Columns(i).Width = GetSetting(App.ProductName, Me.Caption & " Settings", dgdList.Columns(i).Caption & " Width", dgdList.Columns(i).Width)
    Next
    Me.Top = GetSetting(App.ProductName, Me.Caption & " Settings", "Form Top", Me.Top)
    Me.Left = GetSetting(App.ProductName, Me.Caption & " Settings", "Form Left", Me.Left)
    Me.Width = GetSetting(App.ProductName, Me.Caption & " Settings", "Form Width", Me.Width)
    Me.Height = GetSetting(App.ProductName, Me.Caption & " Settings", "Form Height", Me.Height)
        
    If SQLfilter <> vbNullString Then
        sbStatus.Panels("Message").Text = "Filter: " & SQLfilter
    Else
        sbStatus.Panels("Message").Text = vbNullString
    End If
    
    'Resize Columns based on content...
    ResizeWindow = 36
    For i = dgdList.LeftCol To dgdList.Columns.Count - 1
        Set col = dgdList.Columns(i)
        If col.Visible And col.Width > 0 Then
            dgdList.ClearSelCols
            Call ResizeColumn(col)
        End If
    Next i
    
    dgdList_Click
End Sub
Private Sub Form_Load()
    'fDebug = True
    
    Me.Icon = Forms(Forms.Count - 2).Icon
    dgdList.Top = 0
    dgdList.Left = 0
    dgdList.Width = Me.ScaleWidth
    dgdList.Height = Me.ScaleHeight
    dgdList.RowHeight = 192.189     'so I don't forget (this is for MS Sans Serif 8 point font)
End Sub
Private Sub Form_Resize()
    If Me.Width < MinWidth Then Me.Width = MinWidth
    If Me.Height < MinHeight Then Me.Height = MinHeight
    dgdList.Width = Me.ScaleWidth
    dgdList.Height = Me.ScaleHeight - sbStatus.Height
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    
    On Error Resume Next
    If Not RS Is Nothing Then
        If (RS.Status And adRecPendingChanges) = adRecPendingChanges Then RS.UpdateBatch
    End If
    
    'Save the column settings for the next display...
    For i = 0 To dgdList.Columns.Count - 1
        SaveSetting App.ProductName, Me.Caption & " Settings", dgdList.Columns(i).Caption & " Width", dgdList.Columns(i).Width
    Next
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
    dgdList_KeyUp vbKeyF2, 0
End Sub
Private Sub mnuListNew_Click()
    MsgBox "Sorry, New is not implemented yet...", vbExclamation, Me.Caption
End Sub
Private Sub ResizeColumn(col As Column)
    Dim ColumnFormat As New StdDataFormat
    Dim DataWidth As Long
    Dim ResizeWindow As Single
    Dim rsTemp As ADODB.Recordset
    Dim WidestData As Long
    
    Me.MousePointer = vbHourglass
    
    ResizeWindow = 36
    lblA.Caption = col.Caption
    WidestData = lblA.Width
    Set ColumnFormat = col.DataFormat
    If Not RS.BOF And Not RS.EOF Then
        Set rsTemp = RS.Clone(adLockReadOnly)
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
    End If
    Set ColumnFormat = Nothing
    col.Width = WidestData + (4 * ResizeWindow)
    If col.Width > dgdList.Width Then col.Width = col.Width - ResizeWindow
    
ExitSub:
    Me.MousePointer = vbDefault
End Sub
Private Sub sbStatus_PanelClick(ByVal Panel As MSComctlLib.Panel)
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
Private Sub UpdatePosition()
    sbStatus.Panels("Position").Text = "Record " & dgdList.BookMark & " of " & RS.RecordCount
    If fEditMode Then
        sbStatus.Panels("Status").Text = "Edit Mode (" & dgdList.col & ")"
    Else
        sbStatus.Panels("Status").Text = vbNullString
    End If
    If dgdList.EditActive Then sbStatus.Panels("Status").Text = sbStatus.Panels("Status").Text & " *"
End Sub
