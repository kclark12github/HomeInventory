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
            Object.Width           =   3175
            Key             =   "Message"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   2
            Object.Width           =   1270
            MinWidth        =   1270
            TextSave        =   "12:10 PM"
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
    sbStatus.Panels("Status").Text = ""
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
                GoTo ExitSub
            End If
        End If
    Next iCol
    
ExitSub:
    Me.MousePointer = vbDefault
End Sub
Private Sub dgdList_HeadClick(ByVal ColIndex As Integer)
    If RS.BOF And RS.EOF Then Exit Sub
    RS.Sort = vbNullString
    If SortDESC(ColIndex) Then
        RS.Sort = dgdList.Columns(ColIndex).Caption & " DESC"
    Else
        RS.Sort = dgdList.Columns(ColIndex).Caption & " ASC"
    End If
    SortDESC(ColIndex) = Not SortDESC(ColIndex)
End Sub
Private Sub dgdList_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim col As Column
    Dim i As Integer
    
    Select Case KeyCode
        Case vbKeyEscape
            If RS.EditMode = adEditInProgress Then RS.Update
            fEditMode = False
            EditRow = -1
            
            If Not dgdList.EditActive Then dgdList.EditActive = False
            dgdList.AllowUpdate = False
            sbStatus.Panels("Status").Text = ""
        Case vbKeyF2
            fEditMode = True
            EditRow = dgdList.Row
            dgdList.AllowUpdate = True
            
            'Position to the first non-hidden column...
            For Each col In dgdList.Columns
                If col.Visible Then
                    dgdList.col = col.ColIndex
                    Exit For
                End If
            Next col
    
            dgdList.EditActive = True
            sbStatus.Panels("Status").Text = "Edit Mode"
            For i = 0 To dgdList.SelBookmarks.Count - 1
                dgdList.SelBookmarks.Remove 0
            Next i
            dgdList.SelBookmarks.Add dgdList.Bookmark
    End Select
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
        dgdList_KeyUp vbKeyEscape, 0
    End If
    
    If Not fEditMode Then
        For i = 0 To dgdList.SelBookmarks.Count - 1
            dgdList.SelBookmarks.Remove 0
        Next i
        dgdList.SelBookmarks.Add dgdList.Bookmark
    
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
    dgdList.AllowUpdate = fAllowEditMode
    dgdList.AllowAddNew = fAllowEditMode
    Set dgdList.DataSource = RS
    ReDim SortDESC(0 To dgdList.Columns.Count - 1)
    
    For Each fld In RS.Fields
        Set col = dgdList.Columns(fld.Name)
        col.Visible = True
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
        
    If RS.Filter <> vbNullString And RS.Filter <> 0 Then
        sbStatus.Panels("Message").Text = "Filter: " & RS.Filter
    Else
        sbStatus.Panels("Message").Text = vbNullString
    End If
    dgdList_Click
End Sub
Private Sub Form_Load()
    'fDebug = True
    
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
    
    If (RS.Status And adRecPendingChanges) = adRecPendingChanges Then RS.UpdateBatch
    
    'Save the column settings for the next display...
    For i = 0 To dgdList.Columns.Count - 1
        SaveSetting App.ProductName, Me.Caption & " Settings", dgdList.Columns(i).Caption & " Width", dgdList.Columns(i).Width
    Next
    SaveSetting App.ProductName, Me.Caption & " Settings", "Form Top", Me.Top
    SaveSetting App.ProductName, Me.Caption & " Settings", "Form Left", Me.Left
    SaveSetting App.ProductName, Me.Caption & " Settings", "Form Width", Me.Width
    SaveSetting App.ProductName, Me.Caption & " Settings", "Form Height", Me.Height
End Sub
Private Sub mnuListEdit_Click()
    dgdList_KeyUp vbKeyF2, 0
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
    sbStatus.Panels("Position").Text = "Record " & dgdList.Bookmark & " of " & RS.RecordCount
End Sub
