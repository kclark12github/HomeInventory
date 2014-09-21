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
            TextSave        =   "7:49 PM"
            Key             =   "Time"
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid dgdList 
      Height          =   612
      Left            =   1260
      TabIndex        =   0
      Top             =   960
      Width           =   1212
      _ExtentX        =   2138
      _ExtentY        =   1080
      _Version        =   393216
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
   Begin VB.Menu mnuList 
      Caption         =   "&List"
      Visible         =   0   'False
      Begin VB.Menu mnuListMoveFirst 
         Caption         =   "Move &First"
      End
      Begin VB.Menu mnuListMoveLast 
         Caption         =   "Move &Last"
      End
      Begin VB.Menu mnuListSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListFilter 
         Caption         =   "&Filter"
      End
      Begin VB.Menu mnuListSort 
         Caption         =   "&Sort"
      End
   End
End
Attribute VB_Name = "frmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rsList As ADODB.Recordset
Private Key As String
Private SortDESC() As Boolean
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
End Sub
Private Sub dgdList_BeforeDelete(Cancel As Integer)
    If fDebug Then sbStatus.Panels("Message").Text = "BeforeDelete"
End Sub
Private Sub dgdList_BeforeInsert(Cancel As Integer)
    If fDebug Then sbStatus.Panels("Message").Text = "BeforeInsert"
End Sub
Private Sub dgdList_BeforeUpdate(Cancel As Integer)
    If fDebug Then sbStatus.Panels("Message").Text = "BeforeUpdate"
End Sub
Private Sub dgdList_Click()
    UpdatePosition
End Sub
Private Sub dgdList_ColEdit(ByVal ColIndex As Integer)
    sbStatus.Panels("Status").Text = "Edit Mode"
    If fDebug Then sbStatus.Panels("Message").Text = "ColEdit"
End Sub
Private Sub dgdList_HeadClick(ByVal ColIndex As Integer)
    If SortDESC(ColIndex) Then
        rsList.Sort = dgdList.Columns(ColIndex).Caption & " DESC"
    Else
        rsList.Sort = dgdList.Columns(ColIndex).Caption & " ASC"
    End If
    SortDESC(ColIndex) = Not SortDESC(ColIndex)
End Sub
Private Sub dgdList_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            If Not dgdList.EditActive Then dgdList.EditActive = False
            sbStatus.Panels("Status").Text = ""
        Case vbKeyF2
            dgdList.EditActive = True
            sbStatus.Panels("Status").Text = "Edit Mode"
    End Select
End Sub
Private Sub dgdList_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    UpdatePosition
End Sub
Private Sub dgdList_RowResize(Cancel As Integer)
    If fDebug Then sbStatus.Panels("Message").Text = "RowHeight: " & dgdList.RowHeight
End Sub
Private Sub Form_Activate()
    Dim i As Integer
    
    ReDim SortDESC(0 To dgdList.Columns.Count - 1)
    
    'Get the column settings for the display...
    For i = 0 To dgdList.Columns.Count - 1
        dgdList.Columns(i).Width = GetSetting(App.ProductName, Me.Caption & " Settings", dgdList.Columns(i).Caption & " Width", dgdList.Columns(i).Width)
    Next
    Me.Top = GetSetting(App.ProductName, Me.Caption & " Settings", "Form Top", Me.Top)
    Me.Left = GetSetting(App.ProductName, Me.Caption & " Settings", "Form Left", Me.Left)
    Me.Width = GetSetting(App.ProductName, Me.Caption & " Settings", "Form Width", Me.Width)
    Me.Height = GetSetting(App.ProductName, Me.Caption & " Settings", "Form Height", Me.Height)
        
    If rsList.Filter <> vbNullString Then
        sbStatus.Panels("Message").Text = "Filter: " & rsList.Filter
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
    If Me.Width < frmMain.MinWidth Then Me.Width = frmMain.MinWidth
    If Me.Height < frmMain.MinHeight Then Me.Height = frmMain.MinHeight
    dgdList.Width = Me.ScaleWidth
    dgdList.Height = Me.ScaleHeight - sbStatus.Height
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    
    rsList.UpdateBatch
    
    'Save the column settings for the next display...
    For i = 0 To dgdList.Columns.Count - 1
        SaveSetting App.ProductName, Me.Caption & " Settings", dgdList.Columns(i).Caption & " Width", dgdList.Columns(i).Width
    Next
    SaveSetting App.ProductName, Me.Caption & " Settings", "Form Top", Me.Top
    SaveSetting App.ProductName, Me.Caption & " Settings", "Form Left", Me.Left
    SaveSetting App.ProductName, Me.Caption & " Settings", "Form Width", Me.Width
    SaveSetting App.ProductName, Me.Caption & " Settings", "Form Height", Me.Height
End Sub
Private Sub sbStatus_PanelClick(ByVal Panel As MSComctlLib.Panel)
    Select Case UCase(Panel.Key)
        Case "TOP"
            rsList.MoveFirst
            UpdatePosition
        Case "BOTTOM"
            rsList.MoveLast
            UpdatePosition
        Case "FILTER"
            MsgBox "Filter button clicked..."
        Case Else
    End Select
End Sub
Private Sub UpdatePosition()
    sbStatus.Panels("Position").Text = "Record " & dgdList.Bookmark & " of " & rsList.RecordCount & "  "
End Sub
