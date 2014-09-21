VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmVideoResearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Video Research"
   ClientHeight    =   2988
   ClientLeft      =   36
   ClientTop       =   492
   ClientWidth     =   7524
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2988
   ScaleWidth      =   7524
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtSort 
      Height          =   288
      Left            =   1530
      TabIndex        =   4
      Text            =   "Sort"
      Top             =   1260
      Width           =   5892
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   372
      Left            =   6486
      TabIndex        =   7
      Top             =   2520
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   5466
      TabIndex        =   6
      Top             =   2520
      Width           =   972
   End
   Begin MSAdodcLib.Adodc adodcHobby 
      Height          =   312
      Left            =   276
      Top             =   2040
      Width           =   7152
      _ExtentX        =   12615
      _ExtentY        =   550
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txtInventoried 
      Height          =   288
      Left            =   1536
      TabIndex        =   5
      Text            =   "Inventoried"
      Top             =   1560
      Width           =   1812
   End
   Begin MSDataListLib.DataCombo dbcDistributor 
      Height          =   288
      Left            =   1530
      TabIndex        =   3
      Top             =   972
      Width           =   5892
      _ExtentX        =   10393
      _ExtentY        =   508
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   "Distributor"
   End
   Begin VB.TextBox txtCost 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      Height          =   288
      Left            =   5934
      TabIndex        =   2
      Top             =   660
      Width           =   972
   End
   Begin VB.TextBox txtTitle 
      Height          =   288
      Left            =   1530
      TabIndex        =   0
      Text            =   "Title"
      Top             =   372
      Width           =   5892
   End
   Begin MSComctlLib.ImageList imlSmall 
      Left            =   492
      Top             =   2520
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVideoResearch.frx":0000
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVideoResearch.frx":031C
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVideoResearch.frx":0638
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVideoResearch.frx":0A8C
            Key             =   "List"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVideoResearch.frx":1558
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVideoResearch.frx":2224
            Key             =   "New2"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVideoResearch.frx":2CF0
            Key             =   "Delete2"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVideoResearch.frx":37BC
            Key             =   "Modify2"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVideoResearch.frx":4288
            Key             =   "New"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVideoResearch.frx":4D54
            Key             =   "Report"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlLarge 
      Left            =   72
      Top             =   2520
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVideoResearch.frx":51A8
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVideoResearch.frx":55FC
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVideoResearch.frx":5918
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVideoResearch.frx":5C34
            Key             =   "New2"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVideoResearch.frx":6700
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVideoResearch.frx":73CC
            Key             =   "List"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVideoResearch.frx":7E98
            Key             =   "Delete2"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVideoResearch.frx":8964
            Key             =   "Modify2"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVideoResearch.frx":9430
            Key             =   "New"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVideoResearch.frx":9EFC
            Key             =   "Delete"
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo dbcSubject 
      Height          =   288
      Left            =   1530
      TabIndex        =   1
      Top             =   660
      Width           =   3252
      _ExtentX        =   5736
      _ExtentY        =   508
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   "Subject"
   End
   Begin MSComctlLib.Toolbar tbHobby 
      Align           =   1  'Align Top
      Height          =   288
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   7524
      _ExtentX        =   13272
      _ExtentY        =   508
      ButtonWidth     =   1439
      ButtonHeight    =   466
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlSmall"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "List"
            Key             =   "List"
            Object.ToolTipText     =   "List all records"
            ImageKey        =   "List"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            Key             =   "New"
            Object.ToolTipText     =   "New record"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Modify"
            Key             =   "Modify"
            Object.ToolTipText     =   "Modify record"
            ImageKey        =   "Modify"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete record"
            ImageKey        =   "Delete"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Report"
            Key             =   "Report"
            Object.ToolTipText     =   "Report"
            ImageKey        =   "Report"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblSort 
      AutoSize        =   -1  'True
      Caption         =   "Sort:"
      Height          =   192
      Left            =   984
      TabIndex        =   15
      Top             =   1308
      Width           =   324
   End
   Begin VB.Label lblSubject 
      AutoSize        =   -1  'True
      Caption         =   "Subject:"
      Height          =   192
      Left            =   846
      TabIndex        =   14
      Top             =   720
      Width           =   576
   End
   Begin VB.Label lblIDlabel 
      AutoSize        =   -1  'True
      Caption         =   "ID:"
      Height          =   192
      Left            =   6816
      TabIndex        =   13
      Top             =   1620
      Width           =   192
   End
   Begin VB.Label lblDateInventoried 
      AutoSize        =   -1  'True
      Caption         =   "Date Inventoried:"
      Height          =   192
      Left            =   216
      TabIndex        =   12
      Top             =   1608
      Width           =   1212
   End
   Begin VB.Label lblCost 
      AutoSize        =   -1  'True
      Caption         =   "Cost:"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   192
      Left            =   5400
      TabIndex        =   11
      Top             =   720
      Width           =   360
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Title:"
      Height          =   192
      Left            =   1074
      TabIndex        =   10
      Top             =   420
      Width           =   348
   End
   Begin VB.Label lblDistributor 
      AutoSize        =   -1  'True
      Caption         =   "Distributor:"
      Height          =   192
      Left            =   666
      TabIndex        =   9
      Top             =   1020
      Width           =   756
   End
   Begin VB.Label lblID 
      AutoSize        =   -1  'True
      Caption         =   "lblID"
      Height          =   192
      Left            =   7104
      TabIndex        =   8
      Top             =   1620
      Width           =   324
   End
   Begin VB.Menu mnuAction 
      Caption         =   "&Action"
      Begin VB.Menu mnuActionList 
         Caption         =   "&List"
      End
      Begin VB.Menu mnuActionSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActionNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuActionModify 
         Caption         =   "&Modify"
      End
      Begin VB.Menu mnuActionDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuActionSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActionReport 
         Caption         =   "&Report"
      End
   End
End
Attribute VB_Name = "frmVideoResearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoConn As ADODB.Connection
Dim WithEvents rsVideoResearch As ADODB.Recordset
Attribute rsVideoResearch.VB_VarHelpID = -1
Dim rsDistributors As New ADODB.Recordset
Dim rsSubjects As New ADODB.Recordset
Dim strDefaultSort As String
Dim mode As ActionMode
Dim fTransaction As Boolean
Dim DBinfo As DataBaseInfo
Private Sub cmdCancel_Click()
    Select Case mode
        Case modeDisplay
            Unload Me
        Case modeAdd, modeModify
            rsVideoResearch.CancelUpdate
            adoConn.RollbackTrans
            fTransaction = False
            frmMain.ProtectFields Me
            mode = modeDisplay
            adodcHobby.Enabled = True
    End Select
End Sub
Private Sub cmdOK_Click()
    Dim SaveBookmark As String
    
    Select Case mode
        Case modeDisplay
            Unload Me
        Case modeAdd, modeModify
            'Why we need to do this is buggy...
            rsVideoResearch("Distributor") = dbcDistributor.BoundText
            rsVideoResearch("Subject") = dbcSubject.BoundText
            rsVideoResearch.UpdateBatch
            adoConn.CommitTrans
            fTransaction = False
            frmMain.ProtectFields Me
            mode = modeDisplay
            adodcHobby.Enabled = True
            
            SaveBookmark = rsVideoResearch("Sort")
            rsVideoResearch.Requery
            rsVideoResearch.Find "Sort='" & SQLQuote(SaveBookmark) & "'"
            rsDistributors.Requery
            rsSubjects.Requery
    End Select
End Sub
Private Sub dbcDistributor_GotFocus()
    TextSelected
End Sub
Private Sub dbcDistributor_Validate(Cancel As Boolean)
    If dbcDistributor.Text = "" Then
        MsgBox "Distributor must be specified!", vbExclamation, Me.Caption
        dbcDistributor.SetFocus
        Cancel = True
    End If
    If rsDistributors.Bookmark <> dbcDistributor.SelectedItem Then rsDistributors.Bookmark = dbcDistributor.SelectedItem
End Sub
Private Sub dbcSubject_GotFocus()
    TextSelected
End Sub
Private Sub dbcSubject_Validate(Cancel As Boolean)
    If dbcSubject.Text = "" Then
        MsgBox "Subject must be specified!", vbExclamation, Me.Caption
        dbcSubject.SetFocus
        Cancel = True
    End If
    If rsSubjects.Bookmark <> dbcSubject.SelectedItem Then rsSubjects.Bookmark = dbcSubject.SelectedItem
End Sub
Private Sub Form_Load()
    Set adoConn = New ADODB.Connection
    Set rsVideoResearch = New ADODB.Recordset
    Set DBinfo = frmMain.DBcollection("Hobby")
    With DBinfo
        adoConn.Provider = .Provider
        adoConn.CommandTimeout = 60
        adoConn.ConnectionTimeout = 60
        adoConn.Open .PathName, .UserName, .Password
    End With
    rsVideoResearch.CursorLocation = adUseClient
    rsVideoResearch.Open "select * from [Video Research] order by Sort", adoConn, adOpenKeyset, adLockBatchOptimistic
    
    rsDistributors.CursorLocation = adUseClient
    rsDistributors.Open "select distinct Distributor from [Video Research] order by Distributor", adoConn, adOpenStatic, adLockReadOnly
    
    rsSubjects.CursorLocation = adUseClient
    rsSubjects.Open "select distinct Subject from [Video Research] order by Subject", adoConn, adOpenStatic, adLockReadOnly
    
    Set adodcHobby.Recordset = rsVideoResearch
    frmMain.BindField lblID, "ID", rsVideoResearch
    frmMain.BindField dbcDistributor, "Distributor", rsVideoResearch, rsDistributors, "Distributor", "Distributor"
    frmMain.BindField txtTitle, "Title", rsVideoResearch
    frmMain.BindField txtCost, "Cost", rsVideoResearch
    frmMain.BindField dbcSubject, "Subject", rsVideoResearch, rsSubjects, "Subject", "Subject"
    frmMain.BindField txtSort, "Sort", rsVideoResearch
    frmMain.BindField txtInventoried, "DateInventoried", rsVideoResearch

    frmMain.ProtectFields Me
    mode = modeDisplay
    fTransaction = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If fTransaction Then
        MsgBox "Please complete the current operation before closing the window.", vbExclamation, Me.Caption
        Cancel = 1
        Exit Sub
    End If
    
    If Not rsVideoResearch.EOF Then
        If rsVideoResearch.EditMode <> adEditNone Then rsVideoResearch.CancelUpdate
    End If
    If (rsVideoResearch.State And adStateOpen) = adStateOpen Then rsVideoResearch.Close
    Set rsVideoResearch = Nothing
    rsDistributors.Close
    Set rsDistributors = Nothing
    rsSubjects.Close
    Set rsSubjects = Nothing
    
    On Error Resume Next
    adoConn.Close
    If Err.Number = 3246 Then
        adoConn.RollbackTrans
        fTransaction = False
        adoConn.Close
    End If
    Set adoConn = Nothing
End Sub
Private Sub mnuActionList_Click()
    Dim frm As Form
    Dim CurrencyFormat As New StdDataFormat
    Dim Col As Column
    
    CurrencyFormat.Format = "Currency"
    
    Load frmList
    frmList.Caption = Me.Caption & " List"
    If frmMain.Width > Me.Width And frmMain.Height > Me.Height Then
        Set frm = frmMain
    Else
        Set frm = Me
    End If
    frmList.Top = frm.Top
    frmList.Left = frm.Left
    frmList.Width = frm.Width
    frmList.Height = frm.Height
    
    Set frmList.rsList = rsVideoResearch
    Set frmList.mnuList = mnuAction
    Set frmList.dgdList.DataSource = frmList.rsList
    Set frmList.dgdList.Columns("Cost").DataFormat = CurrencyFormat
    For Each Col In frmList.dgdList.Columns
        Col.Alignment = dbgGeneral
    Next Col
    
    adoConn.BeginTrans
    fTransaction = True
    frmList.Show vbModal
    adoConn.CommitTrans
    fTransaction = False
End Sub
Private Sub mnuActionNew_Click()
    mode = modeAdd
    frmMain.OpenFields Me
    adodcHobby.Enabled = False
    rsVideoResearch.AddNew
    adoConn.BeginTrans
    fTransaction = True
    
    txtInventoried.Text = Format(Now(), "mm/dd/yyyy hh:nn AMPM")
    strDefaultSort = ""
    txtTitle.SetFocus
End Sub
Private Sub mnuActionDelete_Click()
    mode = modeDelete
    If MsgBox("Are you sure you want to permanently delete this record...?", vbYesNo, Me.Caption) = vbYes Then
        rsVideoResearch.Delete
        rsVideoResearch.MoveNext
        If rsVideoResearch.EOF Then rsVideoResearch.MoveLast
    End If
    mode = modeDisplay
End Sub
Private Sub mnuActionModify_Click()
    mode = modeModify
    frmMain.OpenFields Me
    adodcHobby.Enabled = False
    adoConn.BeginTrans
    fTransaction = True
    
    txtTitle.SetFocus
End Sub
Private Sub mnuActionReport_Click()
    'Dim Report As New scrHobbyReport
    
    'Report.Database.SetDataSource rsVideoResearch, 3, 1
    'Set frmMain.rdcReport = Report
    'Set frmMain.frmReport = Me
    
    'frmViewReport.Show vbModal
    
    'Set Report = Nothing
End Sub
Private Sub rsVideoResearch_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Dim Caption As String
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    If rsVideoResearch.BOF And rsVideoResearch.EOF Then
        Caption = "No Records"
    ElseIf rsVideoResearch.EOF Then
        Caption = "EOF"
    ElseIf rsVideoResearch.BOF Then
        Caption = "BOF"
    Else
        Caption = "Reference #" & rsVideoResearch.Bookmark & ": " & rsVideoResearch("Sort")
        
        i = InStr(Caption, "&")
        If i > 0 Then Caption = Left(Caption, i) & "&" & Mid(Caption, i + 1)
    End If
    
    adodcHobby.Caption = Caption
    Exit Sub

ErrorHandler:
    MsgBox Err.Description & " (Error " & Err.Number & ")", vbExclamation, Me.Caption
    Resume Next
End Sub
Private Sub tbHobby_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "List"
            mnuActionList_Click
        Case "New"
            mnuActionNew_Click
        Case "Modify"
            mnuActionModify_Click
        Case "Delete"
            mnuActionDelete_Click
        Case "Report"
            mnuActionReport_Click
    End Select
End Sub
Private Sub txtInventoried_GotFocus()
    TextSelected
End Sub
Private Sub txtTitle_GotFocus()
    TextSelected
End Sub
Private Sub txtTitle_Validate(Cancel As Boolean)
    If txtTitle.Text = "" Then
        MsgBox "Title must be specified!", vbExclamation, Me.Caption
        txtTitle.SetFocus
        Cancel = True
    End If
End Sub
Private Sub txtCost_GotFocus()
    TextSelected
End Sub
Private Sub txtCost_KeyPress(KeyAscii As Integer)
    If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then
        If KeyAscii <> Asc(".") Then
            KeyAscii = 0    'Cancel the character.
            Beep            'Sound error signal.
        End If
    End If
End Sub
Private Sub txtCost_Validate(Cancel As Boolean)
    If txtCost.Text = vbNullString Then txtCost.Text = Format(0, "Currency")
    If Not IsNumeric(txtCost.Text) Then
        MsgBox "Invalid Cost entered.", vbExclamation, Me.Caption
        TextSelected
        Cancel = True
    End If
End Sub
Private Sub txtSort_GotFocus()
    TextSelected
    If txtSort.Text = "" Then txtSort.Text = dbcSubject.Text & ": " & txtTitle.Text
End Sub
Private Sub txtSort_Validate(Cancel As Boolean)
    If txtSort.Text = "" Then
        MsgBox "Sort should be specified!", vbExclamation, Me.Caption
        txtSort.SetFocus
        'Cancel = True
    End If
End Sub
