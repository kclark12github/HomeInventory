VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{C2000000-FFFF-1100-8100-000000000001}#8.0#0"; "PVCURR.OCX"
Begin VB.Form frmBooks 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Book Inventory"
   ClientHeight    =   4380
   ClientLeft      =   30
   ClientTop       =   495
   ClientWidth     =   7530
   Icon            =   "frmBooks.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin PVCurrencyLib.PVCurrency pvcPrice 
      Bindings        =   "frmBooks.frx":0442
      Height          =   285
      Left            =   1464
      TabIndex        =   3
      Top             =   1440
      Width           =   975
      _Version        =   524288
      _ExtentX        =   1720
      _ExtentY        =   503
      _StockProps     =   253
      Text            =   "$0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      Alignment       =   2
      EditMode        =   0
      EditModeChange  =   0   'False
      Value           =   0
      ChangeColor     =   -1  'True
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   4125
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1270
            MinWidth        =   1270
            Key             =   "Position"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "Status"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7964
            Key             =   "Message"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   2
            Object.Width           =   1376
            MinWidth        =   1270
            TextSave        =   "10:51 PM"
            Key             =   "Time"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   372
      Left            =   6360
      TabIndex        =   10
      Top             =   3660
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   5340
      TabIndex        =   9
      Top             =   3660
      Width           =   972
   End
   Begin MSAdodcLib.Adodc adodcMain 
      Height          =   315
      Left            =   150
      Top             =   3300
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   556
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
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txtAlphaSort 
      Height          =   288
      Left            =   1464
      TabIndex        =   6
      Text            =   "AlphaSort"
      Top             =   2520
      Width           =   5832
   End
   Begin VB.TextBox txtInventoried 
      Height          =   288
      Left            =   3330
      TabIndex        =   8
      Text            =   "Inventoried"
      Top             =   2880
      Width           =   3195
   End
   Begin MSDataListLib.DataCombo dbcSubject 
      Height          =   315
      Left            =   1470
      TabIndex        =   4
      Top             =   1800
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   "Subject"
   End
   Begin MSDataListLib.DataCombo dbcAuthor 
      Height          =   315
      Left            =   1470
      TabIndex        =   1
      Top             =   720
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   "Author"
   End
   Begin VB.TextBox txtMisc 
      Height          =   288
      Left            =   1464
      TabIndex        =   5
      Text            =   "Misc"
      Top             =   2160
      Width           =   5895
   End
   Begin VB.TextBox txtISBN 
      Height          =   300
      Left            =   1464
      TabIndex        =   2
      Text            =   "ISBN"
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CheckBox chkCataloged 
      Alignment       =   1  'Right Justify
      Caption         =   "Cataloged"
      Height          =   192
      Left            =   480
      TabIndex        =   7
      Top             =   2940
      Width           =   1152
   End
   Begin VB.TextBox txtTitle 
      Height          =   300
      Left            =   1464
      TabIndex        =   0
      Text            =   "Title"
      Top             =   372
      Width           =   5892
   End
   Begin MSComctlLib.Toolbar tbMain 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlSmall"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Report"
            Object.ToolTipText     =   "Report"
            ImageKey        =   "Report"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SQL"
            Object.ToolTipText     =   "SQL Window"
            ImageKey        =   "SQL"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New record"
            ImageKey        =   "NewRecord"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Modify"
            Object.ToolTipText     =   "Modify record"
            ImageKey        =   "Modify"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete record"
            ImageKey        =   "Delete"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh data"
            ImageKey        =   "Refresh"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Filter"
            Object.ToolTipText     =   "Filter"
            ImageKey        =   "Filter"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Search"
            Object.ToolTipText     =   "Search"
            ImageKey        =   "Search"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "List"
            Object.ToolTipText     =   "List all records"
            ImageKey        =   "List"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList imlSmall 
      Left            =   240
      Top             =   60
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBooks.frx":044D
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBooks.frx":0769
            Key             =   "Warning"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBooks.frx":0A91
            Key             =   "List"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBooks.frx":0DB9
            Key             =   "xNew"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBooks.frx":356D
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBooks.frx":39C1
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBooks.frx":3E15
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBooks.frx":48E1
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBooks.frx":4C09
            Key             =   "Sort"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBooks.frx":505D
            Key             =   "SQL"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBooks.frx":54B1
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBooks.frx":5905
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBooks.frx":5D5D
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBooks.frx":5EB9
            Key             =   "Blank"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBooks.frx":6015
            Key             =   "NewRecord"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblIDlabel 
      AutoSize        =   -1  'True
      Caption         =   "ID:"
      Height          =   195
      Left            =   6750
      TabIndex        =   20
      Top             =   2927
      Width           =   195
   End
   Begin VB.Label lblDateInventoried 
      AutoSize        =   -1  'True
      Caption         =   "Date Inventoried:"
      Height          =   195
      Left            =   1995
      TabIndex        =   19
      Top             =   2925
      Width           =   1215
   End
   Begin VB.Label lblAlphaSort 
      AutoSize        =   -1  'True
      Caption         =   "AlphaSort:"
      Height          =   195
      Left            =   606
      TabIndex        =   18
      Top             =   2567
      Width           =   750
   End
   Begin VB.Label lblMisc 
      AutoSize        =   -1  'True
      Caption         =   "Miscellaneous:"
      Height          =   195
      Left            =   276
      TabIndex        =   17
      Top             =   2207
      Width           =   1080
   End
   Begin VB.Label lblSubject 
      AutoSize        =   -1  'True
      Caption         =   "Subject:"
      Height          =   195
      Left            =   786
      TabIndex        =   16
      Top             =   1845
      Width           =   570
   End
   Begin VB.Label lblPrice 
      AutoSize        =   -1  'True
      Caption         =   "Price:"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   195
      Left            =   951
      TabIndex        =   15
      Top             =   1487
      Width           =   405
   End
   Begin VB.Label lblISBN 
      AutoSize        =   -1  'True
      Caption         =   "ISBN:"
      Height          =   192
      Left            =   948
      TabIndex        =   14
      Top             =   1134
      Width           =   408
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Title:"
      Height          =   192
      Left            =   1008
      TabIndex        =   13
      Top             =   426
      Width           =   348
   End
   Begin VB.Label lblAuthor 
      AutoSize        =   -1  'True
      Caption         =   "Author:"
      Height          =   192
      Left            =   864
      TabIndex        =   12
      Top             =   766
      Width           =   492
   End
   Begin VB.Label lblID 
      AutoSize        =   -1  'True
      Caption         =   "lblID"
      Height          =   195
      Left            =   7035
      TabIndex        =   11
      Top             =   2927
      Width           =   330
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileReport 
         Caption         =   "&Report"
      End
      Begin VB.Menu mnuFileSQL 
         Caption         =   "&SQL"
      End
      Begin VB.Menu mnuFileSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuRecords 
      Caption         =   "&Records"
      Begin VB.Menu mnuRecordsNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuRecordsCopy 
         Caption         =   "&Copy/Append"
      End
      Begin VB.Menu mnuRecordsModify 
         Caption         =   "&Modify"
      End
      Begin VB.Menu mnuRecordsDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuRecordsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecordsRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuRecordsSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecordsFilter 
         Caption         =   "&Filter"
      End
      Begin VB.Menu mnuRecordsSearch 
         Caption         =   "&Search"
      End
      Begin VB.Menu mnuRecordsSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecordsList 
         Caption         =   "&List"
      End
   End
End
Attribute VB_Name = "frmBooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'frmBooks - frmBooks.frm
'   Book Inventory Form...
'   Copyright � 1999-2002, Ken Clark
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Description:
'   08/20/02    Started History;
'=================================================================================================================================
Option Explicit
Public FileDSN As String

Dim WithEvents vrsMain As ADODB.Recordset
Attribute vrsMain.VB_VarHelpID = -1
Dim WithEvents CurrencyFormat As StdDataFormat
Attribute CurrencyFormat.VB_VarHelpID = -1
Dim vrsAuthors As New ADODB.Recordset
Dim vrsSubjects As New ADODB.Recordset
Dim strDefaultAlphaSort As String
Private Sub cmdCancel_Click()
    CancelCommand Me, vrsMain
End Sub
Private Sub cmdOK_Click()
    Dim fCancel As Boolean
    Select Case mode
        Case modeAdd, modeModify
            If chkCataloged.Value = vbChecked Then
                fCancel = False
                Call dbcAuthor_Validate(fCancel)
                If fCancel Then GoTo ExitSub
                Call txtInventoried_Validate(fCancel)
                If fCancel Then GoTo ExitSub
                Call txtISBN_Validate(fCancel)
                If fCancel Then GoTo ExitSub
            End If
    End Select
    
    OKCommand Me, vrsMain

ExitSub:
    Exit Sub
End Sub
Private Sub Form_Activate()
    Dim Caption As String
    
    Me.Top = frmMain.saveTop + ((frmMain.Height - Me.Height) / 2)
    Me.Left = frmMain.saveLeft + ((frmMain.Width - Me.Width) / 2)
    If Not vrsMain.BOF And Not vrsMain.EOF Then Caption = "Reference #" & vrsMain.BookMark & ": " & vrsMain(SQLkey)
    UpdatePosition Me, Caption, vrsMain
End Sub
Private Sub Form_Load()
    'EstablishConnection adoConn
    Call ConnectByFileDSN(gstrFileDSN, vbNullString, vbNullString)
    
    Set vrsMain = New ADODB.Recordset
'    vrsMain.CursorLocation = adUseClient
    SQLmain = "Select * From [Books] Order By AlphaSort"
    SQLfilter = vbNullString
    SQLkey = "AlphaSort"
    'rsMain.Open SQLmain, adoConn, adOpenKeyset, adLockBatchOptimistic
    Call MakeVirtualRecordset(SQLmain, vrsMain)
    DBcollection.Add "vrsMain", vrsMain
    
'    rsAuthors.CursorLocation = adUseClient
'    rsAuthors.Open "Select Distinct Author From [Books] Order By Author", adoConn, adOpenStatic, adLockReadOnly
    Call MakeVirtualRecordset("Select Distinct Author From [Books] Order By Author", vrsAuthors)
    DBcollection.Add "vrsAuthors", vrsAuthors
    
'    rsSubjects.CursorLocation = adUseClient
'    rsSubjects.Open "Select Distinct Subject From [Books] Order By Subject", adoConn, adOpenStatic, adLockReadOnly
    Call MakeVirtualRecordset("Select Distinct Subject From [Books] Order By Subject", vrsSubjects)
    DBcollection.Add "vrsSubjects", vrsSubjects
    
    Set adodcMain.Recordset = vrsMain
    BindField lblID, "ID", vrsMain, "ID"
    BindField dbcAuthor, "Author", vrsMain, "Author", vrsAuthors, "Author", "Author"
    BindField txtTitle, "Title", vrsMain, "Title"
    BindField txtISBN, "ISBN", vrsMain, "ISBN"
    BindField pvcPrice, "Price", vrsMain, "Price"
    BindField txtAlphaSort, "AlphaSort", vrsMain, "AlphaSort"
    BindField dbcSubject, "Subject", vrsMain, "Subject", vrsSubjects, "Subject", "Subject"
    BindField txtMisc, "Misc", vrsMain, "Misc"
    BindField chkCataloged, "Cataloged", vrsMain, "Cataloged"
    BindField txtInventoried, "Inventoried", vrsMain, "Inventoried"

    ProtectFields Me
    mode = modeDisplay
    fTransaction = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim iMap As FieldMap
    Dim i As Integer
    
    For i = 1 To FieldMaps.Count
        Set iMap = FieldMaps(1)
        FieldMaps.Remove 1
        Set iMap = Nothing
    Next
    Cancel = CloseConnection(Me)
End Sub
Private Sub mnuRecordsCopy_Click()
    CopyCommand Me, vrsMain, SQLkey
End Sub
Private Sub mnuRecordsFilter_Click()
    FilterCommand Me, vrsMain, SQLkey
End Sub
Private Sub mnuRecordsDelete_Click()
    DeleteCommand Me, vrsMain
End Sub
Private Sub mnuRecordsList_Click()
    ListCommand Me, vrsMain
End Sub
Private Sub mnuRecordsModify_Click()
    ModifyCommand Me
    
    txtTitle.SetFocus
End Sub
Private Sub mnuRecordsNew_Click()
    NewCommand Me, vrsMain
    
    chkCataloged.Value = vbChecked
'    txtInventoried.Text = Format(Now(), fmtDate)
    strDefaultAlphaSort = vbNullString
    txtTitle.SetFocus
End Sub
Private Sub mnuRecordsRefresh_Click()
    RefreshCommand vrsMain, "ID"
End Sub
Private Sub mnuRecordsSearch_Click()
    SearchCommand Me, vrsMain, SQLkey
End Sub
Private Sub mnuFileExit_Click()
    Unload Me
End Sub
Private Sub mnuFileReport_Click()
    ReportCommand Me, vrsMain, App.Path & "\Reports\Books.rpt"
End Sub
Private Sub mnuFileSQL_Click()
    SQLCommand "Books"
End Sub
Private Sub pvcPrice_GotFocus()
    TextSelected
End Sub
Private Sub pvcPrice_GotFocusEvent()
    TextSelected
End Sub
Private Sub vrsMain_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Dim Caption As String
    
    If Not pRecordset.BOF And Not pRecordset.EOF Then Caption = "Reference #" & pRecordset.BookMark & ": " & pRecordset(SQLkey)
    UpdatePosition Me, Caption, pRecordset
End Sub
Private Sub tbMain_ButtonClick(ByVal Button As MSComCtlLib.Button)
    Select Case Button.Key
        Case "Report"
            mnuFileReport_Click
        Case "SQL"
            mnuFileSQL_Click
        Case "New"
            mnuRecordsNew_Click
        Case "Modify"
            mnuRecordsModify_Click
        Case "Delete"
            mnuRecordsDelete_Click
        Case "Refresh"
            mnuRecordsRefresh_Click
        Case "Filter"
            mnuRecordsFilter_Click
        Case "Search"
            mnuRecordsSearch_Click
        Case "List"
            mnuRecordsList_Click
    End Select
End Sub
'=================================================================================
Private Sub dbcAuthor_GotFocus()
    TextSelected
End Sub
Private Sub dbcAuthor_Validate(Cancel As Boolean)
    If Not dbcAuthor.Enabled Then Exit Sub
    If dbcAuthor.Text = vbNullString And chkCataloged.Value = vbChecked Then
        MsgBox "Author must be specified!", vbExclamation, Me.Caption
        dbcAuthor.SetFocus
        Cancel = True
    End If
    If dbcValidate(vrsMain("Author"), dbcAuthor) = 0 Then Cancel = True
    If rsAuthors.BookMark <> dbcAuthor.SelectedItem Then rsAuthors.BookMark = dbcAuthor.SelectedItem
End Sub
Private Sub dbcSubject_GotFocus()
    TextSelected
End Sub
Private Sub dbcSubject_Validate(Cancel As Boolean)
    If Not dbcSubject.Enabled Then Exit Sub
    'If dbcSubject.Text = vbnullstring Then
    '    MsgBox "Subject must be specified!", vbExclamation, Me.Caption
    '    dbcSubject.SetFocus
    '    Cancel = True
    'End If
    If dbcValidate(vrsMain("Subject"), dbcSubject) = 0 Then Cancel = True
    If rsSubjects.BookMark <> dbcSubject.SelectedItem Then rsSubjects.BookMark = dbcSubject.SelectedItem
End Sub
Private Function DefaultAlphaSort() As String
    Dim LastName As String
    Dim Title As String
    Dim iAnd As Integer
    Dim iAmpersand As Integer
    Dim iComma As Integer
    Dim iSemiColon As Integer
    Dim iSeparator As Integer
    Dim iWith As Integer
    Dim CountryCode As Long
    Dim PublisherCode As Long
    
    'Start with the Author's last name...
    LastName = dbcAuthor.Text
    iAnd = InStr(dbcAuthor.Text, " and ")
    iAmpersand = InStr(dbcAuthor.Text, " & ")
    iComma = InStr(dbcAuthor.Text, ",")
    iSemiColon = InStr(dbcAuthor.Text, ";")
    iWith = InStr(dbcAuthor.Text, " with ")
    
    If iComma > 0 Then
        'Assume the comma separates authors, and...
        iSeparator = iComma
    ElseIf iSemiColon > 0 Then
        'Assume the semicolon separates authors, and...
        iSeparator = iSemiColon
    ElseIf iAnd > 0 Then
        'Assume the "and" separates authors, and...
        iSeparator = iAnd
    ElseIf iAmpersand > 0 Then
        'Assume the "&" separates authors, and...
        iSeparator = iAmpersand
    ElseIf iWith > 0 Then
        'Assume the "with" separates authors, and...
        iSeparator = iWith
    End If
    
    If iSeparator > 0 Then
        '...take the first Author...
        LastName = Mid(LastName, 1, iSeparator - 1)
    End If
        
    'OK, we have a single person's name (theoretically)...
    'Grab the last word on the line and assume it's his last name...
    If InStr(LastName, " ") Then
        iSeparator = InStrRev(LastName, " ", Len(LastName))
        LastName = Mid(LastName, iSeparator + 1)
    End If
    
    'Check for "The" at the beginning of the title...
    Title = txtTitle.Text
    If Mid(UCase(Title), 1, 4) = "THE " Then
        Title = Mid(Title, 5) & ", " & Mid(Title, 1, 3)
    End If
    
    'Check ISBN to detect publishers of special series that
    'we'd rather sort by series name than author(s)...
    If Trim(txtISBN.Text) <> vbNullString Then
        CountryCode = ParseStr(txtISBN.Text, 1, "-")        'If we begin to get dups on PublisherCode, we'll have to start using the CountryCode, but not until then...
        PublisherCode = ParseStr(txtISBN.Text, 2, "-")      'PublisherCode
        Select Case PublisherCode
            Case 880588, 874023, 86184, 874023      'CountryCode: 1
                DefaultAlphaSort = "AIRTIME: " & UCase(Title)
            Case 361                                'CountryCode: 962
                DefaultAlphaSort = "CONCORD: <SeriesName> #nnnn; " & UCase(Title)
            Case 7603, 87398, 85045                 'CountryCode: 0
                DefaultAlphaSort = "MOTORBOOKS: " & UCase(LastName) & ": " & UCase(Title)
            Case 914845                             'CountryCode: 0
                DefaultAlphaSort = "MSPRESS: " & UCase(LastName) & ": " & UCase(Title)
            Case 57231, 55615                       'CountryCode: 1
                DefaultAlphaSort = "MSPRESS: " & UCase(LastName) & ": " & UCase(Title)
            Case 896522                             'CountryCode: 1
                DefaultAlphaSort = "NASA MISSION REPORTS: " & UCase(Title)
            Case 87938, 89141, 85045                'CountryCode: 0
                DefaultAlphaSort = "OSPREY: " & UCase(Title)
            Case 85532, 84176                       'CountryCode: 1
                DefaultAlphaSort = "OSPREY: <SeriesName> #nn; " & UCase(Title)
            Case 672                                'CountryCode: 0
                DefaultAlphaSort = "SAMS: " & UCase(Title)
            Case 944055                             'CountryCode: 0
                DefaultAlphaSort = "SCALE MODELING: FLOATING DRYDOCK: " & UCase(Title)
            Case 89024                              'CountryCode: 0
                DefaultAlphaSort = "SCALE MODELING: KALMBACH: " & UCase(Title)
            Case 7643, 88740                        'CountryCode: 0
                DefaultAlphaSort = "SCHIFFER MILITARY HISTORY: " & UCase(LastName & ": " & Title)
            Case 8094, 670, 939526                  'CountryCode: 0
                DefaultAlphaSort = "TIMELIFE: <SeriesName>: BOOK ##; " & UCase(Title)
            Case 85488                              'CountryCode: 1
                DefaultAlphaSort = "TRISERVICE: " & UCase(Title)
            Case 935696, 88038                      'CountryCode: 0
                DefaultAlphaSort = "TSR: #### " & UCase(Title)
            Case 89747                              'CountryCode: 0
                DefaultAlphaSort = "SQUADRON/SIGNAL: <SeriesName> NO. ##; " & UCase(Title)
            Case 930607                             'CountryCode: 1
                DefaultAlphaSort = "VERLINDEN: <SeriesName> NO. ##; " & UCase(Title)
            Case 70932                              'CountryCode: 90
                DefaultAlphaSort = "VERLINDEN: <SeriesName> NO. ##; " & UCase(Title)
            Case 9654829, 9710687                   'CountryCode: 0
                DefaultAlphaSort = "WARSHIP PICTORIAL SERIES #nn; " & UCase(Title)
            Case 933126, 929521                     'CountryCode: 0
                DefaultAlphaSort = "WARSHIP SERIES #n; <ShipDesignation>; U.S.S. <ShipName>"
            Case 929521                             'CountryCode: 0
                DefaultAlphaSort = "WARSHIP'S DATA #n; <ShipDesignation>; U.S.S. <ShipName>"
            Case 57510                              'CountryCode: 1
                DefaultAlphaSort = "WARSHIP'S DATA #n; <ShipDesignation>; U.S.S. <ShipName>"
            Case Else
                Select Case dbcSubject.Text
                    Case "Science Fiction; Star Trek"
                        DefaultAlphaSort = UCase("STAR TREK: " & LastName & ": " & Title)
                    Case "Science Fiction; Star Wars"
                        DefaultAlphaSort = UCase("STAR WARS: " & LastName & ": " & Title)
                    Case Else
                        DefaultAlphaSort = UCase(LastName & ": " & Title)
                End Select
        End Select
    Else
        DefaultAlphaSort = UCase(LastName & ": " & Title)
    End If
End Function
Private Sub txtAlphaSort_GotFocus()
    TextSelected
End Sub
Private Sub txtAlphaSort_KeyPress(KeyAscii As Integer)
    KeyPressUcase KeyAscii
End Sub
Private Sub txtAlphaSort_Validate(Cancel As Boolean)
    If Not txtAlphaSort.Enabled Then Exit Sub
    If txtAlphaSort.Text = vbNullString Then
        txtAlphaSort.Text = DefaultAlphaSort
        'MsgBox "AlphaSort must be specified!", vbExclamation, Me.Caption
        'txtAlphaSort.SetFocus
        'Cancel = True
    Else
        If ParseStr(txtAlphaSort.Text, 1, ":") <> ParseStr(DefaultAlphaSort, 1, ":") Then
            Call MsgBox("AlphaSort doesn't match the configured value for this title/author/ISBN... Expected """ & ParseStr(DefaultAlphaSort, 1, ":") & ": ...""", vbInformation)
        End If
    End If
    txtAlphaSort.Text = UCase(txtAlphaSort.Text)
End Sub
Private Sub txtInventoried_GotFocus()
    TextSelected
End Sub
Private Sub txtInventoried_Validate(Cancel As Boolean)
    On Error Resume Next
    txtInventoried.Text = Format(txtInventoried.Text, fmtDate)
    If txtInventoried.Text = vbNullString Then
        If chkCataloged.Value = vbChecked Then txtInventoried.Text = Format(Now(), fmtDate) Else Exit Sub
    End If
    If Not IsDate(txtInventoried.Text) Then
        MsgBox "Invalid date format", vbExclamation
        Cancel = True
        Exit Sub
    End If
End Sub
Private Sub txtISBN_GotFocus()
    TextSelected
End Sub
Private Sub txtISBN_KeyPress(KeyAscii As Integer)
    KeyPressUcase KeyAscii
End Sub
Private Sub txtISBN_Validate(Cancel As Boolean)
    If Not txtISBN.Enabled Then Exit Sub
    If txtISBN.Text = vbNullString And chkCataloged.Value = vbChecked Then
        MsgBox "ISBN should be specified!", vbInformation, Me.Caption
    Else
        'txtISBN.Text = FormatISBN(txtISBN.Text)
        Select Case Left(txtISBN.Text, 1)
            Case "0" To "9", "M"        'Some forms of ISBN (EAN) may begin with "M"...
                txtISBN.Text = CheckISBN(txtISBN.Text)
            Case Else
        End Select
    End If
End Sub
Private Sub txtMisc_GotFocus()
    TextSelected
End Sub
Private Sub txtTitle_GotFocus()
    TextSelected
End Sub
Private Sub txtTitle_Validate(Cancel As Boolean)
    If Not txtTitle.Enabled Then Exit Sub
    If txtTitle.Text = vbNullString Then
        MsgBox "Title must be specified!", vbExclamation, Me.Caption
        txtTitle.SetFocus
        Cancel = True
    End If
End Sub
