VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmMusic 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Music Inventory"
   ClientHeight    =   4110
   ClientLeft      =   30
   ClientTop       =   495
   ClientWidth     =   7530
   Icon            =   "frmMusic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   22
      Top             =   3855
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
            TextSave        =   "12:26 PM"
            Key             =   "Time"
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox chkLP 
      Alignment       =   1  'Right Justify
      Caption         =   "LP:"
      Height          =   192
      Left            =   6420
      TabIndex        =   7
      Top             =   1741
      Width           =   615
   End
   Begin VB.CheckBox chkCS 
      Alignment       =   1  'Right Justify
      Caption         =   "Cassette:"
      Height          =   192
      Left            =   4980
      TabIndex        =   6
      Top             =   1741
      Width           =   1095
   End
   Begin VB.CheckBox chkCD 
      Alignment       =   1  'Right Justify
      Caption         =   "CD:"
      Height          =   192
      Left            =   3840
      TabIndex        =   5
      Top             =   1741
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   372
      Left            =   6384
      TabIndex        =   12
      Top             =   3420
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   5364
      TabIndex        =   11
      Top             =   3420
      Width           =   972
   End
   Begin MSAdodcLib.Adodc adodcMain 
      Height          =   330
      Left            =   180
      Top             =   2940
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   582
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
      Height          =   300
      Left            =   1530
      TabIndex        =   8
      Text            =   "AlphaSort"
      Top             =   2100
      Width           =   5832
   End
   Begin VB.TextBox txtInventoried 
      Height          =   300
      Left            =   1530
      TabIndex        =   9
      Text            =   "Inventoried"
      Top             =   2520
      Width           =   2835
   End
   Begin MSDataListLib.DataCombo dbcType 
      Height          =   315
      Left            =   1530
      TabIndex        =   4
      Top             =   1680
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   "Type"
   End
   Begin MSDataListLib.DataCombo dbcArtist 
      Height          =   315
      Left            =   1530
      TabIndex        =   0
      Top             =   420
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   "Artist"
   End
   Begin VB.TextBox txtPrice 
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
      Height          =   300
      Left            =   5250
      TabIndex        =   3
      Top             =   1260
      Width           =   1815
   End
   Begin VB.TextBox txtYear 
      Height          =   300
      Left            =   1530
      TabIndex        =   2
      Text            =   "Year"
      Top             =   1260
      Width           =   972
   End
   Begin VB.CheckBox chkInventoried 
      Alignment       =   1  'Right Justify
      Caption         =   "Inventoried"
      Height          =   192
      Left            =   5010
      TabIndex        =   10
      Top             =   2574
      Width           =   1395
   End
   Begin VB.TextBox txtTitle 
      Height          =   300
      Left            =   1530
      TabIndex        =   1
      Text            =   "Title"
      Top             =   840
      Width           =   5892
   End
   Begin MSComctlLib.Toolbar tbMain 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   23
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
            Picture         =   "frmMusic.frx":030A
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMusic.frx":0626
            Key             =   "Warning"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMusic.frx":094E
            Key             =   "List"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMusic.frx":0C76
            Key             =   "xNew"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMusic.frx":342A
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMusic.frx":387E
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMusic.frx":3CD2
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMusic.frx":479E
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMusic.frx":4AC6
            Key             =   "Sort"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMusic.frx":4F1A
            Key             =   "SQL"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMusic.frx":536E
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMusic.frx":57C2
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMusic.frx":5C1A
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMusic.frx":5D76
            Key             =   "Blank"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMusic.frx":5ED2
            Key             =   "NewRecord"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblIDlabel 
      AutoSize        =   -1  'True
      Caption         =   "ID:"
      Height          =   195
      Left            =   6750
      TabIndex        =   21
      Top             =   2573
      Width           =   195
   End
   Begin VB.Label lblDateInventoried 
      AutoSize        =   -1  'True
      Caption         =   "Date Inventoried:"
      Height          =   195
      Left            =   201
      TabIndex        =   20
      Top             =   2573
      Width           =   1215
   End
   Begin VB.Label lblAlphaSort 
      AutoSize        =   -1  'True
      Caption         =   "AlphaSort:"
      Height          =   195
      Left            =   666
      TabIndex        =   19
      Top             =   2153
      Width           =   750
   End
   Begin VB.Label lblType 
      AutoSize        =   -1  'True
      Caption         =   "Type:"
      Height          =   195
      Left            =   996
      TabIndex        =   18
      Top             =   1740
      Width           =   420
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
      Left            =   4725
      TabIndex        =   17
      Top             =   1313
      Width           =   405
   End
   Begin VB.Label lblYear 
      AutoSize        =   -1  'True
      Caption         =   "Year:"
      Height          =   195
      Left            =   1026
      TabIndex        =   16
      Top             =   1313
      Width           =   390
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Title:"
      Height          =   195
      Left            =   1071
      TabIndex        =   15
      Top             =   893
      Width           =   345
   End
   Begin VB.Label lblArtist 
      AutoSize        =   -1  'True
      Caption         =   "Artist:"
      Height          =   192
      Left            =   1032
      TabIndex        =   14
      Top             =   481
      Width           =   384
   End
   Begin VB.Label lblID 
      AutoSize        =   -1  'True
      Caption         =   "lblID"
      Height          =   195
      Left            =   7035
      TabIndex        =   13
      Top             =   2573
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
Attribute VB_Name = "frmMusic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents rsMain As ADODB.Recordset
Attribute rsMain.VB_VarHelpID = -1
Dim rsArtists As New ADODB.Recordset
Dim rsTypes As New ADODB.Recordset
Dim strDefaultAlphaSort As String
Private Sub cmdCancel_Click()
    CancelCommand Me, rsMain
End Sub
Private Sub cmdOK_Click()
    OKCommand Me, rsMain
End Sub
Private Sub Form_Activate()
    Me.Top = frmMain.saveTop + ((frmMain.Height - Me.Height) / 2)
    Me.Left = frmMain.saveLeft + ((frmMain.Width - Me.Width) / 2)
End Sub
Private Sub Form_Load()
    EstablishConnection adoConn
    
    Set rsMain = New ADODB.Recordset
    rsMain.CursorLocation = adUseClient
    SQLmain = "select * from [Music] order by AlphaSort"
    SQLfilter = vbNullString
    SQLkey = "AlphaSort"
    rsMain.Open SQLmain, adoConn, adOpenKeyset, adLockBatchOptimistic
    DBcollection.Add "rsMain", rsMain
    
    rsArtists.CursorLocation = adUseClient
    rsArtists.Open "select distinct Artist from [Music] order by Artist", adoConn, adOpenStatic, adLockReadOnly
    DBcollection.Add "rsArtists", rsArtists
    
    rsTypes.CursorLocation = adUseClient
    rsTypes.Open "select distinct Type from [Music] order by Type", adoConn, adOpenStatic, adLockReadOnly
    DBcollection.Add "rsTypes", rsTypes
    
    Set adodcMain.Recordset = rsMain
    BindField lblID, "ID", rsMain, "ID"
    BindField dbcArtist, "Artist", rsMain, "Artist", rsArtists, "Artist", "Artist"
    BindField txtTitle, "Title", rsMain, "Title"
    BindField txtYear, "Year", rsMain, "Year"
    BindField txtPrice, "Price", rsMain, "Price"
    BindField chkCD, "CD", rsMain, "CD"
    BindField chkCS, "CS", rsMain, "CS"
    BindField chkLP, "LP", rsMain, "LP"
    BindField txtAlphaSort, "AlphaSort", rsMain, "AlphaSort"
    BindField dbcType, "Type", rsMain, "Type", rsTypes, "Type", "Type"
    BindField chkInventoried, "Inventoried", rsMain, "Inventoried"
    BindField txtInventoried, "DateInventoried", rsMain, "Date Inventoried"

    ProtectFields Me
    mode = modeDisplay
    fTransaction = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Cancel = CloseConnection(Me)
End Sub
Private Sub mnuRecordsFilter_Click()
    FilterCommand Me, rsMain, SQLkey
End Sub
Private Sub mnuRecordsDelete_Click()
    DeleteCommand Me, rsMain
End Sub
Private Sub mnuRecordsList_Click()
    ListCommand Me, rsMain
End Sub
Private Sub mnuRecordsModify_Click()
    ModifyCommand Me
    
    dbcArtist.SetFocus
End Sub
Private Sub mnuRecordsNew_Click()
    NewCommand Me, rsMain
    
    txtInventoried.Text = Format(Now(), fmtDate)
    chkInventoried.Value = vbChecked
    chkCD.Value = vbUnchecked
    chkCS.Value = vbUnchecked
    chkLP.Value = vbUnchecked
    strDefaultAlphaSort = vbNullString
    dbcArtist.SetFocus
End Sub
Private Sub mnuRecordsRefresh_Click()
    RefreshCommand rsMain, "ID"
End Sub
Private Sub mnuRecordsSearch_Click()
    SearchCommand Me, rsMain, SQLkey
End Sub
Private Sub mnuFileExit_Click()
    Unload Me
End Sub
Private Sub mnuFileReport_Click()
    ReportCommand Me, rsMain, App.Path & "\Reports\Music.rpt"
End Sub
Private Sub mnuFileSQL_Click()
    SQLCommand "Music"
End Sub
Private Sub rsMain_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
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
Private Sub dbcArtist_GotFocus()
    TextSelected
End Sub
Private Sub dbcArtist_Validate(Cancel As Boolean)
    If Not dbcArtist.Enabled Then Exit Sub
    If dbcArtist.Text = "" Then
        MsgBox "Artist must be specified!", vbExclamation, Me.Caption
        dbcArtist.SetFocus
        Cancel = True
    End If
    If dbcValidate(rsMain("Artist"), dbcArtist) = 0 Then Cancel = True
    If rsArtists.BookMark <> dbcArtist.SelectedItem Then rsArtists.BookMark = dbcArtist.SelectedItem
End Sub
Private Sub dbcType_GotFocus()
    TextSelected
End Sub
Private Sub dbcType_Validate(Cancel As Boolean)
    If Not dbcType.Enabled Then Exit Sub
    If dbcType.Text = "" Then
        MsgBox "Type must be specified!", vbExclamation, Me.Caption
        dbcType.SetFocus
        Cancel = True
    End If
    If dbcValidate(rsMain("Type"), dbcType) = 0 Then Cancel = True
    If rsTypes.BookMark <> dbcType.SelectedItem Then rsTypes.BookMark = dbcType.SelectedItem
End Sub
Private Function DefaultAlphaSort() As String
    Dim LastName As String
    Dim Title As String
    Dim iAnd As Integer
    Dim iAmpersand As Integer
    Dim iComma As Integer
    Dim iSemiColon As Integer
    Dim iSeparator As Integer
    
    'Start with the Artist's last name...
    LastName = dbcArtist.Text
    iAnd = InStr(dbcArtist.Text, " and ")
    iAmpersand = InStr(dbcArtist.Text, " & ")
    iComma = InStr(dbcArtist.Text, ",")
    iSemiColon = InStr(dbcArtist.Text, ";")
    
    If iComma > 0 Then
        'Assume the comma separates Artists, and...
        iSeparator = iComma
    ElseIf iSemiColon > 0 Then
        'Assume the semicolon separates Artists, and...
        iSeparator = iSemiColon
    ElseIf iAnd > 0 Then
        'Assume the "and" separates Artists, and...
        iSeparator = iAnd
    ElseIf iAmpersand > 0 Then
        'Assume the "&" separates Artists, and...
        iSeparator = iAmpersand
    End If
    
    If iSeparator > 0 Then
        '...take the first Artist...
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
    
    If UCase(dbcType.Text) = "SOUNDTRACK" Then
        DefaultAlphaSort = UCase(dbcType.Text & ": " & txtYear.Text & "; " & Title)
    ElseIf InStr(UCase(dbcArtist.Text), "TIME LIFE") > 0 Then
        DefaultAlphaSort = UCase(dbcType.Text & ": TIME LIFE: " & txtYear.Text & "; " & Title)
    Else
        DefaultAlphaSort = UCase(dbcType.Text & ": " & UCase(LastName) & ": " & txtYear.Text & "; " & Title)
    End If
End Function
Private Sub txtAlphaSort_GotFocus()
    If txtAlphaSort.Text = "" Then
        txtAlphaSort.Text = DefaultAlphaSort
    Else
        txtAlphaSort.Text = UCase(txtAlphaSort.Text)
    End If
    TextSelected
End Sub
Private Sub txtAlphaSort_KeyPress(KeyAscii As Integer)
    KeyPressUcase KeyAscii
End Sub
Private Sub txtAlphaSort_Validate(Cancel As Boolean)
    If Not txtAlphaSort.Enabled Then Exit Sub
    If txtAlphaSort.Text = "" Then
        MsgBox "AlphaSort must be specified!", vbExclamation, Me.Caption
        txtAlphaSort.SetFocus
        Cancel = True
    End If
End Sub
Private Sub txtInventoried_GotFocus()
    TextSelected
End Sub
Private Sub txtInventoried_Validate(Cancel As Boolean)
    On Error Resume Next
    txtInventoried.Text = Format(txtInventoried.Text, "mm/dd/yyyy hh:mm AMPM")
    If txtInventoried.Text = vbNullString Then txtInventoried.Text = Format(Now(), fmtDate)
    If Not IsDate(txtInventoried.Text) Then
        MsgBox "Invalid date format", vbExclamation
        Cancel = True
        Exit Sub
    End If
End Sub
Private Sub txtMisc_GotFocus()
    TextSelected
End Sub
Private Sub txtPrice_GotFocus()
    TextSelected
End Sub
Private Sub txtPrice_Validate(Cancel As Boolean)
    ValidateCurrency txtPrice.Text, Cancel
End Sub
Private Sub txtTitle_GotFocus()
    TextSelected
End Sub
Private Sub txtTitle_Validate(Cancel As Boolean)
    If Not txtTitle.Enabled Then Exit Sub
    If txtTitle.Text = "" Then
        MsgBox "Title must be specified!", vbExclamation, Me.Caption
        txtTitle.SetFocus
        Cancel = True
    End If
End Sub
Private Sub txtYear_GotFocus()
    TextSelected
End Sub
Private Sub txtYear_Validate(Cancel As Boolean)
    If Not txtYear.Enabled Then Exit Sub
    If txtYear.Text = "" Then
        MsgBox "Year must be specified!", vbExclamation, Me.Caption
        txtYear.SetFocus
        Cancel = True
    End If
End Sub

