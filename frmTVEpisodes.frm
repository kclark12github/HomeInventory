VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmTVEpisodes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TV Episodes"
   ClientHeight    =   4140
   ClientLeft      =   30
   ClientTop       =   495
   ClientWidth     =   7530
   Icon            =   "frmTVEpisodes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkTaped 
      Caption         =   "Taped"
      Height          =   192
      Left            =   5520
      TabIndex        =   9
      Top             =   2574
      Width           =   1035
   End
   Begin VB.CheckBox chkStoreBought 
      Caption         =   "Store Bought"
      Height          =   192
      Left            =   3840
      TabIndex        =   8
      Top             =   2574
      Width           =   1515
   End
   Begin VB.TextBox txtNumber 
      Height          =   300
      Left            =   5940
      TabIndex        =   2
      Text            =   "Number"
      Top             =   847
      Width           =   1392
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   20
      Top             =   3885
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
            Object.Width           =   8070
            Key             =   "Message"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   2
            Object.Width           =   1270
            MinWidth        =   1270
            TextSave        =   "3:34 PM"
            Key             =   "Time"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   372
      Left            =   6384
      TabIndex        =   11
      Top             =   3420
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   5364
      TabIndex        =   10
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
   Begin VB.TextBox txtInventoried 
      Height          =   300
      Left            =   1530
      TabIndex        =   7
      Text            =   "Inventoried"
      Top             =   2520
      Width           =   2115
   End
   Begin MSDataListLib.DataCombo dbcDistributor 
      Height          =   315
      Left            =   1530
      TabIndex        =   4
      Top             =   1680
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   556
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
      Height          =   300
      Left            =   5520
      TabIndex        =   6
      Top             =   2107
      Width           =   972
   End
   Begin VB.TextBox txtTitle 
      Height          =   300
      Left            =   1530
      TabIndex        =   0
      Text            =   "Title"
      Top             =   420
      Width           =   5868
   End
   Begin MSDataListLib.DataCombo dbcSubject 
      Height          =   315
      Left            =   1530
      TabIndex        =   5
      Top             =   2100
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   "Subject"
   End
   Begin MSDataListLib.DataCombo dbcSeries 
      Height          =   315
      Left            =   1530
      TabIndex        =   1
      Top             =   840
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   "Series"
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
            Picture         =   "frmTVEpisodes.frx":1CFA
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTVEpisodes.frx":2016
            Key             =   "Warning"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTVEpisodes.frx":233E
            Key             =   "List"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTVEpisodes.frx":2666
            Key             =   "xNew"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTVEpisodes.frx":4E1A
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTVEpisodes.frx":526E
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTVEpisodes.frx":56C2
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTVEpisodes.frx":618E
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTVEpisodes.frx":64B6
            Key             =   "Sort"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTVEpisodes.frx":690A
            Key             =   "SQL"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTVEpisodes.frx":6D5E
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTVEpisodes.frx":71B2
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTVEpisodes.frx":760A
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTVEpisodes.frx":7766
            Key             =   "Blank"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTVEpisodes.frx":78C2
            Key             =   "NewRecord"
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo dbcFormat 
      Height          =   315
      Left            =   1530
      TabIndex        =   3
      Top             =   1260
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   "Format"
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
   Begin VB.Label lblFormat 
      AutoSize        =   -1  'True
      Caption         =   "Format:"
      Height          =   195
      Left            =   888
      TabIndex        =   22
      Top             =   1320
      Width           =   540
   End
   Begin VB.Label lblNumber 
      AutoSize        =   -1  'True
      Caption         =   "Number:"
      Height          =   195
      Left            =   5280
      TabIndex        =   21
      Top             =   900
      Width           =   615
   End
   Begin VB.Label lblSeries 
      AutoSize        =   -1  'True
      Caption         =   "Series:"
      Height          =   195
      Left            =   918
      TabIndex        =   19
      Top             =   900
      Width           =   510
   End
   Begin VB.Label lblSubject 
      AutoSize        =   -1  'True
      Caption         =   "Subject:"
      Height          =   195
      Left            =   858
      TabIndex        =   18
      Top             =   2160
      Width           =   570
   End
   Begin VB.Label lblIDlabel 
      AutoSize        =   -1  'True
      Caption         =   "ID:"
      Height          =   195
      Left            =   6750
      TabIndex        =   17
      Top             =   2573
      Width           =   195
   End
   Begin VB.Label lblDateInventoried 
      AutoSize        =   -1  'True
      Caption         =   "Date Inventoried:"
      Height          =   195
      Left            =   213
      TabIndex        =   16
      Top             =   2573
      Width           =   1215
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
      Height          =   195
      Left            =   4980
      TabIndex        =   15
      Top             =   2160
      Width           =   360
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Title:"
      Height          =   192
      Left            =   1080
      TabIndex        =   14
      Top             =   474
      Width           =   348
   End
   Begin VB.Label lblDistributor 
      AutoSize        =   -1  'True
      Caption         =   "Distributor:"
      Height          =   195
      Left            =   678
      TabIndex        =   13
      Top             =   1740
      Width           =   750
   End
   Begin VB.Label lblID 
      AutoSize        =   -1  'True
      Caption         =   "lblID"
      Height          =   195
      Left            =   7050
      TabIndex        =   12
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
Attribute VB_Name = "frmTVEpisodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents rsMain As ADODB.Recordset
Attribute rsMain.VB_VarHelpID = -1
Dim rsSeries As New ADODB.Recordset
Dim rsDistributors As New ADODB.Recordset
Dim rsFormats As New ADODB.Recordset
Dim rsSubjects As New ADODB.Recordset
Dim strDefaultSort As String
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
    SQLmain = "select * from [Episodes] order by Series, Number"
    SQLfilter = vbNullString
    SQLkey = "ID"
    rsMain.Open SQLmain, adoConn, adOpenKeyset, adLockBatchOptimistic
    DBcollection.Add "rsMain", rsMain
    
    rsDistributors.CursorLocation = adUseClient
    rsDistributors.Open "select distinct Distributor from [Episodes] order by Distributor", adoConn, adOpenStatic, adLockReadOnly
    DBcollection.Add "rsDistributors", rsDistributors
    
    rsSeries.CursorLocation = adUseClient
    rsSeries.Open "select distinct Series from [Episodes] order by Series", adoConn, adOpenStatic, adLockReadOnly
    DBcollection.Add "rsSeries", rsSeries
    
    rsSubjects.CursorLocation = adUseClient
    rsSubjects.Open "select distinct Subject from [Episodes] order by Subject", adoConn, adOpenStatic, adLockReadOnly
    DBcollection.Add "rsSubjects", rsSubjects
    
    rsFormats.CursorLocation = adUseClient
    rsFormats.Open "select distinct Format from [Movies] order by Format", adoConn, adOpenStatic, adLockReadOnly
    DBcollection.Add "rsFormats", rsFormats
    
    Set adodcMain.Recordset = rsMain
    BindField lblID, "ID", rsMain, "ID"
    BindField dbcSeries, "Series", rsMain, "Series", rsSeries, "Series", "Series"
    BindField dbcDistributor, "Distributor", rsMain, "Distributor", rsDistributors, "Distributor", "Distributor"
    BindField dbcFormat, "Format", rsMain, "Format", rsFormats, "Format", "Format"
    BindField txtTitle, "Title", rsMain, "Title"
    BindField txtCost, "Cost", rsMain, "Cost"
    BindField dbcSubject, "Subject", rsMain, "Subject", rsSubjects, "Subject", "Subject"
    BindField txtNumber, "Number", rsMain, "Number"
    BindField txtInventoried, "DateInventoried", rsMain, "Date Inventoried"
    BindField chkStoreBought, "StoreBought", rsMain, "Store Bought"
    BindField chkTaped, "Taped", rsMain, "Taped"

    ProtectFields Me
    mode = modeDisplay
    fTransaction = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Cancel = CloseConnection(Me)
End Sub
Private Sub mnuRecordsCopy_Click()
    CopyCommand Me, rsMain, SQLkey
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
    
    txtTitle.SetFocus
End Sub
Private Sub mnuRecordsNew_Click()
    NewCommand Me, rsMain
    
    txtInventoried.Text = Format(Now(), fmtDate)
    txtTitle.SetFocus
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
    ReportCommand Me, rsMain, App.Path & "\Reports\TVEpisodes.rpt"
End Sub
Private Sub mnuFileSQL_Click()
    SQLCommand "Episodes"
End Sub
Private Sub rsMain_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Dim Caption As String
    
    If Not pRecordset.BOF And Not pRecordset.EOF Then
        Caption = "Reference #" & pRecordset.BookMark & ": " & pRecordset("Series")
        If Trim(pRecordset("Number")) <> vbNullString Then
            Caption = Caption & " Episode #" & pRecordset("Number")
        Else
            Caption = Caption & " Episode """ & pRecordset("Title") & """"
        End If
    End If
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
Private Sub dbcDistributor_GotFocus()
    TextSelected
End Sub
Private Sub dbcDistributor_Validate(Cancel As Boolean)
    If Not dbcDistributor.Enabled Then Exit Sub
    If dbcDistributor.Text = "" Then
        MsgBox "Distributor must be specified!", vbExclamation, Me.Caption
        dbcDistributor.SetFocus
        Cancel = True
    End If
    If dbcValidate(rsMain("Distributor"), dbcDistributor) = 0 Then Cancel = True
    If rsDistributors.BookMark <> dbcDistributor.SelectedItem Then rsDistributors.BookMark = dbcDistributor.SelectedItem
End Sub
Private Sub dbcFormat_GotFocus()
    TextSelected
End Sub
Private Sub dbcFormat_Validate(Cancel As Boolean)
    If Not dbcFormat.Enabled Then Exit Sub
    If dbcFormat.Text = "" Then
        MsgBox "Format must be specified!", vbExclamation, Me.Caption
        dbcFormat.SetFocus
        Cancel = True
    End If
    If dbcValidate(rsMain("Format"), dbcFormat) = 0 Then Cancel = True
    If rsFormats.BookMark <> dbcFormat.SelectedItem Then rsFormats.BookMark = dbcFormat.SelectedItem
End Sub
Private Sub dbcSeries_GotFocus()
    TextSelected
End Sub
Private Sub dbcSeries_Validate(Cancel As Boolean)
    If Not dbcSeries.Enabled Then Exit Sub
    If dbcSeries.Text = "" Then
        MsgBox "Series must be specified!", vbExclamation, Me.Caption
        dbcSeries.SetFocus
        Cancel = True
    End If
    If dbcValidate(rsMain("Series"), dbcSeries) = 0 Then Cancel = True
    If rsSeries.BookMark <> dbcSeries.SelectedItem Then rsSeries.BookMark = dbcSeries.SelectedItem
End Sub
Private Sub dbcSubject_GotFocus()
    TextSelected
End Sub
Private Sub dbcSubject_Validate(Cancel As Boolean)
    If Not dbcSubject.Enabled Then Exit Sub
    If dbcSubject.Text = "" Then
        MsgBox "Subject must be specified!", vbExclamation, Me.Caption
        dbcSubject.SetFocus
        Cancel = True
    End If
    If dbcValidate(rsMain("Subject"), dbcSubject) = 0 Then Cancel = True
    If rsSubjects.BookMark <> dbcSubject.SelectedItem Then rsSubjects.BookMark = dbcSubject.SelectedItem
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
Private Sub txtCost_GotFocus()
    TextSelected
End Sub
Private Sub txtCost_Validate(Cancel As Boolean)
    ValidateCurrency txtCost.Text, Cancel
End Sub
Private Sub txtNumber_GotFocus()
    TextSelected
End Sub
