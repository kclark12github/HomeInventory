VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSoftware 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Software Inventory"
   ClientHeight    =   4860
   ClientLeft      =   30
   ClientTop       =   495
   ClientWidth     =   7530
   Icon            =   "frmSoftware.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkCataloged 
      Alignment       =   1  'Right Justify
      Caption         =   "Cataloged"
      Height          =   192
      Left            =   4515
      TabIndex        =   12
      Top             =   3414
      Width           =   1155
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   27
      Top             =   4605
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
            TextSave        =   "3:17 PM"
            Key             =   "Time"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtISBN 
      Height          =   300
      Left            =   4800
      TabIndex        =   3
      Text            =   "ISBN"
      Top             =   1260
      Width           =   2475
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
      Left            =   4800
      TabIndex        =   5
      Top             =   1680
      Width           =   2475
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   372
      Left            =   6353
      TabIndex        =   14
      Top             =   4140
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   5333
      TabIndex        =   13
      Top             =   4140
      Width           =   972
   End
   Begin MSAdodcLib.Adodc adodcMain 
      Height          =   330
      Left            =   195
      Top             =   3780
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
      Left            =   1440
      TabIndex        =   11
      Text            =   "Inventoried"
      Top             =   3360
      Width           =   2535
   End
   Begin MSDataListLib.DataCombo dbcType 
      Height          =   315
      Left            =   1440
      TabIndex        =   8
      Top             =   2520
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   "Type"
   End
   Begin MSDataListLib.DataCombo dbcPublisher 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   840
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   "Publisher"
   End
   Begin VB.TextBox txtCDkey 
      Height          =   300
      Left            =   4800
      TabIndex        =   9
      Text            =   "CDkey"
      Top             =   2527
      Width           =   2472
   End
   Begin VB.TextBox txtValue 
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
      Left            =   1440
      TabIndex        =   4
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox txtVersion 
      Height          =   300
      Left            =   1440
      TabIndex        =   2
      Text            =   "Version"
      Top             =   1260
      Width           =   2175
   End
   Begin VB.TextBox txtTitle 
      Height          =   300
      Left            =   1440
      TabIndex        =   0
      Text            =   "Title"
      Top             =   420
      Width           =   5892
   End
   Begin MSDataListLib.DataCombo dbcPlatform 
      Height          =   315
      Left            =   1440
      TabIndex        =   6
      Top             =   2100
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   "Platform"
   End
   Begin MSComctlLib.ImageList imlSmall 
      Left            =   60
      Top             =   480
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
            Picture         =   "frmSoftware.frx":030A
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":0626
            Key             =   "Warning"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":094E
            Key             =   "List"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":0C76
            Key             =   "xNew"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":342A
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":387E
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":3CD2
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":479E
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":4AC6
            Key             =   "Sort"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":4F1A
            Key             =   "SQL"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":536E
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":57C2
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":5C1A
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":5D76
            Key             =   "Blank"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":5ED2
            Key             =   "NewRecord"
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo dbcMedia 
      Height          =   315
      Left            =   4800
      TabIndex        =   7
      Top             =   2100
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   "Media"
   End
   Begin MSComctlLib.Toolbar tbMain 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   30
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
   Begin MSDataListLib.DataCombo dbcLocation 
      Height          =   315
      Left            =   1440
      TabIndex        =   10
      Top             =   2940
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   "Location"
   End
   Begin VB.Label lblLocation 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Location:"
      Height          =   195
      Left            =   705
      TabIndex        =   29
      Top             =   3000
      Width           =   645
   End
   Begin VB.Label lblMedia 
      AutoSize        =   -1  'True
      Caption         =   "Media:"
      Height          =   195
      Left            =   4275
      TabIndex        =   28
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label lblPlatform 
      AutoSize        =   -1  'True
      Caption         =   "Platform:"
      Height          =   195
      Left            =   735
      TabIndex        =   26
      Top             =   2160
      Width           =   630
   End
   Begin VB.Label lblISBN 
      AutoSize        =   -1  'True
      Caption         =   "ISBN:"
      Height          =   195
      Left            =   4320
      TabIndex        =   25
      Top             =   1313
      Width           =   405
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
      Left            =   4410
      TabIndex        =   24
      Top             =   1733
      Width           =   360
   End
   Begin VB.Label lblIDlabel 
      AutoSize        =   -1  'True
      Caption         =   "ID:"
      Height          =   195
      Left            =   6750
      TabIndex        =   23
      Top             =   3413
      Width           =   195
   End
   Begin VB.Label lblDateInventoried 
      AutoSize        =   -1  'True
      Caption         =   "Date Inventoried:"
      Height          =   195
      Left            =   150
      TabIndex        =   22
      Top             =   3413
      Width           =   1215
   End
   Begin VB.Label lblCDkey 
      AutoSize        =   -1  'True
      Caption         =   "CD Key:"
      Height          =   195
      Left            =   4185
      TabIndex        =   21
      Top             =   2580
      Width           =   570
   End
   Begin VB.Label lblType 
      AutoSize        =   -1  'True
      Caption         =   "Type:"
      Height          =   195
      Left            =   930
      TabIndex        =   20
      Top             =   2580
      Width           =   420
   End
   Begin VB.Label lblValue 
      AutoSize        =   -1  'True
      Caption         =   "Value:"
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
      Left            =   900
      TabIndex        =   19
      Top             =   1733
      Width           =   450
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      Caption         =   "Version:"
      Height          =   195
      Left            =   765
      TabIndex        =   18
      Top             =   1313
      Width           =   585
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Title:"
      Height          =   195
      Left            =   1005
      TabIndex        =   17
      Top             =   480
      Width           =   345
   End
   Begin VB.Label lblPublisher 
      AutoSize        =   -1  'True
      Caption         =   "Publisher:"
      Height          =   195
      Left            =   645
      TabIndex        =   16
      Top             =   900
      Width           =   705
   End
   Begin VB.Label lblID 
      AutoSize        =   -1  'True
      Caption         =   "ID"
      Height          =   195
      Left            =   7035
      TabIndex        =   15
      Top             =   3413
      Width           =   150
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
Attribute VB_Name = "frmSoftware"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents rsMain As ADODB.Recordset
Attribute rsMain.VB_VarHelpID = -1
Dim rsPublishers As New ADODB.Recordset
Dim rsTypes As New ADODB.Recordset
Dim rsPlatforms As New ADODB.Recordset
Dim rsMedias As New ADODB.Recordset
Dim rsLocations As New ADODB.Recordset
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
    SQLmain = "select * from [Software] order by Type,Title"
    SQLfilter = vbNullString
    SQLkey = "Title"
    rsMain.Open SQLmain, adoConn, adOpenKeyset, adLockBatchOptimistic
    DBcollection.Add "rsMain", rsMain
    
    rsPublishers.CursorLocation = adUseClient
    rsPublishers.Open "select distinct Publisher from [Software] order by Publisher", adoConn, adOpenStatic, adLockReadOnly
    DBcollection.Add "rsPublishers", rsPublishers
    
    rsPlatforms.CursorLocation = adUseClient
    rsPlatforms.Open "select distinct Platform from [Software] order by Platform", adoConn, adOpenStatic, adLockReadOnly
    DBcollection.Add "rsPlatforms", rsPlatforms
    
    rsMedias.CursorLocation = adUseClient
    rsMedias.Open "select distinct Media from [Software] order by Media", adoConn, adOpenStatic, adLockReadOnly
    DBcollection.Add "rsMedias", rsMedias
    
    rsLocations.CursorLocation = adUseClient
    rsLocations.Open "select distinct Location from [Software] order by Location", adoConn, adOpenStatic, adLockReadOnly
    DBcollection.Add "rsLocations", rsLocations
    
    rsTypes.CursorLocation = adUseClient
    rsTypes.Open "select distinct Type from [Software] order by Type", adoConn, adOpenStatic, adLockReadOnly
    DBcollection.Add "rsTypes", rsTypes
    
    Set adodcMain.Recordset = rsMain
    BindField lblID, "ID", rsMain, "ID"
    BindField dbcPublisher, "Publisher", rsMain, "Publisher", rsPublishers, "Publisher", "Publisher"
    BindField txtTitle, "Title", rsMain, "Title"
    BindField txtVersion, "Version", rsMain, "Version"
    BindField txtISBN, "ISBN", rsMain, "ISBN"
    BindField txtValue, "Value", rsMain, "Value"
    BindField txtCost, "Cost", rsMain, "Cost"
    BindField dbcType, "Type", rsMain, "Type", rsTypes, "Type", "Type"
    BindField dbcPlatform, "Platform", rsMain, "Platform", rsPlatforms, "Platform", "Platform"
    BindField dbcMedia, "Media", rsMain, "Media", rsMedias, "Media", "Media"
    BindField dbcLocation, "Location", rsMain, "Location", rsLocations, "Location", "Location"
    BindField txtCDkey, "CDkey", rsMain, "CDkey"
    BindField chkCataloged, "Cataloged", rsMain, "Cataloged"
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
    
    txtTitle.SetFocus
End Sub
Private Sub mnuRecordsNew_Click()
    NewCommand Me, rsMain
    
    txtInventoried.Text = Format(Now(), fmtDate)
    chkCataloged.Value = vbChecked
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
    ReportCommand Me, rsMain, App.Path & "\Reports\Software.rpt"
End Sub
Private Sub mnuFileSQL_Click()
    SQLCommand "Software"
End Sub
Private Sub rsMain_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Dim Caption As String
    
    If Not pRecordset.BOF And Not pRecordset.EOF Then Caption = "Reference #" & pRecordset.BookMark & ": " & pRecordset("Type") & "; " & pRecordset("Title")
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
Private Sub dbcLocation_GotFocus()
    TextSelected
End Sub
Private Sub dbcLocation_Validate(Cancel As Boolean)
    If Not dbcLocation.Enabled Then Exit Sub
    If dbcLocation.Text = vbNullString Then
        MsgBox "Location must be specified!", vbExclamation, Me.Caption
        dbcLocation.SetFocus
        Cancel = True
    End If
    If dbcValidate(rsMain("Location"), dbcLocation) = 0 Then Cancel = True
    If rsLocations.BookMark <> dbcLocation.SelectedItem Then rsLocations.BookMark = dbcLocation.SelectedItem
End Sub
Private Sub dbcMedia_GotFocus()
    TextSelected
End Sub
Private Sub dbcMedia_Validate(Cancel As Boolean)
    If Not dbcMedia.Enabled Then Exit Sub
    If dbcMedia.Text = vbNullString Then
        MsgBox "Media must be specified!", vbExclamation, Me.Caption
        dbcMedia.SetFocus
        Cancel = True
    End If
    If dbcValidate(rsMain("Media"), dbcMedia) = 0 Then Cancel = True
    If rsMedias.BookMark <> dbcMedia.SelectedItem Then rsMedias.BookMark = dbcMedia.SelectedItem
End Sub
Private Sub dbcPublisher_GotFocus()
    TextSelected
End Sub
Private Sub dbcPublisher_Validate(Cancel As Boolean)
    If Not dbcPublisher.Enabled Then Exit Sub
    If dbcPublisher.Text = "" Then
        MsgBox "Publisher must be specified!", vbExclamation, Me.Caption
        dbcPublisher.SetFocus
        Cancel = True
    End If
    If dbcValidate(rsMain("Publisher"), dbcPublisher) = 0 Then Cancel = True
    If rsPublishers.BookMark <> dbcPublisher.SelectedItem Then rsPublishers.BookMark = dbcPublisher.SelectedItem
End Sub
Private Sub dbcPlatform_GotFocus()
    TextSelected
End Sub
Private Sub dbcPlatform_Validate(Cancel As Boolean)
    If Not dbcPlatform.Enabled Then Exit Sub
    If dbcPlatform.Text = "" Then
        MsgBox "Platform must be specified!", vbExclamation, Me.Caption
        dbcPlatform.SetFocus
        Cancel = True
    End If
    If dbcValidate(rsMain("Platform"), dbcPlatform) = 0 Then Cancel = True
    If rsPlatforms.BookMark <> dbcPlatform.SelectedItem Then rsPlatforms.BookMark = dbcPlatform.SelectedItem
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
Private Sub txtCDkey_GotFocus()
    TextSelected
End Sub
Private Sub txtCost_GotFocus()
    TextSelected
End Sub
Private Sub txtCost_Validate(Cancel As Boolean)
    ValidateCurrency txtCost.Text, Cancel
End Sub
Private Sub txtInventoried_GotFocus()
    TextSelected
End Sub
Private Sub txtInventoried_Validate(Cancel As Boolean)
    On Error Resume Next
    txtInventoried.Text = Format(txtInventoried.Text, "mm/dd/yyyy hh:mm AMPM")
    If txtInventoried.Text = vbNullString Or Format(txtInventoried.Text, "mm/dd/yyyy") = "07/31/1963" Then txtInventoried.Text = Format(Now(), fmtDate)
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
    If txtISBN.Text = vbNullString Then
        txtISBN.Text = "Unknown"
        'MsgBox "ISBN must be specified!", vbExclamation, Me.Caption
        'txtISBN.SetFocus
        'Cancel = True
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
Private Sub txtValue_GotFocus()
    TextSelected
End Sub
Private Sub txtValue_Validate(Cancel As Boolean)
    ValidateCurrency txtValue.Text, Cancel
End Sub
Private Sub txtVersion_GotFocus()
    TextSelected
End Sub
Private Sub txtVersion_Validate(Cancel As Boolean)
    If Not txtVersion.Enabled Then Exit Sub
    If txtVersion.Text = vbNullString Then
        txtVersion.Text = "Unknown"
        'MsgBox "Version must be specified!", vbExclamation, Me.Caption
        'txtVersion.SetFocus
        'Cancel = True
    End If
End Sub

