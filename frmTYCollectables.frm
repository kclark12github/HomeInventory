VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{C2000000-FFFF-1100-8100-000000000001}#8.0#0"; "PVCURR.OCX"
Begin VB.Form frmTYCollectables 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TY Collectables"
   ClientHeight    =   5625
   ClientLeft      =   30
   ClientTop       =   495
   ClientWidth     =   8100
   Icon            =   "frmTYCollectables.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraTYCollectables 
      Height          =   3675
      Index           =   0
      Left            =   180
      TabIndex        =   17
      Top             =   660
      Width           =   7692
      Begin PVCurrencyLib.PVCurrency pvcValue 
         Bindings        =   "frmTYCollectables.frx":08CA
         Height          =   285
         Left            =   3525
         TabIndex        =   5
         Top             =   900
         Width           =   1275
         _Version        =   524288
         _ExtentX        =   2249
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
      Begin PVCurrencyLib.PVCurrency pvcPrice 
         Bindings        =   "frmTYCollectables.frx":08D5
         Height          =   285
         Left            =   1530
         TabIndex        =   4
         Top             =   900
         Width           =   1275
         _Version        =   524288
         _ExtentX        =   2249
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
      Begin VB.TextBox txtDateVerified 
         Height          =   300
         Left            =   1530
         TabIndex        =   13
         Text            =   "Date Verified"
         Top             =   3060
         Width           =   3135
      End
      Begin VB.TextBox txtName 
         Height          =   300
         Left            =   1524
         TabIndex        =   2
         Text            =   "Name"
         Top             =   180
         Width           =   5892
      End
      Begin VB.TextBox txtDateInventoried 
         Height          =   300
         Left            =   1524
         TabIndex        =   12
         Text            =   "Date Inventoried"
         Top             =   2700
         Width           =   3135
      End
      Begin VB.TextBox txtReference 
         Height          =   300
         Left            =   5880
         TabIndex        =   9
         Text            =   "Reference"
         Top             =   1627
         Width           =   1452
      End
      Begin VB.CheckBox chkOutOfProduction 
         Caption         =   "Out of Production"
         Height          =   192
         Left            =   5460
         TabIndex        =   6
         Top             =   961
         Width           =   1932
      End
      Begin MSDataListLib.DataCombo dbcManufacturer 
         Height          =   315
         Left            =   1524
         TabIndex        =   7
         Top             =   1260
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   "Manufacturer"
      End
      Begin MSDataListLib.DataCombo dbcSeries 
         Height          =   315
         Left            =   1524
         TabIndex        =   8
         Top             =   1620
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   "Series"
      End
      Begin MSDataListLib.DataCombo dbcType 
         Height          =   315
         Left            =   1530
         TabIndex        =   3
         Top             =   540
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   "Type"
      End
      Begin MSDataListLib.DataCombo dbcCondition 
         Height          =   315
         Left            =   1524
         TabIndex        =   10
         Top             =   2340
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   "Condition"
      End
      Begin MSDataListLib.DataCombo dbcLocation 
         Height          =   315
         Left            =   4095
         TabIndex        =   11
         Top             =   2340
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   "Location"
      End
      Begin VB.Label lblValue 
         Alignment       =   1  'Right Justify
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
         Left            =   2955
         TabIndex        =   32
         Top             =   960
         Width           =   450
      End
      Begin VB.Label lblManufacturer 
         AutoSize        =   -1  'True
         Caption         =   "Manufacturer:"
         Height          =   195
         Left            =   456
         TabIndex        =   27
         Top             =   1320
         Width           =   960
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   195
         Left            =   936
         TabIndex        =   26
         Top             =   240
         Width           =   480
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
         Left            =   1011
         TabIndex        =   25
         Top             =   960
         Width           =   405
      End
      Begin VB.Label lblDateInventoried 
         AutoSize        =   -1  'True
         Caption         =   "Date Inventoried:"
         Height          =   195
         Left            =   201
         TabIndex        =   24
         Top             =   2753
         Width           =   1215
      End
      Begin VB.Label lblReference 
         AutoSize        =   -1  'True
         Caption         =   "Reference:"
         Height          =   195
         Left            =   4980
         TabIndex        =   23
         Top             =   1680
         Width           =   795
      End
      Begin VB.Label lblSeries 
         AutoSize        =   -1  'True
         Caption         =   "Series:"
         Height          =   195
         Left            =   936
         TabIndex        =   22
         Top             =   1680
         Width           =   480
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         Caption         =   "Type:"
         Height          =   195
         Left            =   996
         TabIndex        =   21
         Top             =   600
         Width           =   420
      End
      Begin VB.Label lblDateVerified 
         AutoSize        =   -1  'True
         Caption         =   "Date Verified:"
         Height          =   195
         Left            =   420
         TabIndex        =   20
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label lblLocation 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Location:"
         Height          =   195
         Left            =   3360
         TabIndex        =   19
         Top             =   2400
         Width           =   645
      End
      Begin VB.Label lblCondition 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Condition:"
         Height          =   195
         Left            =   711
         TabIndex        =   18
         Top             =   2400
         Width           =   705
      End
   End
   Begin VB.Frame fraTYCollectables 
      Height          =   3675
      Index           =   1
      Left            =   180
      TabIndex        =   30
      Top             =   660
      Width           =   7692
      Begin RichTextLib.RichTextBox rtxtNotes 
         Height          =   3435
         Left            =   60
         TabIndex        =   1
         Top             =   180
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   6059
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmTYCollectables.frx":08E0
      End
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   5370
      Width           =   8100
      _ExtentX        =   14288
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
            Object.Width           =   9075
            Key             =   "Message"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   2
            Object.Width           =   1270
            MinWidth        =   1270
            TextSave        =   "7:40 PM"
            Key             =   "Time"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   372
      Left            =   6984
      TabIndex        =   15
      Top             =   4920
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   5964
      TabIndex        =   14
      Top             =   4920
      Width           =   972
   End
   Begin MSAdodcLib.Adodc adodcMain 
      Height          =   330
      Left            =   495
      Top             =   4500
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
   Begin MSComctlLib.TabStrip tsTYCollectables 
      Height          =   4035
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   7117
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            Key             =   "General"
            Object.Tag             =   "General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Notes"
            Key             =   "Notes"
            Object.Tag             =   "Notes"
            Object.ToolTipText     =   "Notes"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbMain 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   8100
      _ExtentX        =   14288
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
            Picture         =   "frmTYCollectables.frx":09A9
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTYCollectables.frx":0CC5
            Key             =   "Warning"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTYCollectables.frx":0FED
            Key             =   "List"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTYCollectables.frx":1315
            Key             =   "xNew"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTYCollectables.frx":3AC9
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTYCollectables.frx":3F1D
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTYCollectables.frx":4371
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTYCollectables.frx":4E3D
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTYCollectables.frx":5165
            Key             =   "Sort"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTYCollectables.frx":55B9
            Key             =   "SQL"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTYCollectables.frx":5A0D
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTYCollectables.frx":5E61
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTYCollectables.frx":62B9
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTYCollectables.frx":6415
            Key             =   "Blank"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTYCollectables.frx":6571
            Key             =   "NewRecord"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblID 
      AutoSize        =   -1  'True
      Caption         =   "lblID"
      Height          =   195
      Left            =   429
      TabIndex        =   29
      Top             =   5040
      Width           =   330
   End
   Begin VB.Label lblIDlabel 
      AutoSize        =   -1  'True
      Caption         =   "ID:"
      Height          =   195
      Left            =   144
      TabIndex        =   28
      Top             =   5040
      Width           =   195
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
Attribute VB_Name = "frmTYCollectables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents rsMain As ADODB.Recordset
Attribute rsMain.VB_VarHelpID = -1
Dim rsManufacturers As New ADODB.Recordset
Dim rsSeries As New ADODB.Recordset
Dim rsScales As New ADODB.Recordset
Dim rsNations As New ADODB.Recordset
Dim rsConditions As New ADODB.Recordset
Dim rsLocations As New ADODB.Recordset
Dim rsTypes As New ADODB.Recordset
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
    'rsMain.CursorLocation = adUseServer
    rsMain.CursorLocation = adUseClient
    SQLmain = "select * from [TY Collectables] order by Manufacturer,Type,Series,Reference"
    SQLfilter = vbNullString
    SQLkey = "Reference"
    rsMain.Open SQLmain, adoConn, adOpenKeyset, adLockBatchOptimistic
    DBcollection.Add "rsMain", rsMain
    
    rsManufacturers.CursorLocation = adUseClient
    rsManufacturers.Open "select distinct Manufacturer from [TY Collectables] order by Manufacturer", adoConn, adOpenStatic, adLockReadOnly
    DBcollection.Add "rsManufacturers", rsManufacturers
    
    rsSeries.CursorLocation = adUseClient
    rsSeries.Open "select distinct Series from [TY Collectables] order by Series", adoConn, adOpenStatic, adLockReadOnly
    DBcollection.Add "rsSeries", rsSeries
    
    rsTypes.CursorLocation = adUseClient
    rsTypes.Open "select distinct Type from [TY Collectables] order by Type", adoConn, adOpenStatic, adLockReadOnly
    DBcollection.Add "rsTypes", rsTypes
    
    rsConditions.CursorLocation = adUseClient
    rsConditions.Open "select distinct Condition from [TY Collectables] order by Condition", adoConn, adOpenStatic, adLockReadOnly
    DBcollection.Add "rsConditions", rsConditions
    
    rsLocations.CursorLocation = adUseClient
    rsLocations.Open "select distinct Location from [TY Collectables] order by Location", adoConn, adOpenStatic, adLockReadOnly
    DBcollection.Add "rsLocations", rsLocations
    
    Set adodcMain.Recordset = rsMain
    BindField lblID, "ID", rsMain, "ID"
    BindField dbcManufacturer, "Manufacturer", rsMain, "Manufacturer", rsManufacturers, "Manufacturer", "Manufacturer"
    BindField txtName, "Name", rsMain, "Name"
    BindField pvcPrice, "Price", rsMain, "Price"
    BindField pvcValue, "Value", rsMain, "Value"
    BindField dbcType, "Type", rsMain, "Type", rsTypes, "Type", "Type"
    BindField txtReference, "Reference", rsMain, "Reference"
    BindField dbcSeries, "Series", rsMain, "Series", rsSeries, "Series", "Series"
    BindField dbcCondition, "Condition", rsMain, "Condition", rsConditions, "Condition", "Condition"
    BindField dbcLocation, "Location", rsMain, "Location", rsLocations, "Location", "Location"
    BindField chkOutOfProduction, "OutOfProduction", rsMain, "Out of Production"
    'BindField txtCount, "Count", rsMain, "Count"
    BindField txtDateInventoried, "DateInventoried", rsMain, "Date Inventoried"
    BindField txtDateVerified, "DateVerified", rsMain, "Date Verified"
    BindField rtxtNotes, "Notes", rsMain, vbNullString

    Set tsTYCollectables.SelectedItem = tsTYCollectables.Tabs(1)
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
    ListCommand Me, rsMain  ', False
End Sub
Private Sub mnuRecordsModify_Click()
    ModifyCommand Me
    
    Set tsTYCollectables.SelectedItem = tsTYCollectables.Tabs(1)
    txtName.SetFocus
End Sub
Private Sub mnuRecordsNew_Click()
    NewCommand Me, rsMain
    
    Set tsTYCollectables.SelectedItem = tsTYCollectables.Tabs(1)
    
    'Defaults...
    dbcCondition.BoundText = "New (boxed)"
    dbcLocation.BoundText = "Closet"
    txtDateInventoried.Text = Format(Now(), fmtDate)
    txtDateVerified.Text = Format(Now(), fmtDate)
    
    txtName.SetFocus
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
    ReportCommand Me, rsMain, App.Path & "\Reports\TYCollectables.rpt"
End Sub
Private Sub mnuFileSQL_Click()
    SQLCommand "TY Collectables"
End Sub
Private Sub rsMain_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Dim Caption As String
    
    If pError Is Nothing Then
        Call Trace(trcEnter, "rsMain_MoveComplete(""" & adoEventReason(adReason) & """, Nothing, """ & adoEventStatus(adStatusCancel) & """, pRecordset)")
    Else
        Call Trace(trcEnter, "rsMain_MoveComplete(""" & adoEventReason(adReason) & """, """ & pError.Description & """, """ & adoEventStatus(adStatusCancel) & """, pRecordset)")
    End If
    
    If Not pRecordset.BOF And Not pRecordset.EOF Then
        Caption = "Reference #" & pRecordset.BookMark & ": " & pRecordset("Type") & "; " & pRecordset("Series") & " " & pRecordset("Reference") & " - " & pRecordset("Name")
    End If
    UpdatePosition Me, Caption, pRecordset
    Call Trace(trcExit, "rsMain_MoveComplete")
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
Private Sub tsTYCollectables_Click()
    Dim i As Integer
    
    With tsTYCollectables
        For i = 0 To .Tabs.Count - 1
            If i = .SelectedItem.Index - 1 Then
                fraTYCollectables(i).Enabled = True
                fraTYCollectables(i).ZOrder
            Else
                fraTYCollectables(i).Enabled = False
            End If
        Next
    End With
End Sub
'=================================================================================
Private Sub dbcCondition_GotFocus()
    TextSelected
End Sub
Private Sub dbcCondition_Validate(Cancel As Boolean)
    If Trim(dbcCondition.Text) = vbNullString Then dbcCondition.Text = "Unknown"
    If dbcValidate(rsMain("Condition"), dbcCondition) = 0 Then Cancel = True
    If IsNull(dbcCondition.SelectedItem) Then Exit Sub
    If rsConditions.BookMark <> dbcCondition.SelectedItem Then rsConditions.BookMark = dbcCondition.SelectedItem
End Sub
Private Sub dbcLocation_GotFocus()
    TextSelected
End Sub
Private Sub dbcLocation_Validate(Cancel As Boolean)
    If Trim(dbcLocation.Text) = vbNullString Then dbcLocation.Text = "Unknown"
    If dbcValidate(rsMain("Location"), dbcLocation) = 0 Then Cancel = True
    If IsNull(dbcLocation.SelectedItem) Then Exit Sub
    If rsLocations.BookMark <> dbcLocation.SelectedItem Then rsLocations.BookMark = dbcLocation.SelectedItem
End Sub
Private Sub dbcManufacturer_GotFocus()
    TextSelected
End Sub
Private Sub dbcManufacturer_Validate(Cancel As Boolean)
    If Not dbcManufacturer.Enabled Then Exit Sub
    If dbcManufacturer.Text = vbNullString Then
        MsgBox "Manufacturer must be specified!", vbExclamation, Me.Caption
        dbcManufacturer.SetFocus
        Cancel = True
    End If
    If dbcValidate(rsMain("Manufacturer"), dbcManufacturer) = 0 Then Cancel = True
    If IsNull(dbcManufacturer.SelectedItem) Then Exit Sub
    If rsManufacturers.BookMark <> dbcManufacturer.SelectedItem Then rsManufacturers.BookMark = dbcManufacturer.SelectedItem
End Sub
Private Sub dbcSeries_GotFocus()
    TextSelected
End Sub
Private Sub dbcSeries_Validate(Cancel As Boolean)
    If Trim(dbcSeries.Text) = vbNullString Then dbcSeries.Text = "Unknown"
    If dbcValidate(rsMain("Series"), dbcSeries) = 0 Then Cancel = True
    If IsNull(dbcSeries.SelectedItem) Then Exit Sub
    If rsSeries.BookMark <> dbcSeries.SelectedItem Then rsSeries.BookMark = dbcSeries.SelectedItem
End Sub
Private Sub dbcType_GotFocus()
    TextSelected
End Sub
Private Sub dbcType_Validate(Cancel As Boolean)
    If Not dbcType.Enabled Then Exit Sub
    If dbcType.Text = vbNullString Then
        MsgBox "Type must be specified!", vbExclamation, Me.Caption
        dbcType.SetFocus
        Cancel = True
    End If
    If dbcValidate(rsMain("Type"), dbcType) = 0 Then Cancel = True
    If IsNull(dbcType.SelectedItem) Then Exit Sub
    If rsTypes.BookMark <> dbcType.SelectedItem Then rsTypes.BookMark = dbcType.SelectedItem
End Sub
Private Sub pvcPrice_GotFocus()
    TextSelected
End Sub
Private Sub pvcPrice_GotFocusEvent()
    TextSelected
End Sub
Private Sub pvcValue_GotFocus()
    TextSelected
End Sub
Private Sub pvcValue_GotFocusEvent()
    TextSelected
End Sub
Private Sub txtDateVerified_GotFocus()
    TextSelected
End Sub
Private Sub txtDateVerified_Validate(Cancel As Boolean)
    On Error Resume Next
    txtDateVerified.Text = Format(txtDateVerified.Text, "mm/dd/yyyy hh:mm AMPM")
    If txtDateVerified.Text = vbNullString Then txtDateVerified.Text = Format(Now(), fmtDate)
    If Not IsDate(txtDateVerified.Text) Then
        MsgBox "Invalid date format", vbExclamation
        Cancel = True
        Exit Sub
    End If
End Sub
Private Sub txtDateInventoried_GotFocus()
    TextSelected
End Sub
Private Sub txtDateInventoried_Validate(Cancel As Boolean)
    On Error Resume Next
    txtDateInventoried.Text = Format(txtDateInventoried.Text, "mm/dd/yyyy hh:mm AMPM")
    If txtDateInventoried.Text = vbNullString Then txtDateInventoried.Text = Format(Now(), fmtDate)
    If Not IsDate(txtDateInventoried.Text) Then
        MsgBox "Invalid date format", vbExclamation
        Cancel = True
        Exit Sub
    End If
End Sub
Private Sub txtName_GotFocus()
    TextSelected
End Sub
Private Sub txtName_Validate(Cancel As Boolean)
    If Not txtName.Enabled Then Exit Sub
    If txtName.Text = vbNullString Then
        MsgBox "Name must be specified!", vbExclamation, Me.Caption
        Cancel = True
    End If
End Sub
Private Sub txtReference_GotFocus()
    TextSelected
End Sub
Private Sub txtReference_KeyPress(KeyAscii As Integer)
    KeyPressUcase KeyAscii
End Sub
Private Sub txtReference_Validate(Cancel As Boolean)
    If txtReference.Text = vbNullString Then txtReference.Text = "Unknown"
End Sub
