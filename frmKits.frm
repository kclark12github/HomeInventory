VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{C2000000-FFFF-1100-8100-000000000001}#8.0#0"; "PVCURR.OCX"
Begin VB.Form frmKits 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Model Kits"
   ClientHeight    =   5625
   ClientLeft      =   30
   ClientTop       =   495
   ClientWidth     =   8100
   Icon            =   "frmKits.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraKits 
      Height          =   3675
      Index           =   0
      Left            =   180
      TabIndex        =   18
      Top             =   660
      Width           =   7692
      Begin PVCurrencyLib.PVCurrency pvcPrice 
         Bindings        =   "frmKits.frx":0442
         Height          =   285
         Left            =   3240
         TabIndex        =   36
         Top             =   900
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
      Begin VB.TextBox txtDateVerified 
         Height          =   300
         Left            =   1530
         TabIndex        =   14
         Text            =   "Date Verified"
         Top             =   3060
         Width           =   3135
      End
      Begin VB.TextBox txtName 
         Height          =   300
         Left            =   1524
         TabIndex        =   4
         Text            =   "Name"
         Top             =   540
         Width           =   5892
      End
      Begin VB.TextBox txtDesignation 
         Height          =   300
         Left            =   1524
         TabIndex        =   2
         Text            =   "Designation"
         Top             =   180
         Width           =   2172
      End
      Begin VB.TextBox txtDateInventoried 
         Height          =   300
         Left            =   1524
         TabIndex        =   13
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
         Left            =   4620
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
      Begin MSDataListLib.DataCombo dbcCatalog 
         Height          =   315
         Left            =   1524
         TabIndex        =   8
         Top             =   1620
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   "Catalog"
      End
      Begin MSDataListLib.DataCombo dbcNation 
         Height          =   315
         Left            =   1524
         TabIndex        =   10
         Top             =   1980
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   "Nation"
      End
      Begin MSDataListLib.DataCombo dbcScale 
         Height          =   315
         Left            =   1524
         TabIndex        =   5
         Top             =   900
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   "Scale"
      End
      Begin MSDataListLib.DataCombo dbcType 
         Height          =   315
         Left            =   4410
         TabIndex        =   3
         Top             =   180
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   "Type"
      End
      Begin MSDataListLib.DataCombo dbcCondition 
         Height          =   315
         Left            =   1524
         TabIndex        =   11
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
         TabIndex        =   12
         Top             =   2340
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   "Location"
      End
      Begin VB.Label lblManufacturer 
         AutoSize        =   -1  'True
         Caption         =   "Manufacturer:"
         Height          =   195
         Left            =   456
         TabIndex        =   31
         Top             =   1320
         Width           =   960
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   195
         Left            =   936
         TabIndex        =   30
         Top             =   593
         Width           =   480
      End
      Begin VB.Label lblDesignation 
         AutoSize        =   -1  'True
         Caption         =   "Designation:"
         Height          =   192
         Left            =   516
         TabIndex        =   29
         Top             =   234
         Width           =   900
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
         Left            =   2745
         TabIndex        =   28
         Top             =   960
         Width           =   405
      End
      Begin VB.Label lblNation 
         AutoSize        =   -1  'True
         Caption         =   "Nation:"
         Height          =   195
         Left            =   906
         TabIndex        =   27
         Top             =   2040
         Width           =   510
      End
      Begin VB.Label lblDateInventoried 
         AutoSize        =   -1  'True
         Caption         =   "Date Inventoried:"
         Height          =   195
         Left            =   201
         TabIndex        =   26
         Top             =   2753
         Width           =   1215
      End
      Begin VB.Label lblScale 
         AutoSize        =   -1  'True
         Caption         =   "Scale:"
         Height          =   195
         Left            =   966
         TabIndex        =   25
         Top             =   960
         Width           =   450
      End
      Begin VB.Label lblReference 
         AutoSize        =   -1  'True
         Caption         =   "Reference:"
         Height          =   195
         Left            =   4980
         TabIndex        =   24
         Top             =   1680
         Width           =   795
      End
      Begin VB.Label lblCatalog 
         AutoSize        =   -1  'True
         Caption         =   "Catalog:"
         Height          =   195
         Left            =   816
         TabIndex        =   23
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         Caption         =   "Type:"
         Height          =   192
         Left            =   3900
         TabIndex        =   22
         Top             =   241
         Width           =   420
      End
      Begin VB.Label lblDateVerified 
         AutoSize        =   -1  'True
         Caption         =   "Date Verified:"
         Height          =   195
         Left            =   420
         TabIndex        =   21
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label lblLocation 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Location:"
         Height          =   195
         Left            =   3360
         TabIndex        =   20
         Top             =   2400
         Width           =   645
      End
      Begin VB.Label lblCondition 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Condition:"
         Height          =   195
         Left            =   711
         TabIndex        =   19
         Top             =   2400
         Width           =   705
      End
   End
   Begin VB.Frame fraKits 
      Height          =   3675
      Index           =   1
      Left            =   180
      TabIndex        =   34
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
         TextRTF         =   $"frmKits.frx":044D
      End
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   17
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
      TabIndex        =   16
      Top             =   4920
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   5964
      TabIndex        =   15
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
   Begin MSComctlLib.TabStrip tsKits 
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
      TabIndex        =   35
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
            Picture         =   "frmKits.frx":0516
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKits.frx":0832
            Key             =   "Warning"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKits.frx":0B5A
            Key             =   "List"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKits.frx":0E82
            Key             =   "xNew"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKits.frx":3636
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKits.frx":3A8A
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKits.frx":3EDE
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKits.frx":49AA
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKits.frx":4CD2
            Key             =   "Sort"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKits.frx":5126
            Key             =   "SQL"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKits.frx":557A
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKits.frx":59CE
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKits.frx":5E26
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKits.frx":5F82
            Key             =   "Blank"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKits.frx":60DE
            Key             =   "NewRecord"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblID 
      AutoSize        =   -1  'True
      Caption         =   "lblID"
      Height          =   195
      Left            =   429
      TabIndex        =   33
      Top             =   5040
      Width           =   330
   End
   Begin VB.Label lblIDlabel 
      AutoSize        =   -1  'True
      Caption         =   "ID:"
      Height          =   195
      Left            =   144
      TabIndex        =   32
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
Attribute VB_Name = "frmKits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents rsMain As ADODB.Recordset
Attribute rsMain.VB_VarHelpID = -1
Dim rsManufacturers As New ADODB.Recordset
Dim rsCatalogs As New ADODB.Recordset
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
    SQLmain = "select * from [Kits] order by Manufacturer,Type,Reference,Scale"
    SQLfilter = vbNullString
    SQLkey = "Reference"
    rsMain.Open SQLmain, adoConn, adOpenKeyset, adLockBatchOptimistic
    DBcollection.Add "rsMain", rsMain
    
    rsManufacturers.CursorLocation = adUseClient
    rsManufacturers.Open "select distinct Manufacturer from [Kits] order by Manufacturer", adoConn, adOpenStatic, adLockReadOnly
    DBcollection.Add "rsManufacturers", rsManufacturers
    
    rsCatalogs.CursorLocation = adUseClient
    rsCatalogs.Open "select distinct Catalog from [Kits] order by Catalog", adoConn, adOpenStatic, adLockReadOnly
    DBcollection.Add "rsCatalogs", rsCatalogs
    
    rsScales.CursorLocation = adUseClient
    rsScales.Open "select distinct Scale from [Kits] order by Scale", adoConn, adOpenStatic, adLockReadOnly
    DBcollection.Add "rsScales", rsScales
    
    rsTypes.CursorLocation = adUseClient
    rsTypes.Open "select distinct Type from [Kits] order by Type", adoConn, adOpenStatic, adLockReadOnly
    DBcollection.Add "rsTypes", rsTypes
    
    rsNations.CursorLocation = adUseClient
    rsNations.Open "select distinct Nation from [Kits] order by Nation", adoConn, adOpenStatic, adLockReadOnly
    DBcollection.Add "rsNations", rsNations
    
    rsConditions.CursorLocation = adUseClient
    rsConditions.Open "select distinct Condition from [Kits] order by Condition", adoConn, adOpenStatic, adLockReadOnly
    DBcollection.Add "rsConditions", rsConditions
    
    rsLocations.CursorLocation = adUseClient
    rsLocations.Open "select distinct Location from [Kits] order by Location", adoConn, adOpenStatic, adLockReadOnly
    DBcollection.Add "rsLocations", rsLocations
    
    Set adodcMain.Recordset = rsMain
    BindField lblID, "ID", rsMain, "ID"
    BindField dbcManufacturer, "Manufacturer", rsMain, "Manufacturer", rsManufacturers, "Manufacturer", "Manufacturer"
    BindField txtDesignation, "Designation", rsMain, "Designation"
    BindField txtName, "Name", rsMain, "Name"
    BindField pvcPrice, "Price", rsMain, "Price"
    BindField dbcScale, "Scale", rsMain, "Scale", rsScales, "Scale", "Scale"
    BindField dbcType, "Type", rsMain, "Type", rsTypes, "Type", "Type"
    BindField txtReference, "Reference", rsMain, "Reference"
    BindField dbcCatalog, "Catalog", rsMain, "Catalog", rsCatalogs, "Catalog", "Catalog"
    BindField dbcNation, "Nation", rsMain, "Nation", rsNations, "Nation", "Nation"
    BindField dbcCondition, "Condition", rsMain, "Condition", rsConditions, "Condition", "Condition"
    BindField dbcLocation, "Location", rsMain, "Location", rsLocations, "Location", "Location"
    BindField chkOutOfProduction, "OutOfProduction", rsMain, "Out of Production"
    'BindField txtCount, "Count", rsMain, "Count"
    BindField txtDateInventoried, "DateInventoried", rsMain, "Date Inventoried"
    BindField txtDateVerified, "DateVerified", rsMain, "Date Verified"
    BindField rtxtNotes, "Notes", rsMain, vbNullString

    Set tsKits.SelectedItem = tsKits.Tabs(1)
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
    
    Set tsKits.SelectedItem = tsKits.Tabs(1)
    txtDesignation.SetFocus
End Sub
Private Sub mnuRecordsNew_Click()
    NewCommand Me, rsMain
    
    Set tsKits.SelectedItem = tsKits.Tabs(1)
    
    'Defaults...
    dbcCondition.BoundText = "New (boxed)"
    dbcLocation.BoundText = "Closet"
    txtDateInventoried.Text = Format(Now(), fmtDate)
    txtDateVerified.Text = Format(Now(), fmtDate)
    
    txtDesignation.SetFocus
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
    ReportCommand Me, rsMain, App.Path & "\Reports\Kits.rpt"
End Sub
Private Sub mnuFileSQL_Click()
    SQLCommand "Kits"
End Sub
Private Sub rsMain_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Dim Caption As String
    
    If pError Is Nothing Then
        Call Trace(trcEnter, "rsMain_MoveComplete(""" & adoEventReason(adReason) & """, Nothing, """ & adoEventStatus(adStatusCancel) & """, pRecordset)")
    Else
        Call Trace(trcEnter, "rsMain_MoveComplete(""" & adoEventReason(adReason) & """, """ & pError.Description & """, """ & adoEventStatus(adStatusCancel) & """, pRecordset)")
    End If
    
    If Not pRecordset.BOF And Not pRecordset.EOF Then
        Caption = "Reference #" & pRecordset.BookMark & ": "
        If IsNumeric(rsMain("Scale")) Then Caption = Caption & "1/"
        Caption = Caption & pRecordset("Scale") & " Scale; " & pRecordset("Designation") & " " & pRecordset("Name")
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
Private Sub tsKits_Click()
    Dim i As Integer
    
    With tsKits
        For i = 0 To .Tabs.Count - 1
            If i = .SelectedItem.Index - 1 Then
                fraKits(i).Enabled = True
                fraKits(i).ZOrder
            Else
                fraKits(i).Enabled = False
            End If
        Next
    End With
End Sub
'=================================================================================
Private Sub dbcCatalog_GotFocus()
    TextSelected
End Sub
Private Sub dbcCatalog_Validate(Cancel As Boolean)
    If Trim(dbcCatalog.Text) = vbNullString Then dbcCatalog.Text = "Unknown"
    If dbcValidate(rsMain("Catalog"), dbcCatalog) = 0 Then Cancel = True
    If rsCatalogs.BookMark <> dbcCatalog.SelectedItem Then rsCatalogs.BookMark = dbcCatalog.SelectedItem
End Sub
Private Sub dbcCondition_GotFocus()
    TextSelected
End Sub
Private Sub dbcCondition_Validate(Cancel As Boolean)
    If Trim(dbcCondition.Text) = vbNullString Then dbcCondition.Text = "Unknown"
    If dbcValidate(rsMain("Condition"), dbcCondition) = 0 Then Cancel = True
    If rsConditions.BookMark <> dbcCondition.SelectedItem Then rsConditions.BookMark = dbcCondition.SelectedItem
End Sub
Private Sub dbcLocation_GotFocus()
    TextSelected
End Sub
Private Sub dbcLocation_Validate(Cancel As Boolean)
    If Trim(dbcLocation.Text) = vbNullString Then dbcLocation.Text = "Unknown"
    If dbcValidate(rsMain("Location"), dbcLocation) = 0 Then Cancel = True
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
    If rsManufacturers.BookMark <> dbcManufacturer.SelectedItem Then rsManufacturers.BookMark = dbcManufacturer.SelectedItem
End Sub
Private Sub dbcNation_GotFocus()
    TextSelected
End Sub
Private Sub dbcNation_Validate(Cancel As Boolean)
    If Not dbcNation.Enabled Then Exit Sub
    If dbcNation.Text = vbNullString Then
        MsgBox "Nation must be specified!", vbExclamation, Me.Caption
        dbcNation.SetFocus
        Cancel = True
    End If
    If dbcValidate(rsMain("Nation"), dbcNation) = 0 Then Cancel = True
    If rsNations.BookMark <> dbcNation.SelectedItem Then rsNations.BookMark = dbcNation.SelectedItem
End Sub
Private Sub dbcScale_GotFocus()
    TextSelected
End Sub
Private Sub dbcScale_Validate(Cancel As Boolean)
    If dbcScale.Text = vbNullString Then dbcScale.Text = "Unknown"
    If dbcValidate(rsMain("Scale"), dbcScale) = 0 Then Cancel = True
    If rsScales.BookMark <> dbcScale.SelectedItem Then rsScales.BookMark = dbcScale.SelectedItem
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
    If rsTypes.BookMark <> dbcType.SelectedItem Then rsTypes.BookMark = dbcType.SelectedItem
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
Private Sub txtDesignation_GotFocus()
    TextSelected
End Sub
Private Sub txtDesignation_KeyPress(KeyAscii As Integer)
    KeyPressUcase KeyAscii
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
