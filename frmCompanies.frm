VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCompanies 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hobby Companies"
   ClientHeight    =   4608
   ClientLeft      =   36
   ClientTop       =   492
   ClientWidth     =   7524
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4608
   ScaleWidth      =   7524
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   20
      Top             =   4356
      Width           =   7524
      _ExtentX        =   13272
      _ExtentY        =   445
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
            Object.Width           =   8086
            Key             =   "Message"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   2
            Object.Width           =   1270
            MinWidth        =   1270
            TextSave        =   "4:52 PM"
            Key             =   "Time"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtWebSite 
      Height          =   288
      Left            =   1356
      TabIndex        =   7
      Text            =   "WebSite"
      Top             =   2940
      Width           =   5292
   End
   Begin VB.TextBox txtName 
      Height          =   288
      Left            =   1374
      TabIndex        =   0
      Text            =   "Name"
      Top             =   372
      Width           =   5892
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   372
      Left            =   6480
      TabIndex        =   9
      Top             =   3960
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   5460
      TabIndex        =   8
      Top             =   3960
      Width           =   972
   End
   Begin MSAdodcLib.Adodc adodcMain 
      Height          =   312
      Left            =   264
      Top             =   3480
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
      Caption         =   ""
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
   Begin VB.TextBox txtAddress 
      Height          =   1308
      Left            =   1374
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "frmCompanies.frx":0000
      Top             =   1596
      Width           =   5832
   End
   Begin VB.TextBox txtPhone 
      Height          =   288
      Left            =   4314
      TabIndex        =   5
      Text            =   "Phone"
      Top             =   1272
      Width           =   1812
   End
   Begin MSDataListLib.DataCombo dbcProductType 
      Height          =   288
      Left            =   1374
      TabIndex        =   4
      Top             =   1260
      Width           =   2052
      _ExtentX        =   3620
      _ExtentY        =   508
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   "ProductType"
   End
   Begin VB.TextBox txtAccount 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   0
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   288
      Left            =   4314
      TabIndex        =   3
      Text            =   "Account"
      Top             =   972
      Width           =   2352
   End
   Begin VB.TextBox txtCode 
      Height          =   288
      Left            =   1374
      TabIndex        =   2
      Text            =   "Code"
      Top             =   972
      Width           =   1692
   End
   Begin VB.TextBox txtShortName 
      Height          =   288
      Left            =   1374
      TabIndex        =   1
      Text            =   "Short Name"
      Top             =   672
      Width           =   5892
   End
   Begin MSComctlLib.ImageList imlSmall 
      Left            =   240
      Top             =   3840
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompanies.frx":0008
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompanies.frx":0324
            Key             =   "Warning"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompanies.frx":064C
            Key             =   "List"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompanies.frx":0974
            Key             =   "xNew"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompanies.frx":3128
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompanies.frx":357C
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompanies.frx":39D0
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompanies.frx":449C
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompanies.frx":47C4
            Key             =   "Sort"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompanies.frx":4C18
            Key             =   "SQL"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompanies.frx":506C
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompanies.frx":54C0
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompanies.frx":5918
            Key             =   "Blank"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompanies.frx":5D70
            Key             =   "NewRecord"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompanies.frx":6B84
            Key             =   "Delete"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbAction 
      Align           =   1  'Align Top
      Height          =   288
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   7524
      _ExtentX        =   13272
      _ExtentY        =   508
      ButtonWidth     =   487
      ButtonHeight    =   466
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
   Begin VB.Label lblWebSite 
      AutoSize        =   -1  'True
      Caption         =   "Web Site:"
      Height          =   192
      Left            =   360
      TabIndex        =   19
      Top             =   2988
      Width           =   696
   End
   Begin VB.Label lblIDlabel 
      AutoSize        =   -1  'True
      Caption         =   "ID:"
      Height          =   192
      Left            =   6804
      TabIndex        =   18
      Top             =   3060
      Width           =   192
   End
   Begin VB.Label lblPhone 
      AutoSize        =   -1  'True
      Caption         =   "Phone:"
      Height          =   192
      Left            =   3702
      TabIndex        =   17
      Top             =   1320
      Width           =   504
   End
   Begin VB.Label lblAddress 
      AutoSize        =   -1  'True
      Caption         =   "Address"
      Height          =   192
      Left            =   522
      TabIndex        =   16
      Top             =   1620
      Width           =   612
   End
   Begin VB.Label lblProductType 
      AutoSize        =   -1  'True
      Caption         =   "Product Type:"
      Height          =   192
      Left            =   258
      TabIndex        =   15
      Top             =   1320
      Width           =   1008
   End
   Begin VB.Label lblAccount 
      AutoSize        =   -1  'True
      Caption         =   "Account:"
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
      Left            =   3570
      TabIndex        =   14
      Top             =   1020
      Width           =   612
   End
   Begin VB.Label lblCode 
      AutoSize        =   -1  'True
      Caption         =   "Code:"
      Height          =   192
      Left            =   834
      TabIndex        =   13
      Top             =   1020
      Width           =   432
   End
   Begin VB.Label lblShortName 
      AutoSize        =   -1  'True
      Caption         =   "Short Name:"
      Height          =   192
      Left            =   378
      TabIndex        =   12
      Top             =   720
      Width           =   888
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   192
      Left            =   786
      TabIndex        =   11
      Top             =   420
      Width           =   480
   End
   Begin VB.Label lblID 
      AutoSize        =   -1  'True
      Caption         =   "lblID"
      Height          =   192
      Left            =   7092
      TabIndex        =   10
      Top             =   3060
      Width           =   324
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
      Begin VB.Menu mnuActionSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecordsRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuActionSep2 
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
Attribute VB_Name = "frmCompanies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents rsMain As ADODB.Recordset
Attribute rsMain.VB_VarHelpID = -1
Dim rsProductTypes As New ADODB.Recordset
Private Sub cmdCancel_Click()
    CancelCommand Me, rsMain
End Sub
Private Sub cmdOK_Click()
    OKCommand Me, rsMain
End Sub
Private Sub Form_Load()
    Set adoConn = New ADODB.Connection
    adoConn.Open "FileDSN=" & gstrFileDSN
    
    Set rsMain = New ADODB.Recordset
    rsMain.CursorLocation = adUseClient
    SQLmain = "select * from [Companies] order by Code"
    SQLfilter = vbNullString
    SQLkey = "Code"
    rsMain.Open SQLmain, adoConn, adOpenKeyset, adLockBatchOptimistic
    DBcollection.Add "rsMain", rsMain
    
    rsProductTypes.CursorLocation = adUseClient
    rsProductTypes.Open "select distinct ProductType from [Companies] order by ProductType", adoConn, adOpenStatic, adLockReadOnly
    DBcollection.Add "rsProductTypes", rsProductTypes
    
    Set adodcMain.Recordset = rsMain
    BindField lblID, "ID", rsMain
    BindField txtName, "Name", rsMain
    BindField txtShortName, "ShortName", rsMain
    BindField txtCode, "Code", rsMain
    BindField txtAccount, "Account", rsMain
    BindField txtPhone, "Phone", rsMain
    BindField txtAddress, "Address", rsMain
    BindField txtWebSite, "Website", rsMain
    BindField dbcProductType, "ProductType", rsMain, rsProductTypes, "ProductType", "ProductType"

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
    
    txtName.SetFocus
End Sub
Private Sub mnuRecordsNew_Click()
    NewCommand Me, rsMain
    
    txtName.SetFocus
End Sub
Private Sub mnuRecordsRefresh_Click()
    RefreshCommand rsMain, SQLkey
End Sub
Private Sub mnuFileReport_Click()
    ReportCommand Me, rsMain, App.Path & "\Reports\Companies.rpt"
End Sub
Private Sub mnuFileSQL_Click()
    SQLCommand "Companies"
End Sub
Private Sub rsMain_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Dim Caption As String
    
    If Not pRecordset.BOF And Not pRecordset.EOF Then Caption = "Reference #" & pRecordset.Bookmark & ": " & pRecordset("Name")
    UpdatePosition Me, Caption, pRecordset
End Sub
Private Sub tbAction_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "List"
            mnuRecordsList_Click
        Case "Refresh"
            mnuRecordsRefresh_Click
        Case "Filter"
            mnuRecordsFilter_Click
        Case "New"
            mnuRecordsNew_Click
        Case "Modify"
            mnuRecordsModify_Click
        Case "Delete"
            mnuRecordsDelete_Click
        Case "Report"
            mnuFileReport_Click
        Case "SQL"
            mnuFileSQL_Click
    End Select
End Sub
'=================================================================================
Private Sub dbcProductType_GotFocus()
    TextSelected
End Sub
Private Sub dbcProductType_Validate(Cancel As Boolean)
    If rsProductTypes.Bookmark <> dbcProductType.SelectedItem Then rsProductTypes.Bookmark = dbcProductType.SelectedItem
End Sub
Private Sub txtName_GotFocus()
    TextSelected
End Sub
Private Sub txtName_Validate(Cancel As Boolean)
    If Not txtName.Enabled Then Exit Sub
    If txtName.Text = "" Then
        MsgBox "Name must be specified!", vbExclamation, Me.Caption
        txtName.SetFocus
        Cancel = True
    End If
End Sub
Private Sub txtShortName_GotFocus()
    TextSelected
End Sub
Private Sub txtShortName_Validate(Cancel As Boolean)
    If Not txtShortName.Enabled Then Exit Sub
    If txtShortName.Text = "" Then
        MsgBox "Short Name must be specified!", vbExclamation, Me.Caption
        txtShortName.SetFocus
        Cancel = True
    End If
End Sub
Private Sub txtCode_GotFocus()
    TextSelected
End Sub
Private Sub txtCode_KeyPress(KeyAscii As Integer)
    KeyPressUcase KeyAscii
End Sub
Private Sub txtCode_Validate(Cancel As Boolean)
    If Not txtCode.Enabled Then Exit Sub
    If txtCode.Text = "" Then
        MsgBox "Code should be specified!", vbExclamation, Me.Caption
    End If
End Sub
Private Sub txtAccount_GotFocus()
    TextSelected
End Sub
Private Sub txtPhone_GotFocus()
    TextSelected
End Sub
Private Sub txtAddress_GotFocus()
    TextSelected
End Sub
Private Sub txtWebSite_GotFocus()
    TextSelected
End Sub
