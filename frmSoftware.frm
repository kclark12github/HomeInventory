VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSoftware 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Software Inventory"
   ClientHeight    =   3828
   ClientLeft      =   36
   ClientTop       =   492
   ClientWidth     =   7524
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3828
   ScaleWidth      =   7524
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   25
      Top             =   3576
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
            TextSave        =   "12:09 PM"
            Key             =   "Time"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtISBN 
      Height          =   288
      Left            =   4800
      TabIndex        =   3
      Text            =   "ISBN"
      Top             =   960
      Width           =   1512
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
      Left            =   4800
      TabIndex        =   5
      Top             =   1260
      Width           =   972
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   372
      Left            =   6420
      TabIndex        =   12
      Top             =   3180
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   5400
      TabIndex        =   11
      Top             =   3180
      Width           =   972
   End
   Begin MSAdodcLib.Adodc adodcMain 
      Height          =   312
      Left            =   204
      Top             =   2700
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
   Begin VB.TextBox txtInventoried 
      Height          =   288
      Left            =   1458
      TabIndex        =   9
      Text            =   "Inventoried"
      Top             =   2232
      Width           =   1812
   End
   Begin MSDataListLib.DataCombo dbcType 
      Height          =   288
      Left            =   1464
      TabIndex        =   6
      Top             =   1572
      Width           =   2292
      _ExtentX        =   4043
      _ExtentY        =   508
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   "Type"
   End
   Begin MSDataListLib.DataCombo dbcPublisher 
      Height          =   288
      Left            =   1464
      TabIndex        =   0
      Top             =   372
      Width           =   5892
      _ExtentX        =   10393
      _ExtentY        =   508
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   "Publisher"
   End
   Begin VB.TextBox txtCDkey 
      Height          =   288
      Left            =   1458
      TabIndex        =   8
      Text            =   "CDkey"
      Top             =   1896
      Width           =   3792
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
      Height          =   288
      Left            =   1458
      TabIndex        =   4
      Top             =   1272
      Width           =   972
   End
   Begin VB.TextBox txtVersion 
      Height          =   288
      Left            =   1458
      TabIndex        =   2
      Text            =   "Version"
      Top             =   960
      Width           =   1512
   End
   Begin VB.CheckBox chkCataloged 
      Alignment       =   1  'Right Justify
      Caption         =   "Cataloged"
      Height          =   192
      Left            =   3558
      TabIndex        =   10
      Top             =   2280
      Width           =   1152
   End
   Begin VB.TextBox txtTitle 
      Height          =   288
      Left            =   1458
      TabIndex        =   1
      Text            =   "Title"
      Top             =   672
      Width           =   5892
   End
   Begin MSDataListLib.DataCombo dbcPlatform 
      Height          =   288
      Left            =   4800
      TabIndex        =   7
      Top             =   1560
      Width           =   2292
      _ExtentX        =   4043
      _ExtentY        =   508
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   "Platform"
   End
   Begin MSComctlLib.Toolbar tbMain 
      Align           =   1  'Align Top
      Height          =   288
      Left            =   0
      TabIndex        =   26
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
   Begin MSComctlLib.ImageList imlSmall 
      Left            =   240
      Top             =   60
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
            Picture         =   "frmSoftware.frx":0000
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":031C
            Key             =   "Warning"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":0644
            Key             =   "List"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":096C
            Key             =   "xNew"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":3120
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":3574
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":39C8
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":4494
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":47BC
            Key             =   "Sort"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":4C10
            Key             =   "SQL"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":5064
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":54B8
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":5910
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":5A6C
            Key             =   "Blank"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":5BC8
            Key             =   "NewRecord"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblPlatform 
      AutoSize        =   -1  'True
      Caption         =   "Platform:"
      Height          =   192
      Left            =   4104
      TabIndex        =   24
      Top             =   1608
      Width           =   624
   End
   Begin VB.Label lblISBN 
      AutoSize        =   -1  'True
      Caption         =   "ISBN:"
      Height          =   192
      Left            =   4320
      TabIndex        =   23
      Top             =   1008
      Width           =   408
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
      Left            =   4368
      TabIndex        =   22
      Top             =   1308
      Width           =   360
   End
   Begin VB.Label lblIDlabel 
      AutoSize        =   -1  'True
      Caption         =   "ID:"
      Height          =   192
      Left            =   6744
      TabIndex        =   21
      Top             =   2280
      Width           =   192
   End
   Begin VB.Label lblDateInventoried 
      AutoSize        =   -1  'True
      Caption         =   "Date Inventoried:"
      Height          =   192
      Left            =   144
      TabIndex        =   20
      Top             =   2280
      Width           =   1212
   End
   Begin VB.Label lblCDkey 
      AutoSize        =   -1  'True
      Caption         =   "CD Key:"
      Height          =   192
      Left            =   828
      TabIndex        =   19
      Top             =   1920
      Width           =   576
   End
   Begin VB.Label lblType 
      AutoSize        =   -1  'True
      Caption         =   "Type:"
      Height          =   192
      Left            =   984
      TabIndex        =   18
      Top             =   1620
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
      Height          =   192
      Left            =   948
      TabIndex        =   17
      Top             =   1320
      Width           =   456
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      Caption         =   "Version:"
      Height          =   192
      Left            =   768
      TabIndex        =   16
      Top             =   1020
      Width           =   588
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Title:"
      Height          =   192
      Left            =   1008
      TabIndex        =   15
      Top             =   720
      Width           =   348
   End
   Begin VB.Label lblPublisher 
      AutoSize        =   -1  'True
      Caption         =   "Publisher:"
      Height          =   192
      Left            =   648
      TabIndex        =   14
      Top             =   420
      Width           =   708
   End
   Begin VB.Label lblID 
      AutoSize        =   -1  'True
      Caption         =   "lblID"
      Height          =   192
      Left            =   7032
      TabIndex        =   13
      Top             =   2280
      Width           =   324
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
Private Sub cmdCancel_Click()
    CancelCommand Me, rsMain
End Sub
Private Sub cmdOK_Click()
    OKCommand Me, rsMain
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
    
    rsTypes.CursorLocation = adUseClient
    rsTypes.Open "select distinct Type from [Software] order by Type", adoConn, adOpenStatic, adLockReadOnly
    DBcollection.Add "rsTypes", rsTypes
    
    Set adodcMain.Recordset = rsMain
    BindField lblID, "ID", rsMain
    BindField dbcPublisher, "Publisher", rsMain, rsPublishers, "Publisher", "Publisher"
    BindField txtTitle, "Title", rsMain
    BindField txtVersion, "Version", rsMain
    BindField txtISBN, "ISBN", rsMain
    BindField txtValue, "Value", rsMain
    BindField txtCost, "Cost", rsMain
    BindField dbcType, "Type", rsMain, rsTypes, "Type", "Type"
    BindField dbcPlatform, "Platform", rsMain, rsPlatforms, "Platform", "Platform"
    BindField txtCDkey, "CDkey", rsMain
    BindField chkCataloged, "Cataloged", rsMain
    BindField txtInventoried, "DateInventoried", rsMain

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
    
    dbcPublisher.SetFocus
End Sub
Private Sub mnuRecordsNew_Click()
    NewCommand Me, rsMain
    
    txtInventoried.Text = Format(Now(), fmtDate)
    chkCataloged.Value = vbChecked
    dbcPublisher.SetFocus
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
    
    If Not pRecordset.BOF And Not pRecordset.EOF Then Caption = "Reference #" & pRecordset.Bookmark & ": " & pRecordset("Type") & "; " & pRecordset("Title")
    UpdatePosition Me, Caption, pRecordset
End Sub
Private Sub tbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
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
    If Not txtInventoried.Enabled Then Exit Sub
    If txtInventoried.Text = "" Then
        MsgBox "Date Inventoried must be specified!", vbExclamation, Me.Caption
        txtInventoried.SetFocus
        Cancel = True
    End If
End Sub
Private Sub txtISBN_GotFocus()
    TextSelected
End Sub
Private Sub txtISBN_KeyPress(KeyAscii As Integer)
    KeyPressUcase KeyAscii
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

