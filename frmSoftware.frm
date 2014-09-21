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
      TabIndex        =   26
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
            TextSave        =   "7:55 PM"
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
   Begin MSAdodcLib.Adodc adodcSoftware 
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
   Begin MSComctlLib.ImageList imlSmall 
      Left            =   480
      Top             =   3060
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":0000
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":0644
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":096C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":3120
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":3574
            Key             =   "List"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":4040
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":4494
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":4F60
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":5288
            Key             =   "Sort"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":56DC
            Key             =   "Filter"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlLarge 
      Left            =   60
      Top             =   3060
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":5B30
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":5F84
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":6A50
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":6D6C
            Key             =   "List"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":7838
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":7C8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":A440
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware.frx":A894
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbAction 
      Align           =   1  'Align Top
      Height          =   288
      Left            =   0
      TabIndex        =   25
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
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "List"
            Object.ToolTipText     =   "List all records"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh data"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Filter"
            Object.ToolTipText     =   "Filter"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New record"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Modify"
            Object.ToolTipText     =   "Modify record"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete record"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Report"
            Object.ToolTipText     =   "Report"
            ImageIndex      =   7
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Test"
                  Text            =   "Test"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Test2"
                  Text            =   "Test2"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
      BorderStyle     =   1
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
   Begin VB.Menu mnuAction 
      Caption         =   "&Action"
      Begin VB.Menu mnuActionList 
         Caption         =   "&List"
      End
      Begin VB.Menu mnuActionRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuActionFilter 
         Caption         =   "&Filter"
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
Attribute VB_Name = "frmSoftware"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoConn As ADODB.Connection
Dim WithEvents rsSoftware As ADODB.Recordset
Attribute rsSoftware.VB_VarHelpID = -1
Dim rsPublishers As New ADODB.Recordset
Dim rsTypes As New ADODB.Recordset
Dim rsPlatforms As New ADODB.Recordset
Dim mode As ActionMode
Dim fTransaction As Boolean
Dim DBinfo As DataBaseInfo
Private Sub cmdCancel_Click()
    Select Case mode
        Case modeDisplay
            Unload Me
        Case modeAdd, modeModify
            rsSoftware.CancelUpdate
            If mode = modeAdd Then rsSoftware.MoveLast
            adoConn.RollbackTrans
            fTransaction = False
            frmMain.ProtectFields Me
            mode = modeDisplay
            adodcSoftware.Enabled = True
    End Select
End Sub
Private Sub cmdOK_Click()
    Dim SaveBookmark As String
    
    Select Case mode
        Case modeDisplay
            Unload Me
        Case modeAdd, modeModify
            'Why we need to do this is buggy...
            rsSoftware("Publisher") = dbcPublisher.Text
            rsSoftware("Platform") = dbcPlatform.Text
            rsSoftware("Type") = dbcType.Text
            rsSoftware.UpdateBatch
            adoConn.CommitTrans
            fTransaction = False
            frmMain.ProtectFields Me
            mode = modeDisplay
            adodcSoftware.Enabled = True
            
            rsPublishers.Requery
            rsTypes.Requery
            rsPlatforms.Requery
    End Select
End Sub
Private Sub dbcPublisher_GotFocus()
    TextSelected
End Sub
Private Sub dbcPublisher_Validate(Cancel As Boolean)
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
    If dbcType.Text = "" Then
        MsgBox "Type must be specified!", vbExclamation, Me.Caption
        dbcType.SetFocus
        Cancel = True
    End If
End Sub
Private Sub Form_Load()
    Set adoConn = New ADODB.Connection
    Set rsSoftware = New ADODB.Recordset
    Set DBinfo = frmMain.DBcollection("Software")
    With DBinfo
        adoConn.Provider = .Provider
        adoConn.CommandTimeout = 60
        adoConn.ConnectionTimeout = 60
        adoConn.Open .PathName, .UserName, .Password
    End With
    rsSoftware.CursorLocation = adUseClient
    rsSoftware.Open "select * from [Software] order by Type,Title", adoConn, adOpenKeyset, adLockBatchOptimistic

    rsPublishers.CursorLocation = adUseClient
    rsPublishers.Open "select distinct Publisher from [Software] order by Publisher", adoConn, adOpenStatic, adLockReadOnly
    
    rsPlatforms.CursorLocation = adUseClient
    rsPlatforms.Open "select distinct Platform from [Software] order by Platform", adoConn, adOpenStatic, adLockReadOnly
    
    rsTypes.CursorLocation = adUseClient
    rsTypes.Open "select distinct Type from [Software] order by Type", adoConn, adOpenStatic, adLockReadOnly
    
    Set adodcSoftware.Recordset = rsSoftware
    frmMain.BindField lblID, "ID", rsSoftware
    frmMain.BindField dbcPublisher, "Publisher", rsSoftware, rsPublishers, "Publisher", "Publisher"
    frmMain.BindField txtTitle, "Title", rsSoftware
    frmMain.BindField txtVersion, "Version", rsSoftware
    frmMain.BindField txtISBN, "ISBN", rsSoftware
    frmMain.BindField txtValue, "Value", rsSoftware
    frmMain.BindField txtCost, "Cost", rsSoftware
    frmMain.BindField dbcType, "Type", rsSoftware, rsTypes, "Type", "Type"
    frmMain.BindField dbcPlatform, "Platform", rsSoftware, rsPlatforms, "Platform", "Platform"
    frmMain.BindField txtCDkey, "CDkey", rsSoftware
    frmMain.BindField chkCataloged, "Cataloged", rsSoftware
    frmMain.BindField txtInventoried, "DateInventoried", rsSoftware

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
    
    If Not rsSoftware.EOF Then
        If rsSoftware.EditMode <> adEditNone Then rsSoftware.CancelUpdate
    End If
    If (rsSoftware.State And adStateOpen) = adStateOpen Then rsSoftware.Close
    Set rsSoftware = Nothing
    rsPublishers.Close
    Set rsPublishers = Nothing
    rsPlatforms.Close
    Set rsPlatforms = Nothing
    rsTypes.Close
    Set rsTypes = Nothing
    
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
    
    Set frmList.rsList = rsSoftware
    Set frmList.dgdList.DataSource = frmList.rsList
    Set frmList.dgdList.Columns("Value").DataFormat = CurrencyFormat
    Set frmList.dgdList.Columns("Cost").DataFormat = CurrencyFormat
    For Each Col In frmList.dgdList.Columns
        Col.Alignment = dbgGeneral
    Next Col
    
    adoConn.BeginTrans
    fTransaction = True
    frmList.Show vbModal
    If rsSoftware.Filter <> vbNullString And rsSoftware.Filter <> 0 Then
        sbStatus.Panels("Message").Text = "Filter: " & rsSoftware.Filter
    End If
    adoConn.CommitTrans
    fTransaction = False
End Sub
Private Sub mnuActionRefresh_Click()
    Dim SaveBookmark As String
    
    SaveBookmark = rsSoftware("Title")
    rsSoftware.Requery
    rsSoftware.Find "Title='" & SQLQuote(SaveBookmark) & "'"
End Sub
Private Sub mnuActionFilter_Click()
    Dim frm As Form
    
    Load frmFilter
    frmFilter.Caption = Me.Caption & " Filter"
    If frmMain.Width > Me.Width And frmMain.Height > Me.Height Then
        Set frm = frmMain
    Else
        Set frm = Me
    End If
    frmFilter.Top = frm.Top
    frmFilter.Left = frm.Left
    frmFilter.Width = frm.Width
    frmFilter.Height = frm.Height
    
    Set frmFilter.RS = rsSoftware
    frmFilter.Show vbModal
    If rsSoftware.Filter <> vbNullString And rsSoftware.Filter <> 0 Then
        sbStatus.Panels("Message").Text = "Filter: " & rsSoftware.Filter
    End If
End Sub
Private Sub mnuActionNew_Click()
    mode = modeAdd
    frmMain.OpenFields Me
    adodcSoftware.Enabled = False
    rsSoftware.AddNew
    adoConn.BeginTrans
    fTransaction = True
    
    txtInventoried.Text = Format(Now(), "mm/dd/yyyy hh:nn AMPM")
    chkCataloged.Value = vbChecked
    dbcPublisher.SetFocus
End Sub
Private Sub mnuActionDelete_Click()
    mode = modeDelete
    If MsgBox("Are you sure you want to permanently delete this record...?", vbYesNo, Me.Caption) = vbYes Then
        rsSoftware.Delete
        rsSoftware.MoveNext
        If rsSoftware.EOF Then rsSoftware.MoveLast
    End If
    mode = modeDisplay
End Sub
Private Sub mnuActionModify_Click()
    mode = modeModify
    frmMain.OpenFields Me
    adodcSoftware.Enabled = False
    adoConn.BeginTrans
    fTransaction = True
    
    dbcPublisher.SetFocus
End Sub
Private Sub mnuActionReport_Click()
    Dim frm As Form
    Dim Report As New scrSoftwareReport
    Dim vRS As ADODB.Recordset
    
    MakeVirtualRecordset adoConn, rsSoftware, vRS
    
    Load frmViewReport
    frmViewReport.Caption = Me.Caption & " Report"
    If frmMain.Width > Me.Width And frmMain.Height > Me.Height Then
        Set frm = frmMain
    Else
        Set frm = Me
    End If
    frmViewReport.Top = frm.Top
    frmViewReport.Left = frm.Left
    frmViewReport.Width = frm.Width
    frmViewReport.Height = frm.Height
    frmViewReport.WindowState = vbMaximized
    
    Report.Database.SetDataSource vRS, 3, 1
    Report.ReadRecords
    
    frmViewReport.scrViewer.ReportSource = Report
    frmViewReport.Show vbModal
    
    Set Report = Nothing
    vRS.Close
    Set vRS = Nothing
End Sub
Private Sub rsSoftware_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Dim Caption As String
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    If rsSoftware.BOF And rsSoftware.EOF Then
        Caption = "No Records"
    ElseIf rsSoftware.EOF Then
        Caption = "EOF"
    ElseIf rsSoftware.BOF Then
        Caption = "BOF"
    Else
        Caption = "Reference #" & rsSoftware.Bookmark & ": " & rsSoftware("Type") & "; " & rsSoftware("Title")
    
        i = InStr(Caption, "&")
        If i > 0 Then Caption = Left(Caption, i) & "&" & Mid(Caption, i + 1)
        If rsSoftware.Filter <> vbNullString And rsSoftware.Filter <> 0 Then
            sbStatus.Panels("Message").Text = "Filter: " & rsSoftware.Filter
        End If
    End If
    
    adodcSoftware.Caption = Caption
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description & " (Error " & Err.Number & ")", vbExclamation, Me.Caption
    Resume Next
End Sub
Private Sub tbAction_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "List"
            mnuActionList_Click
        Case "Refresh"
            mnuActionRefresh_Click
        Case "Filter"
            mnuActionFilter_Click
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
Private Sub txtCDkey_GotFocus()
    TextSelected
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
Private Sub txtInventoried_GotFocus()
    TextSelected
End Sub
Private Sub txtInventoried_Validate(Cancel As Boolean)
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
    Dim Char As String
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
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
Private Sub txtValue_GotFocus()
    TextSelected
End Sub
Private Sub txtValue_KeyPress(KeyAscii As Integer)
    If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then
        If KeyAscii <> Asc(".") Then
            KeyAscii = 0    'Cancel the character.
            Beep            'Sound error signal.
        End If
    End If
End Sub
Private Sub txtValue_Validate(Cancel As Boolean)
    If txtValue.Text = vbNullString Then txtValue.Text = Format(0, "Currency")
    If Not IsNumeric(txtValue.Text) Then
        MsgBox "Invalid Value entered.", vbExclamation, Me.Caption
        TextSelected
        Cancel = True
    End If
End Sub
Private Sub txtVersion_GotFocus()
    TextSelected
End Sub

