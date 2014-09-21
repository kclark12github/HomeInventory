VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmBlueAngelsHistory 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Blue Angels History"
   ClientHeight    =   3480
   ClientLeft      =   36
   ClientTop       =   492
   ClientWidth     =   7524
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   7524
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   12
      Top             =   3228
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
            TextSave        =   "10:20 PM"
            Key             =   "Time"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtDecalSets 
      Height          =   888
      Left            =   5100
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "frmBlueAngelsHistory.frx":0000
      Top             =   960
      Width           =   2292
   End
   Begin VB.TextBox txtKits 
      Height          =   888
      Left            =   1530
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frmBlueAngelsHistory.frx":000B
      Top             =   960
      Width           =   2592
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   372
      Left            =   6486
      TabIndex        =   5
      Top             =   2820
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   5466
      TabIndex        =   4
      Top             =   2820
      Width           =   972
   End
   Begin MSAdodcLib.Adodc adodcMain 
      Height          =   312
      Left            =   270
      Top             =   2340
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
   Begin VB.TextBox txtDates 
      Height          =   288
      Left            =   1530
      TabIndex        =   1
      Text            =   "Dates"
      Top             =   672
      Width           =   1812
   End
   Begin VB.TextBox txtAircraftType 
      Height          =   288
      Left            =   1530
      TabIndex        =   0
      Text            =   "Aircraft Type"
      Top             =   372
      Width           =   5892
   End
   Begin MSComctlLib.ImageList imlSmall 
      Left            =   480
      Top             =   2760
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlueAngelsHistory.frx":0010
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlueAngelsHistory.frx":032C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlueAngelsHistory.frx":0654
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlueAngelsHistory.frx":097C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlueAngelsHistory.frx":3130
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlueAngelsHistory.frx":3584
            Key             =   "List"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlueAngelsHistory.frx":4050
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlueAngelsHistory.frx":44A4
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlueAngelsHistory.frx":4F70
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlueAngelsHistory.frx":5298
            Key             =   "Sort"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlueAngelsHistory.frx":56EC
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlueAngelsHistory.frx":5B40
            Key             =   "SQL"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlueAngelsHistory.frx":5F94
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlLarge 
      Left            =   60
      Top             =   2760
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
            Picture         =   "frmBlueAngelsHistory.frx":63E8
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlueAngelsHistory.frx":683C
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlueAngelsHistory.frx":7308
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlueAngelsHistory.frx":7624
            Key             =   "List"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlueAngelsHistory.frx":80F0
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlueAngelsHistory.frx":8544
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlueAngelsHistory.frx":ACF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlueAngelsHistory.frx":B14C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbAction 
      Align           =   1  'Align Top
      Height          =   288
      Left            =   0
      TabIndex        =   13
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
         NumButtons      =   11
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
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SQL"
            ImageIndex      =   13
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label lblDecalSets 
      AutoSize        =   -1  'True
      Caption         =   "Decal Sets:"
      Height          =   192
      Left            =   4260
      TabIndex        =   11
      Top             =   1020
      Width           =   828
   End
   Begin VB.Label lblIDlabel 
      AutoSize        =   -1  'True
      Caption         =   "ID:"
      Height          =   192
      Left            =   6810
      TabIndex        =   10
      Top             =   1920
      Width           =   192
   End
   Begin VB.Label lblDates 
      AutoSize        =   -1  'True
      Caption         =   "Dates:"
      Height          =   192
      Left            =   960
      TabIndex        =   9
      Top             =   720
      Width           =   468
   End
   Begin VB.Label lblAircraftType 
      AutoSize        =   -1  'True
      Caption         =   "Aircraft Type:"
      Height          =   192
      Left            =   474
      TabIndex        =   8
      Top             =   420
      Width           =   948
   End
   Begin VB.Label lblKits 
      AutoSize        =   -1  'True
      Caption         =   "Kits:"
      Height          =   192
      Left            =   1134
      TabIndex        =   7
      Top             =   1020
      Width           =   288
   End
   Begin VB.Label lblID 
      AutoSize        =   -1  'True
      Caption         =   "lblID"
      Height          =   192
      Left            =   7098
      TabIndex        =   6
      Top             =   1920
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
      Begin VB.Menu mnuActionSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActionSQL 
         Caption         =   "&SQL"
      End
   End
End
Attribute VB_Name = "frmBlueAngelsHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents rsMain As ADODB.Recordset
Attribute rsMain.VB_VarHelpID = -1
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
    SQLmain = "select * from [Blue Angels History] order by Dates"
    SQLfilter = vbNullString
    SQLkey = "Aircraft Type"
    rsMain.Open SQLmain, adoConn, adOpenKeyset, adLockBatchOptimistic
    DBcollection.Add "rsMain", rsMain
    
    Set adodcMain.Recordset = rsMain
    BindField lblID, "ID", rsMain
    BindField txtAircraftType, "Aircraft Type", rsMain
    BindField txtDates, "Dates", rsMain
    BindField txtKits, "Kits", rsMain
    BindField txtDecalSets, "Decal Sets", rsMain
    
    ProtectFields Me
    mode = modeDisplay
    fTransaction = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Cancel = CloseConnection(Me)
End Sub
Private Sub mnuActionFilter_Click()
    FilterCommand Me, rsMain, SQLkey
End Sub
Private Sub mnuActionDelete_Click()
    DeleteCommand Me, rsMain
End Sub
Private Sub mnuActionList_Click()
    ListCommand Me, rsMain
End Sub
Private Sub mnuActionModify_Click()
    ModifyCommand Me
End Sub
Private Sub mnuActionNew_Click()
    NewCommand Me, rsMain
End Sub
Private Sub mnuActionRefresh_Click()
    RefreshCommand rsMain, SQLkey
End Sub
Private Sub mnuActionReport_Click()
    ReportCommand Me, rsMain, App.Path & "\Reports\BlueAngels.rpt"
End Sub
Private Sub mnuActionSQL_Click()
    SQLCommand "Blue Angels History"
End Sub
Private Sub rsMain_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Dim Caption As String
    
    If Not pRecordset.BOF And Not pRecordset.EOF Then Caption = "Reference #" & pRecordset.Bookmark & ": " & pRecordset("Dates") & ": " & pRecordset("Aircraft Type")
    UpdatePosition Me, Caption, pRecordset
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
        Case "SQL"
            mnuActionSQL_Click
    End Select
End Sub
'=================================================================================
Private Sub txtAircraftType_GotFocus()
    TextSelected
End Sub
Private Sub txtAircraftType_Validate(Cancel As Boolean)
    If Not txtAircraftType.Enabled Then Exit Sub
    If txtAircraftType.Text = "" Then
        MsgBox "Aircraft Type must be specified!", vbExclamation, Me.Caption
        txtAircraftType.SetFocus
        Cancel = True
    End If
End Sub
Private Sub txtDates_GotFocus()
    TextSelected
End Sub
Private Sub txtDates_Validate(Cancel As Boolean)
    If Not txtDates.Enabled Then Exit Sub
    If txtDates.Text = "" Then
        MsgBox "Dates must be specified!", vbExclamation, Me.Caption
        txtDates.SetFocus
        Cancel = True
    End If
End Sub
Private Sub txtKits_GotFocus()
    TextSelected
End Sub
Private Sub txtDecalSets_GotFocus()
    TextSelected
End Sub

