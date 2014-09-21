VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmUSNClassifications 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "US Navy Ship Classifications"
   ClientHeight    =   2160
   ClientLeft      =   120
   ClientTop       =   348
   ClientWidth     =   6984
   Icon            =   "frmUSNClassifications.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   6984
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   4920
      TabIndex        =   5
      Top             =   1500
      Width           =   972
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   372
      Left            =   5916
      TabIndex        =   4
      Top             =   1500
      Width           =   972
   End
   Begin VB.TextBox txtDescription 
      Height          =   288
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "Description"
      Top             =   720
      Width           =   5532
   End
   Begin VB.TextBox txtType 
      Height          =   288
      Left            =   1320
      TabIndex        =   0
      Text            =   "Type"
      Top             =   360
      Width           =   4392
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   6
      Top             =   1908
      Width           =   6984
      _ExtentX        =   12319
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
            Object.Width           =   7133
            Key             =   "Message"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   2
            Object.Width           =   1270
            MinWidth        =   1270
            TextSave        =   "2:04 PM"
            Key             =   "Time"
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adodcMain 
      Height          =   312
      Left            =   120
      Top             =   1140
      Width           =   6732
      _ExtentX        =   11875
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
      Connect         =   "0"
      OLEDBString     =   "0"
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
   Begin MSComctlLib.Toolbar tbMain 
      Align           =   1  'Align Top
      Height          =   288
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   6984
      _ExtentX        =   12319
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
            Picture         =   "frmUSNClassifications.frx":0682
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNClassifications.frx":099E
            Key             =   "Warning"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNClassifications.frx":0CC6
            Key             =   "List"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNClassifications.frx":0FEE
            Key             =   "xNew"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNClassifications.frx":37A2
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNClassifications.frx":3BF6
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNClassifications.frx":404A
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNClassifications.frx":4B16
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNClassifications.frx":4E3E
            Key             =   "Sort"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNClassifications.frx":5292
            Key             =   "SQL"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNClassifications.frx":56E6
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNClassifications.frx":5B3A
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNClassifications.frx":5F92
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNClassifications.frx":60EE
            Key             =   "Blank"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNClassifications.frx":624A
            Key             =   "NewRecord"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblClassification 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Classification:"
      Height          =   192
      Left            =   240
      TabIndex        =   2
      Top             =   408
      Width           =   984
   End
   Begin VB.Label lblID 
      AutoSize        =   -1  'True
      Caption         =   "lblID"
      Height          =   192
      Left            =   360
      TabIndex        =   8
      Top             =   1656
      Width           =   324
   End
   Begin VB.Label lblIDlabel 
      AutoSize        =   -1  'True
      Caption         =   "ID:"
      Height          =   192
      Left            =   96
      TabIndex        =   7
      Top             =   1656
      Width           =   192
   End
   Begin VB.Label lblDescription 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Description:"
      Height          =   192
      Left            =   372
      TabIndex        =   3
      Top             =   768
      Width           =   852
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
Attribute VB_Name = "frmUSNClassifications"
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
Private Sub Form_Load()
    EstablishConnection adoConn
    
    Set rsMain = New ADODB.Recordset
    rsMain.CursorLocation = adUseClient
    SQLmain = "select * from [Classification] order by Type"
    SQLfilter = vbNullString
    SQLkey = "Reference"
    rsMain.Open SQLmain, adoConn, adOpenKeyset, adLockBatchOptimistic
    DBcollection.Add "rsMain", rsMain
    
    Set adodcMain.Recordset = rsMain
    BindField lblID, "ID", rsMain, "ID"
    BindField txtType, "Type", rsMain, "Type"
    BindField txtDescription, "Description", rsMain, "Description"

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
    
    txtType.SetFocus
End Sub
Private Sub mnuRecordsNew_Click()
    NewCommand Me, rsMain
    
    txtType.SetFocus
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
    ReportCommand Me, rsMain, App.Path & "\Reports\USN Ship Classifications.rpt"
End Sub
Private Sub mnuFileSQL_Click()
    SQLCommand "Classification"
End Sub
Private Sub rsMain_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Dim Caption As String
    
    If Not pRecordset.BOF And Not pRecordset.EOF Then Caption = "Reference #" & pRecordset.BookMark & ": " & pRecordset("Type") & " - " & pRecordset("Description")
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
Private Sub txtDescription_GotFocus()
    TextSelected
End Sub
Private Sub txtDescription_Validate(Cancel As Boolean)
    If Not txtDescription.Enabled Then Exit Sub
    If txtDescription.Text = vbNullString Then
        MsgBox "Description must be specified!", vbExclamation, Me.Caption
        Cancel = True
    End If
End Sub
Private Sub txtType_GotFocus()
    TextSelected
End Sub
Private Sub txtType_Validate(Cancel As Boolean)
    If Not txtType.Enabled Then Exit Sub
    If txtType.Text = vbNullString Then
        MsgBox "Type must be specified!", vbExclamation, Me.Caption
        Cancel = True
    End If
End Sub
