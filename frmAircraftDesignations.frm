VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAircraftDesignations 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aircraft Designations"
   ClientHeight    =   4860
   ClientLeft      =   36
   ClientTop       =   492
   ClientWidth     =   7524
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   7524
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtNumber 
      Height          =   288
      Left            =   1458
      TabIndex        =   19
      Text            =   "Number"
      Top             =   1596
      Width           =   1872
   End
   Begin MSComctlLib.Toolbar tbDesignations 
      Align           =   1  'Align Top
      Height          =   288
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   7524
      _ExtentX        =   13272
      _ExtentY        =   508
      ButtonWidth     =   1439
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlSmall"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "List"
            Key             =   "List"
            Object.ToolTipText     =   "List all records"
            ImageKey        =   "List"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            Key             =   "New"
            Object.ToolTipText     =   "New record"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Modify"
            Key             =   "Modify"
            Object.ToolTipText     =   "Modify record"
            ImageKey        =   "Modify"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete record"
            ImageKey        =   "Delete"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Report"
            Key             =   "Report"
            Object.ToolTipText     =   "Report"
            ImageKey        =   "Report"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSmall 
      Left            =   120
      Top             =   4380
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAircraftDesignations.frx":0000
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAircraftDesignations.frx":031C
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAircraftDesignations.frx":0638
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAircraftDesignations.frx":0A8C
            Key             =   "List"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAircraftDesignations.frx":1558
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAircraftDesignations.frx":2224
            Key             =   "New2"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAircraftDesignations.frx":2CF0
            Key             =   "Delete2"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAircraftDesignations.frx":37BC
            Key             =   "Modify2"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAircraftDesignations.frx":4288
            Key             =   "New"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAircraftDesignations.frx":4D54
            Key             =   "Report"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   372
      Left            =   6420
      TabIndex        =   8
      Top             =   4380
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   5400
      TabIndex        =   7
      Top             =   4380
      Width           =   972
   End
   Begin MSAdodcLib.Adodc adodcDesignations 
      Height          =   312
      Left            =   204
      Top             =   3900
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
   Begin VB.TextBox txtNotes 
      Height          =   1428
      Left            =   1458
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "frmAircraftDesignations.frx":51A8
      Top             =   1956
      Width           =   5832
   End
   Begin VB.TextBox txtServiceDate 
      Height          =   288
      Left            =   4284
      TabIndex        =   6
      Text            =   "Service Date"
      Top             =   972
      Width           =   1812
   End
   Begin MSDataListLib.DataCombo dbcType 
      Height          =   288
      Left            =   1458
      TabIndex        =   3
      Top             =   1272
      Width           =   5892
      _ExtentX        =   10393
      _ExtentY        =   508
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   "Type"
   End
   Begin MSDataListLib.DataCombo dbcManufacturer 
      Height          =   288
      Left            =   1458
      TabIndex        =   0
      Top             =   672
      Width           =   5892
      _ExtentX        =   10393
      _ExtentY        =   508
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   "Manufacturer"
   End
   Begin VB.TextBox txtVersion 
      Height          =   288
      Left            =   5184
      TabIndex        =   4
      Text            =   "Version"
      Top             =   1596
      Width           =   1872
   End
   Begin VB.TextBox txtDesignation 
      Height          =   288
      Left            =   1458
      TabIndex        =   2
      Text            =   "Designation"
      Top             =   972
      Width           =   1512
   End
   Begin VB.TextBox txtName 
      Height          =   288
      Left            =   1458
      TabIndex        =   1
      Text            =   "Name"
      Top             =   372
      Width           =   5892
   End
   Begin MSComctlLib.ImageList imlLarge 
      Left            =   540
      Top             =   4380
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAircraftDesignations.frx":51AE
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAircraftDesignations.frx":5602
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAircraftDesignations.frx":591E
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAircraftDesignations.frx":5C3A
            Key             =   "New2"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAircraftDesignations.frx":6706
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAircraftDesignations.frx":73D2
            Key             =   "List"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAircraftDesignations.frx":7E9E
            Key             =   "Delete2"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAircraftDesignations.frx":896A
            Key             =   "Modify2"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAircraftDesignations.frx":9436
            Key             =   "New"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAircraftDesignations.frx":9F02
            Key             =   "Delete"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblNumber 
      AutoSize        =   -1  'True
      Caption         =   "Number:"
      Height          =   192
      Left            =   780
      TabIndex        =   20
      Top             =   1644
      Width           =   612
   End
   Begin VB.Label lblIDlabel 
      AutoSize        =   -1  'True
      Caption         =   "ID:"
      Height          =   192
      Left            =   6744
      TabIndex        =   17
      Top             =   3480
      Width           =   192
   End
   Begin VB.Label lblServiceDate 
      AutoSize        =   -1  'True
      Caption         =   "Service Date:"
      Height          =   192
      Left            =   3204
      TabIndex        =   16
      Top             =   1020
      Width           =   972
   End
   Begin VB.Label lblNotes 
      AutoSize        =   -1  'True
      Caption         =   "Notes:"
      Height          =   192
      Left            =   924
      TabIndex        =   15
      Top             =   1980
      Width           =   468
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      Caption         =   "Version:"
      Height          =   192
      Left            =   4488
      TabIndex        =   14
      Top             =   1644
      Width           =   588
   End
   Begin VB.Label lblType 
      AutoSize        =   -1  'True
      Caption         =   "Type:"
      Height          =   192
      Left            =   936
      TabIndex        =   13
      Top             =   1320
      Width           =   420
   End
   Begin VB.Label lblDesignation 
      AutoSize        =   -1  'True
      Caption         =   "Designation:"
      Height          =   192
      Left            =   456
      TabIndex        =   12
      Top             =   1020
      Width           =   900
   End
   Begin VB.Label lblManufacturer 
      AutoSize        =   -1  'True
      Caption         =   "Manufacturer:"
      Height          =   192
      Left            =   396
      TabIndex        =   11
      Top             =   720
      Width           =   960
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   192
      Left            =   876
      TabIndex        =   10
      Top             =   420
      Width           =   480
   End
   Begin VB.Label lblID 
      AutoSize        =   -1  'True
      Caption         =   "lblID"
      Height          =   192
      Left            =   7032
      TabIndex        =   9
      Top             =   3480
      Width           =   324
   End
   Begin VB.Menu mnuAction 
      Caption         =   "&Action"
      Begin VB.Menu mnuActionList 
         Caption         =   "&List"
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
Attribute VB_Name = "frmAircraftDesignations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoConn As ADODB.Connection
Dim WithEvents rsDesignations As ADODB.Recordset
Attribute rsDesignations.VB_VarHelpID = -1
Dim rsManufacturers As New ADODB.Recordset
Dim rsTypes As New ADODB.Recordset
Dim mode As ActionMode
Dim fTransaction As Boolean
Dim DBinfo As DataBaseInfo
Private Sub cmdCancel_Click()
    Select Case mode
        Case modeDisplay
            Unload Me
        Case modeAdd, modeModify
            rsDesignations.CancelUpdate
            adoConn.RollbackTrans
            fTransaction = False
            frmMain.ProtectFields Me
            mode = modeDisplay
            adodcDesignations.Enabled = True
    End Select
End Sub
Private Sub cmdOK_Click()
    Dim SaveBookmark As String
    
    Select Case mode
        Case modeDisplay
            Unload Me
        Case modeAdd, modeModify
            'Why we need to do this is buggy...
            rsDesignations("Manufacturer") = dbcManufacturer.Text
            rsDesignations("Type") = dbcType.Text
            rsDesignations.UpdateBatch
            adoConn.CommitTrans
            fTransaction = False
            frmMain.ProtectFields Me
            mode = modeDisplay
            adodcDesignations.Enabled = True
            
            SaveBookmark = rsDesignations("Name")
            rsDesignations.Requery
            rsDesignations.Find "Name='" & SaveBookmark & "'"
            rsManufacturers.Requery
            rsTypes.Requery
    End Select
End Sub
Private Sub dbcManufacturer_GotFocus()
    TextSelected
End Sub
Private Sub dbcManufacturer_Validate(Cancel As Boolean)
    If dbcManufacturer.Text = "" Then
        MsgBox "Manufacturer must be specified!", vbExclamation, Me.Caption
        dbcManufacturer.SetFocus
        Cancel = True
    End If
End Sub
Private Sub dbcType_GotFocus()
    TextSelected
End Sub
Private Sub Form_Load()
    Set adoConn = New ADODB.Connection
    Set rsDesignations = New ADODB.Recordset
    Set DBinfo = frmMain.DBcollection("Hobby")
    With DBinfo
        adoConn.Provider = .Provider
        adoConn.CommandTimeout = 60
        adoConn.ConnectionTimeout = 60
        adoConn.Open .PathName, .UserName, .Password
    End With
    rsDesignations.CursorLocation = adUseClient
    rsDesignations.Open "select * from [Aircraft Designations] order by Type,Number,Version", adoConn, adOpenKeyset, adLockBatchOptimistic

    rsManufacturers.CursorLocation = adUseClient
    rsManufacturers.Open "select distinct Manufacturer from [Aircraft Designations] order by Manufacturer", adoConn, adOpenStatic, adLockReadOnly
    
    rsTypes.CursorLocation = adUseClient
    rsTypes.Open "select distinct Type from [Aircraft Designations] order by Type", adoConn, adOpenStatic, adLockReadOnly
    
    Set adodcDesignations.Recordset = rsDesignations
    frmMain.BindField lblID, "ID", rsDesignations
    frmMain.BindField dbcManufacturer, "Manufacturer", rsDesignations, rsManufacturers, "Manufacturer", "Manufacturer"
    frmMain.BindField txtName, "Name", rsDesignations
    frmMain.BindField txtNumber, "Number", rsDesignations
    frmMain.BindField txtDesignation, "Designation", rsDesignations
    frmMain.BindField txtVersion, "Version", rsDesignations
    frmMain.BindField dbcType, "Type", rsDesignations, rsTypes, "Type", "Type"
    frmMain.BindField txtNotes, "Notes", rsDesignations
    frmMain.BindField txtServiceDate, "Service Date", rsDesignations
    
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
    
    If rsDesignations.EditMode <> adEditNone Then rsDesignations.CancelUpdate
    If rsDesignations.State = adStateOpen Then rsDesignations.Close
    Set rsDesignations = Nothing
    rsManufacturers.Close
    Set rsManufacturers = Nothing
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
    
    Set frmList.rsList = rsDesignations
    Set frmList.mnuList = mnuAction
    Set frmList.dgdList.DataSource = frmList.rsList
    'Set frmList.dgdList.Columns("Price").DataFormat = CurrencyFormat
    For Each Col In frmList.dgdList.Columns
        Col.Alignment = dbgGeneral
    Next Col
    
    adoConn.BeginTrans
    fTransaction = True
    frmList.Show vbModal
    adoConn.CommitTrans
    fTransaction = False
End Sub
Private Sub mnuActionNew_Click()
    mode = modeAdd
    frmMain.OpenFields Me
    adodcDesignations.Enabled = False
    rsDesignations.AddNew
    adoConn.BeginTrans
    fTransaction = True
    
    dbcManufacturer.SetFocus
End Sub
Private Sub mnuActionDelete_Click()
    mode = modeDelete
    If MsgBox("Are you sure you want to permanently delete this record...?", vbYesNo, Me.Caption) = vbYes Then
        rsDesignations.Delete
        rsDesignations.MoveNext
        If rsDesignations.EOF Then rsDesignations.MoveLast
    End If
    mode = modeDisplay
End Sub
Private Sub mnuActionModify_Click()
    mode = modeModify
    frmMain.OpenFields Me
    adodcDesignations.Enabled = False
    adoConn.BeginTrans
    fTransaction = True
    
    dbcManufacturer.SetFocus
End Sub
Private Sub mnuActionReport_Click()
    'Dim Report As New scrMusicReport
    
    'Report.Database.SetDataSource rsMusic, 3, 1
    'Set frmMain.rdcReport = Report
    'Set frmMain.frmReport = Me
    
    'frmViewReport.Show vbModal
    
    'Set Report = Nothing
End Sub
Private Sub rsDesignations_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Dim Caption As String
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    If rsDesignations.BOF And rsDesignations.EOF Then
        adodcDesignations.Caption = "No Records"
    ElseIf rsDesignations.EOF Then
        Caption = "EOF"
    ElseIf rsDesignations.BOF Then
        Caption = "BOF"
    Else
        Caption = "Reference #" & rsDesignations.Bookmark & ": " & rsDesignations("Designation") & " " & rsDesignations("Name")
    
        i = InStr(Caption, "&")
        If i > 0 Then Caption = Left(Caption, i) & "&" & Mid(Caption, i + 1)
    End If
    
    adodcDesignations.Caption = Caption
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description & " (Error " & Err.Number & ")", vbExclamation, Me.Caption
    Resume Next
End Sub
Private Sub tbDesignations_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "List"
            mnuActionList_Click
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
Private Sub txtDesignation_GotFocus()
    TextSelected
End Sub
Private Sub txtDesignation_KeyPress(KeyAscii As Integer)
    Dim Char As String
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
End Sub
Private Sub txtDesignation_Validate(Cancel As Boolean)
    If txtDesignation.Text = "" Then
        MsgBox "Designation must be specified!", vbExclamation, Me.Caption
        txtDesignation.SetFocus
        Cancel = True
    End If
End Sub
Private Sub txtServiceDate_GotFocus()
    TextSelected
End Sub
Private Sub txtVersion_GotFocus()
    TextSelected
End Sub
Private Sub txtNotes_GotFocus()
    TextSelected
End Sub
Private Sub txtNumber_GotFocus()
    TextSelected
End Sub
Private Sub txtName_GotFocus()
    TextSelected
End Sub
Private Sub txtName_Validate(Cancel As Boolean)
    If txtName.Text = "" Then
        MsgBox "Name must be specified!", vbExclamation, Me.Caption
        txtName.SetFocus
        Cancel = True
    End If
End Sub
