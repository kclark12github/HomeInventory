VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmDetailSets 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detail Sets"
   ClientHeight    =   3276
   ClientLeft      =   36
   ClientTop       =   492
   ClientWidth     =   7524
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3276
   ScaleWidth      =   7524
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCount 
      Height          =   288
      Left            =   5934
      TabIndex        =   8
      Text            =   "Count"
      Top             =   1560
      Width           =   972
   End
   Begin VB.TextBox txtReference 
      Height          =   288
      Left            =   5934
      TabIndex        =   6
      Text            =   "Reference"
      Top             =   1260
      Width           =   1452
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   372
      Left            =   6486
      TabIndex        =   11
      Top             =   2820
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   5466
      TabIndex        =   10
      Top             =   2820
      Width           =   972
   End
   Begin MSAdodcLib.Adodc adodcHobby 
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
      Left            =   1530
      TabIndex        =   9
      Text            =   "Inventoried"
      Top             =   1872
      Width           =   1812
   End
   Begin MSDataListLib.DataCombo dbcManufacturer 
      Height          =   288
      Left            =   1530
      TabIndex        =   4
      Top             =   972
      Width           =   5892
      _ExtentX        =   10393
      _ExtentY        =   508
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   "Manufacturer"
   End
   Begin VB.TextBox txtPrice 
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
      Left            =   5934
      TabIndex        =   3
      Top             =   660
      Width           =   972
   End
   Begin VB.TextBox txtName 
      Height          =   288
      Left            =   1530
      TabIndex        =   0
      Text            =   "Name"
      Top             =   372
      Width           =   5892
   End
   Begin MSComctlLib.ImageList imlSmall 
      Left            =   486
      Top             =   2820
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
            Picture         =   "frmDetailSets.frx":0000
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetailSets.frx":031C
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetailSets.frx":0638
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetailSets.frx":0A8C
            Key             =   "List"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetailSets.frx":1558
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetailSets.frx":2224
            Key             =   "New2"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetailSets.frx":2CF0
            Key             =   "Delete2"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetailSets.frx":37BC
            Key             =   "Modify2"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetailSets.frx":4288
            Key             =   "New"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetailSets.frx":4D54
            Key             =   "Report"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlLarge 
      Left            =   66
      Top             =   2820
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
            Picture         =   "frmDetailSets.frx":51A8
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetailSets.frx":55FC
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetailSets.frx":5918
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetailSets.frx":5C34
            Key             =   "New2"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetailSets.frx":6700
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetailSets.frx":73CC
            Key             =   "List"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetailSets.frx":7E98
            Key             =   "Delete2"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetailSets.frx":8964
            Key             =   "Modify2"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetailSets.frx":9430
            Key             =   "New"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetailSets.frx":9EFC
            Key             =   "Delete"
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo dbcCatalog 
      Height          =   288
      Left            =   1530
      TabIndex        =   5
      Top             =   1260
      Width           =   3252
      _ExtentX        =   5736
      _ExtentY        =   508
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   "Catalog"
   End
   Begin MSDataListLib.DataCombo dbcNation 
      Height          =   288
      Left            =   1530
      TabIndex        =   7
      Top             =   1560
      Width           =   3252
      _ExtentX        =   5736
      _ExtentY        =   508
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   "Nation"
   End
   Begin MSDataListLib.DataCombo dbcScale 
      Height          =   288
      Left            =   1530
      TabIndex        =   1
      Top             =   672
      Width           =   1092
      _ExtentX        =   1926
      _ExtentY        =   508
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   "Scale"
   End
   Begin MSDataListLib.DataCombo dbcType 
      Height          =   288
      Left            =   3510
      TabIndex        =   2
      Top             =   660
      Width           =   1512
      _ExtentX        =   2667
      _ExtentY        =   508
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   "Type"
   End
   Begin MSComctlLib.Toolbar tbHobby 
      Align           =   1  'Align Top
      Height          =   288
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   7524
      _ExtentX        =   13272
      _ExtentY        =   508
      ButtonWidth     =   1439
      ButtonHeight    =   466
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
   Begin VB.Label lblType 
      AutoSize        =   -1  'True
      Caption         =   "Type:"
      Height          =   192
      Left            =   3006
      TabIndex        =   23
      Top             =   720
      Width           =   420
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      Caption         =   "Count:"
      Height          =   192
      Left            =   5346
      TabIndex        =   22
      Top             =   1608
      Width           =   444
   End
   Begin VB.Label lblCatalog 
      AutoSize        =   -1  'True
      Caption         =   "Catalog:"
      Height          =   192
      Left            =   822
      TabIndex        =   21
      Top             =   1308
      Width           =   600
   End
   Begin VB.Label lblReference 
      AutoSize        =   -1  'True
      Caption         =   "Reference:"
      Height          =   192
      Left            =   5046
      TabIndex        =   20
      Top             =   1308
      Width           =   792
   End
   Begin VB.Label lblScale 
      AutoSize        =   -1  'True
      Caption         =   "Scale:"
      Height          =   192
      Left            =   966
      TabIndex        =   19
      Top             =   720
      Width           =   456
   End
   Begin VB.Label lblIDlabel 
      AutoSize        =   -1  'True
      Caption         =   "ID:"
      Height          =   192
      Left            =   6810
      TabIndex        =   18
      Top             =   1920
      Width           =   192
   End
   Begin VB.Label lblDateInventoried 
      AutoSize        =   -1  'True
      Caption         =   "Date Inventoried:"
      Height          =   192
      Left            =   210
      TabIndex        =   17
      Top             =   1920
      Width           =   1212
   End
   Begin VB.Label lblNation 
      AutoSize        =   -1  'True
      Caption         =   "Nation:"
      Height          =   192
      Left            =   918
      TabIndex        =   16
      Top             =   1608
      Width           =   504
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
      Height          =   192
      Left            =   5394
      TabIndex        =   15
      Top             =   720
      Width           =   408
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   192
      Left            =   942
      TabIndex        =   14
      Top             =   420
      Width           =   480
   End
   Begin VB.Label lblManufacturer 
      AutoSize        =   -1  'True
      Caption         =   "Manufacturer:"
      Height          =   192
      Left            =   462
      TabIndex        =   13
      Top             =   1020
      Width           =   960
   End
   Begin VB.Label lblID 
      AutoSize        =   -1  'True
      Caption         =   "lblID"
      Height          =   192
      Left            =   7098
      TabIndex        =   12
      Top             =   1920
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
Attribute VB_Name = "frmDetailSets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoConn As ADODB.Connection
Dim WithEvents rsDetailSets As ADODB.Recordset
Attribute rsDetailSets.VB_VarHelpID = -1
Dim rsManufacturers As New ADODB.Recordset
Dim rsCatalogs As New ADODB.Recordset
Dim rsScales As New ADODB.Recordset
Dim rsNations As New ADODB.Recordset
Dim rsTypes As New ADODB.Recordset
Dim mode As ActionMode
Dim fTransaction As Boolean
Dim DBinfo As DataBaseInfo
Private Sub cmdCancel_Click()
    Select Case mode
        Case modeDisplay
            Unload Me
        Case modeAdd, modeModify
            rsDetailSets.CancelUpdate
            adoConn.RollbackTrans
            fTransaction = False
            frmMain.ProtectFields Me
            mode = modeDisplay
            adodcHobby.Enabled = True
    End Select
End Sub
Private Sub cmdOK_Click()
    Dim SaveBookmark As String
    
    Select Case mode
        Case modeDisplay
            Unload Me
        Case modeAdd, modeModify
            'Why we need to do this is buggy...
            rsDetailSets("Manufacturer") = dbcManufacturer.BoundText
            rsDetailSets("Catalog") = dbcCatalog.BoundText
            rsDetailSets.UpdateBatch
            adoConn.CommitTrans
            fTransaction = False
            frmMain.ProtectFields Me
            mode = modeDisplay
            adodcHobby.Enabled = True
            
            SaveBookmark = rsDetailSets("Reference")
            rsDetailSets.Requery
            rsDetailSets.Find "Reference='" & SaveBookmark & "'"
            rsManufacturers.Requery
            rsCatalogs.Requery
            rsScales.Requery
            rsNations.Requery
            rsTypes.Requery
    End Select
End Sub
Private Sub dbcCatalog_GotFocus()
    TextSelected
End Sub
Private Sub dbcCatalog_Validate(Cancel As Boolean)
    If Trim(dbcCatalog.Text) = vbNullString Then dbcCatalog.Text = "Unknown"
    If rsCatalogs.Bookmark <> dbcCatalog.SelectedItem Then rsCatalogs.Bookmark = dbcCatalog.SelectedItem
End Sub
Private Sub dbcManufacturer_GotFocus()
    TextSelected
End Sub
Private Sub dbcManufacturer_Validate(Cancel As Boolean)
    If dbcManufacturer.Text = vbNullString Then
        MsgBox "Manufacturer must be specified!", vbExclamation, Me.Caption
        dbcManufacturer.SetFocus
        Cancel = True
    End If
    If rsManufacturers.Bookmark <> dbcManufacturer.SelectedItem Then rsManufacturers.Bookmark = dbcManufacturer.SelectedItem
End Sub
Private Sub dbcNation_GotFocus()
    TextSelected
End Sub
Private Sub dbcNation_Validate(Cancel As Boolean)
    If dbcNation.Text = vbNullString Then
        MsgBox "Nation must be specified!", vbExclamation, Me.Caption
        dbcNation.SetFocus
        Cancel = True
    End If
    If rsNations.Bookmark <> dbcNation.SelectedItem Then rsNations.Bookmark = dbcNation.SelectedItem
End Sub
Private Sub dbcScale_GotFocus()
    TextSelected
End Sub
Private Sub dbcScale_Validate(Cancel As Boolean)
    If dbcScale.Text = vbNullString Then dbcScale.Text = "Unknown"
    If rsScales.Bookmark <> dbcScale.SelectedItem Then rsScales.Bookmark = dbcScale.SelectedItem
End Sub
Private Sub dbcType_GotFocus()
    TextSelected
End Sub
Private Sub dbcType_Validate(Cancel As Boolean)
    If dbcType.Text = vbNullString Then
        MsgBox "Type must be specified!", vbExclamation, Me.Caption
        dbcType.SetFocus
        Cancel = True
    End If
    If rsTypes.Bookmark <> dbcType.SelectedItem Then rsTypes.Bookmark = dbcType.SelectedItem
End Sub
Private Sub Form_Load()
    Set adoConn = New ADODB.Connection
    Set rsDetailSets = New ADODB.Recordset
    Set DBinfo = frmMain.DBcollection("Hobby")
    With DBinfo
        adoConn.Provider = .Provider
        adoConn.CommandTimeout = 60
        adoConn.ConnectionTimeout = 60
        adoConn.Open .PathName, .UserName, .Password
    End With
    rsDetailSets.CursorLocation = adUseClient
    rsDetailSets.Open "select * from [Detail Sets] order by Scale,Name", adoConn, adOpenKeyset, adLockBatchOptimistic
    
    rsManufacturers.CursorLocation = adUseClient
    rsManufacturers.Open "select distinct Manufacturer from [Detail Sets] order by Manufacturer", adoConn, adOpenStatic, adLockReadOnly
    
    rsCatalogs.CursorLocation = adUseClient
    rsCatalogs.Open "select distinct Catalog from [Detail Sets] order by Catalog", adoConn, adOpenStatic, adLockReadOnly
    
    rsScales.CursorLocation = adUseClient
    rsScales.Open "select distinct Scale from [Detail Sets] order by Scale", adoConn, adOpenStatic, adLockReadOnly
    
    rsNations.CursorLocation = adUseClient
    rsNations.Open "select distinct Nation from [Detail Sets] order by Nation", adoConn, adOpenStatic, adLockReadOnly
    
    rsTypes.CursorLocation = adUseClient
    rsTypes.Open "select distinct Type from [Detail Sets] order by Type", adoConn, adOpenStatic, adLockReadOnly
    
    Set adodcHobby.Recordset = rsDetailSets
    frmMain.BindField lblID, "ID", rsDetailSets
    frmMain.BindField dbcManufacturer, "Manufacturer", rsDetailSets, rsManufacturers, "Manufacturer", "Manufacturer"
    frmMain.BindField txtName, "Name", rsDetailSets
    frmMain.BindField txtPrice, "Price", rsDetailSets
    frmMain.BindField dbcScale, "Scale", rsDetailSets, rsScales, "Scale", "Scale"
    frmMain.BindField txtReference, "Reference", rsDetailSets
    frmMain.BindField dbcCatalog, "Catalog", rsDetailSets, rsCatalogs, "Catalog", "Catalog"
    frmMain.BindField dbcNation, "Nation", rsDetailSets, rsNations, "Nation", "Nation"
    frmMain.BindField dbcType, "Type", rsDetailSets, rsTypes, "Type", "Type"
    frmMain.BindField txtCount, "Count", rsDetailSets
    frmMain.BindField txtInventoried, "DateInventoried", rsDetailSets

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
    
    If Not rsDetailSets.EOF Then
        If rsDetailSets.EditMode <> adEditNone Then rsDetailSets.CancelUpdate
    End If
    If (rsDetailSets.State And adStateOpen) = adStateOpen Then rsDetailSets.Close
    Set rsDetailSets = Nothing
    rsManufacturers.Close
    Set rsManufacturers = Nothing
    rsCatalogs.Close
    Set rsCatalogs = Nothing
    rsScales.Close
    Set rsScales = Nothing
    rsNations.Close
    Set rsNations = Nothing
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
    
    Set frmList.rsList = rsDetailSets
    Set frmList.mnuList = mnuAction
    Set frmList.dgdList.DataSource = frmList.rsList
    Set frmList.dgdList.Columns("Price").DataFormat = CurrencyFormat
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
    adodcHobby.Enabled = False
    rsDetailSets.AddNew
    adoConn.BeginTrans
    fTransaction = True
    
    txtInventoried.Text = Format(Now(), "mm/dd/yyyy hh:nn AMPM")
    txtName.SetFocus
End Sub
Private Sub mnuActionDelete_Click()
    mode = modeDelete
    If MsgBox("Are you sure you want to permanently delete this record...?", vbYesNo, Me.Caption) = vbYes Then
        rsDetailSets.Delete
        rsDetailSets.MoveNext
        If rsDetailSets.EOF Then rsDetailSets.MoveLast
    End If
    mode = modeDisplay
End Sub
Private Sub mnuActionModify_Click()
    mode = modeModify
    frmMain.OpenFields Me
    adodcHobby.Enabled = False
    adoConn.BeginTrans
    fTransaction = True
    
    txtName.SetFocus
End Sub
Private Sub mnuActionReport_Click()
    'Dim Report As New scrHobbyReport
    
    'Report.Database.SetDataSource rsDetailSets, 3, 1
    'Set frmMain.rdcReport = Report
    'Set frmMain.frmReport = Me
    
    'frmViewReport.Show vbModal
    
    'Set Report = Nothing
End Sub
Private Sub rsDetailSets_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Dim Caption As String
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    If rsDetailSets.BOF And rsDetailSets.EOF Then
        Caption = "No Records"
    ElseIf rsDetailSets.EOF Then
        Caption = "EOF"
    ElseIf rsDetailSets.BOF Then
        Caption = "BOF"
    Else
        If IsNumeric(rsDetailSets("Scale")) Then
            Caption = "Reference #" & rsDetailSets.Bookmark & ": 1/" & rsDetailSets("Scale") & " Scale; " & rsDetailSets("Name")
        Else
            Caption = "Reference #" & rsDetailSets.Bookmark & ": " & rsDetailSets("Scale") & " Scale; " & rsDetailSets("Name")
        End If
        
        i = InStr(Caption, "&")
        If i > 0 Then Caption = Left(Caption, i) & "&" & Mid(Caption, i + 1)
    End If
    
    adodcHobby.Caption = Caption
    Exit Sub

ErrorHandler:
    MsgBox Err.Description & " (Error " & Err.Number & ")", vbExclamation, Me.Caption
    Resume Next
End Sub
Private Sub tbHobby_ButtonClick(ByVal Button As MSComctlLib.Button)
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
Private Sub txtCount_GotFocus()
    TextSelected
End Sub
Private Sub txtCount_Validate(Cancel As Boolean)
    If txtCount.Text = vbNullString Then txtCount.Text = 1
End Sub
Private Sub txtInventoried_GotFocus()
    TextSelected
End Sub
Private Sub txtName_GotFocus()
    TextSelected
End Sub
Private Sub txtName_Validate(Cancel As Boolean)
    If txtName.Text = vbNullString Then
        MsgBox "Name must be specified!", vbExclamation, Me.Caption
        txtName.SetFocus
        Cancel = True
    End If
End Sub
Private Sub txtPrice_GotFocus()
    TextSelected
End Sub
Private Sub txtPrice_KeyPress(KeyAscii As Integer)
    If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then
        If KeyAscii <> Asc(".") Then
            KeyAscii = 0    'Cancel the character.
            Beep            'Sound error signal.
        End If
    End If
End Sub
Private Sub txtPrice_Validate(Cancel As Boolean)
    If txtPrice.Text = vbNullString Then txtPrice.Text = Format(0, "Currency")
    If Not IsNumeric(txtPrice.Text) Then
        MsgBox "Invalid price entered.", vbExclamation, Me.Caption
        TextSelected
        Cancel = True
    End If
End Sub
Private Sub txtReference_GotFocus()
    TextSelected
End Sub
Private Sub txtReference_KeyPress(KeyAscii As Integer)
    Dim Char As String
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
End Sub
Private Sub txtReference_Validate(Cancel As Boolean)
    If txtReference.Text = vbNullString Then txtReference.Text = "Unknown"
End Sub
