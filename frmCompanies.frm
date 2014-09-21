VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCompanies 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hobby Companies"
   ClientHeight    =   4476
   ClientLeft      =   36
   ClientTop       =   492
   ClientWidth     =   7524
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4476
   ScaleWidth      =   7524
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
   Begin MSAdodcLib.Adodc adodcCompanies 
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
      Left            =   480
      Top             =   3900
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
            Picture         =   "frmCompanies.frx":0008
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompanies.frx":0324
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompanies.frx":0640
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompanies.frx":0A94
            Key             =   "List"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompanies.frx":1560
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompanies.frx":222C
            Key             =   "New2"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompanies.frx":2CF8
            Key             =   "Delete2"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompanies.frx":37C4
            Key             =   "Modify2"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompanies.frx":4290
            Key             =   "New"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompanies.frx":4D5C
            Key             =   "Report"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlLarge 
      Left            =   60
      Top             =   3900
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
            Picture         =   "frmCompanies.frx":51B0
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompanies.frx":5604
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompanies.frx":5920
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompanies.frx":5C3C
            Key             =   "New2"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompanies.frx":6708
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompanies.frx":73D4
            Key             =   "List"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompanies.frx":7EA0
            Key             =   "Delete2"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompanies.frx":896C
            Key             =   "Modify2"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompanies.frx":9438
            Key             =   "New"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompanies.frx":9F04
            Key             =   "Delete"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbCompanies 
      Align           =   1  'Align Top
      Height          =   288
      Left            =   0
      TabIndex        =   19
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
   Begin VB.Label lblWebSite 
      AutoSize        =   -1  'True
      Caption         =   "Web Site:"
      Height          =   192
      Left            =   360
      TabIndex        =   20
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
Attribute VB_Name = "frmCompanies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoConn As ADODB.Connection
Dim WithEvents rsCompanies As ADODB.Recordset
Attribute rsCompanies.VB_VarHelpID = -1
Dim rsProductTypes As New ADODB.Recordset
Dim mode As ActionMode
Dim fTransaction As Boolean
Dim DBinfo As DataBaseInfo
Private Sub cmdCancel_Click()
    Select Case mode
        Case modeDisplay
            Unload Me
        Case modeAdd, modeModify
            rsCompanies.CancelUpdate
            adoConn.RollbackTrans
            fTransaction = False
            frmMain.ProtectFields Me
            mode = modeDisplay
            adodcCompanies.Enabled = True
    End Select
End Sub
Private Sub cmdOK_Click()
    Dim SaveBookmark As String
    
    Select Case mode
        Case modeDisplay
            Unload Me
        Case modeAdd, modeModify
            'Why we need to do this is buggy...
            rsCompanies("ProductType") = dbcProductType.Text
            rsCompanies.UpdateBatch
            adoConn.CommitTrans
            fTransaction = False
            frmMain.ProtectFields Me
            mode = modeDisplay
            adodcCompanies.Enabled = True
            
            SaveBookmark = rsCompanies("Code")
            rsCompanies.Requery
            rsCompanies.Find "Code='" & SaveBookmark & "'"
            rsProductTypes.Requery
    End Select
End Sub
Private Sub dbcProductType_GotFocus()
    TextSelected
End Sub
Private Sub dbcProductType_Validate(Cancel As Boolean)
    If rsProductTypes.Bookmark <> dbcProductType.SelectedItem Then rsProductTypes.Bookmark = dbcProductType.SelectedItem
End Sub
Private Sub Form_Load()
    Set adoConn = New ADODB.Connection
    Set rsCompanies = New ADODB.Recordset
    Set DBinfo = frmMain.DBcollection("Hobby")
    With DBinfo
        adoConn.Provider = .Provider
        adoConn.CommandTimeout = 60
        adoConn.ConnectionTimeout = 60
        adoConn.Open .PathName, .UserName, .Password
    End With
    rsCompanies.CursorLocation = adUseClient
    rsCompanies.Open "select * from [Companies] order by Code", adoConn, adOpenKeyset, adLockBatchOptimistic
    
    rsProductTypes.CursorLocation = adUseClient
    rsProductTypes.Open "select distinct ProductType from [Companies] order by ProductType", adoConn, adOpenStatic, adLockReadOnly
    
    Set adodcCompanies.Recordset = rsCompanies
    frmMain.BindField lblID, "ID", rsCompanies
    frmMain.BindField txtName, "Name", rsCompanies
    frmMain.BindField txtShortName, "ShortName", rsCompanies
    frmMain.BindField txtCode, "Code", rsCompanies
    frmMain.BindField txtAccount, "Account", rsCompanies
    frmMain.BindField txtPhone, "Phone", rsCompanies
    frmMain.BindField txtAddress, "Address", rsCompanies
    frmMain.BindField txtWebSite, "Website", rsCompanies
    frmMain.BindField dbcProductType, "ProductType", rsCompanies, rsProductTypes, "ProductType", "ProductType"

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
    
    If rsCompanies.EditMode <> adEditNone Then rsCompanies.CancelUpdate
    If rsCompanies.State = adStateOpen Then rsCompanies.Close
    Set rsCompanies = Nothing
    rsProductTypes.Close
    Set rsProductTypes = Nothing
    
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
    
    Set frmList.rsList = rsCompanies
    Set frmList.mnuList = mnuAction
    Set frmList.dgdList.DataSource = frmList.rsList
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
    adodcCompanies.Enabled = False
    rsCompanies.AddNew
    adoConn.BeginTrans
    fTransaction = True
    
    txtName.SetFocus
End Sub
Private Sub mnuActionDelete_Click()
    mode = modeDelete
    If MsgBox("Are you sure you want to permanently delete this record...?", vbYesNo, Me.Caption) = vbYes Then
        rsCompanies.Delete
        rsCompanies.MoveNext
        If rsCompanies.EOF Then rsCompanies.MoveLast
    End If
    mode = modeDisplay
End Sub
Private Sub mnuActionModify_Click()
    mode = modeModify
    frmMain.OpenFields Me
    adodcCompanies.Enabled = False
    adoConn.BeginTrans
    fTransaction = True
    
    txtName.SetFocus
End Sub
Private Sub mnuActionReport_Click()
    'Dim Report As New scrCompaniesReport
    
    'Report.Database.SetDataSource rsCompanies, 3, 1
    'Set frmMain.rdcReport = Report
    'Set frmMain.frmReport = Me
    
    'frmViewReport.Show vbModal
    
    'Set Report = Nothing
End Sub
Private Sub rsCompanies_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Dim Caption As String
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    If rsCompanies.BOF And rsCompanies.EOF Then
        adodcCompanies.Caption = "No Records"
    ElseIf rsCompanies.EOF Then
        Caption = "EOF"
    ElseIf rsCompanies.BOF Then
        Caption = "BOF"
    Else
        Caption = "Reference #" & rsCompanies.Bookmark & ": " & rsCompanies("Code")
        
        i = InStr(Caption, "&")
        If i > 0 Then Caption = Left(Caption, i) & "&" & Mid(Caption, i + 1)
    End If
    
    adodcCompanies.Caption = Caption
    Exit Sub

ErrorHandler:
    MsgBox Err.Description & " (Error " & Err.Number & ")", vbExclamation, Me.Caption
    Resume Next
End Sub
Private Sub tbCompanies_ButtonClick(ByVal Button As MSComctlLib.Button)
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
Private Sub txtShortName_GotFocus()
    TextSelected
End Sub
Private Sub txtShortName_Validate(Cancel As Boolean)
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
    Dim Char As String
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
End Sub
Private Sub txtCode_Validate(Cancel As Boolean)
    If txtCode.Text = "" Then
        MsgBox "Code must be specified!", vbExclamation, Me.Caption
        txtCode.SetFocus
        Cancel = True
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