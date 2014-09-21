VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmTools 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tools"
   ClientHeight    =   3468
   ClientLeft      =   36
   ClientTop       =   492
   ClientWidth     =   7524
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3468
   ScaleWidth      =   7524
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   19
      Top             =   3216
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
            TextSave        =   "8:45 PM"
            Key             =   "Time"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtCount 
      Height          =   288
      Left            =   5934
      TabIndex        =   6
      Text            =   "Count"
      Top             =   1560
      Width           =   972
   End
   Begin VB.TextBox txtReference 
      Height          =   288
      Left            =   5934
      TabIndex        =   4
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
      TabIndex        =   8
      Top             =   2820
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   5466
      TabIndex        =   7
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
      TabIndex        =   5
      Text            =   "Inventoried"
      Top             =   1572
      Width           =   1812
   End
   Begin MSDataListLib.DataCombo dbcManufacturer 
      Height          =   288
      Left            =   1536
      TabIndex        =   2
      Top             =   960
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
      Left            =   1530
      TabIndex        =   1
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
   Begin MSDataListLib.DataCombo dbcCatalog 
      Height          =   288
      Left            =   1530
      TabIndex        =   3
      Top             =   1260
      Width           =   3252
      _ExtentX        =   5736
      _ExtentY        =   508
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   "Catalog"
   End
   Begin MSComctlLib.ImageList imlSmall 
      Left            =   480
      Top             =   2700
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
            Picture         =   "frmTools.frx":0000
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTools.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTools.frx":0644
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTools.frx":096C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTools.frx":3120
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTools.frx":3574
            Key             =   "List"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTools.frx":4040
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTools.frx":4494
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTools.frx":4F60
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTools.frx":5288
            Key             =   "Sort"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTools.frx":56DC
            Key             =   "Filter"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlLarge 
      Left            =   60
      Top             =   2700
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
            Picture         =   "frmTools.frx":5B30
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTools.frx":5F84
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTools.frx":6A50
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTools.frx":6D6C
            Key             =   "List"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTools.frx":7838
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTools.frx":7C8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTools.frx":A440
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTools.frx":A894
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbAction 
      Align           =   1  'Align Top
      Height          =   288
      Left            =   0
      TabIndex        =   18
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
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      Caption         =   "Count:"
      Height          =   192
      Left            =   5346
      TabIndex        =   17
      Top             =   1608
      Width           =   444
   End
   Begin VB.Label lblCatalog 
      AutoSize        =   -1  'True
      Caption         =   "Catalog:"
      Height          =   192
      Left            =   822
      TabIndex        =   16
      Top             =   1308
      Width           =   600
   End
   Begin VB.Label lblReference 
      AutoSize        =   -1  'True
      Caption         =   "Reference:"
      Height          =   192
      Left            =   5046
      TabIndex        =   15
      Top             =   1308
      Width           =   792
   End
   Begin VB.Label lblIDlabel 
      AutoSize        =   -1  'True
      Caption         =   "ID:"
      Height          =   192
      Left            =   6810
      TabIndex        =   14
      Top             =   1920
      Width           =   192
   End
   Begin VB.Label lblDateInventoried 
      AutoSize        =   -1  'True
      Caption         =   "Date Inventoried:"
      Height          =   192
      Left            =   216
      TabIndex        =   13
      Top             =   1620
      Width           =   1212
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
      Left            =   1014
      TabIndex        =   12
      Top             =   720
      Width           =   408
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   192
      Left            =   942
      TabIndex        =   11
      Top             =   420
      Width           =   480
   End
   Begin VB.Label lblManufacturer 
      AutoSize        =   -1  'True
      Caption         =   "Manufacturer:"
      Height          =   192
      Left            =   462
      TabIndex        =   10
      Top             =   1020
      Width           =   960
   End
   Begin VB.Label lblID 
      AutoSize        =   -1  'True
      Caption         =   "lblID"
      Height          =   192
      Left            =   7098
      TabIndex        =   9
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
   End
End
Attribute VB_Name = "frmTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoConn As ADODB.Connection
Dim WithEvents rsMain As ADODB.Recordset
Attribute rsMain.VB_VarHelpID = -1
Dim rsManufacturers As New ADODB.Recordset
Dim rsCatalogs As New ADODB.Recordset
Dim mode As ActionMode
Dim fTransaction As Boolean
Dim DBinfo As DataBaseInfo
Private Sub cmdCancel_Click()
    Select Case mode
        Case modeDisplay
            Unload Me
        Case modeAdd, modeModify
            rsMain.CancelUpdate
            If mode = modeAdd Then rsMain.MoveLast
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
            rsMain("Manufacturer") = dbcManufacturer.BoundText
            rsMain("Catalog") = dbcCatalog.BoundText
            rsMain.UpdateBatch
            adoConn.CommitTrans
            fTransaction = False
            frmMain.ProtectFields Me
            mode = modeDisplay
            adodcHobby.Enabled = True
            
            mnuActionRefresh_Click
    End Select
End Sub
Private Sub dbcCatalog_GotFocus()
    TextSelected
End Sub
Private Sub dbcCatalog_Validate(Cancel As Boolean)
    If dbcCatalog.Text = vbNullString Then dbcCatalog.Text = "Unknown"
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
Private Sub Form_Load()
    Set adoConn = New ADODB.Connection
    Set rsMain = New ADODB.Recordset
    Set DBinfo = frmMain.DBcollection("Hobby")
    With DBinfo
        adoConn.Provider = .Provider
        adoConn.CommandTimeout = 60
        adoConn.ConnectionTimeout = 60
        adoConn.Open .PathName, .UserName, .Password
    End With
    rsMain.CursorLocation = adUseClient
    rsMain.Open "select * from [Tools] order by Manufacturer,Name", adoConn, adOpenKeyset, adLockBatchOptimistic
    
    rsManufacturers.CursorLocation = adUseClient
    rsManufacturers.Open "select distinct Manufacturer from [Tools] order by Manufacturer", adoConn, adOpenStatic, adLockReadOnly
    
    rsCatalogs.CursorLocation = adUseClient
    rsCatalogs.Open "select distinct Catalog from [Tools] order by Catalog", adoConn, adOpenStatic, adLockReadOnly
    
    Set adodcHobby.Recordset = rsMain
    frmMain.BindField lblID, "ID", rsMain
    frmMain.BindField dbcManufacturer, "Manufacturer", rsMain, rsManufacturers, "Manufacturer", "Manufacturer"
    frmMain.BindField txtName, "Name", rsMain
    frmMain.BindField txtPrice, "Price", rsMain
    frmMain.BindField txtReference, "Reference", rsMain
    frmMain.BindField dbcCatalog, "Catalog", rsMain, rsCatalogs, "Catalog", "Catalog"
    frmMain.BindField txtCount, "Count", rsMain
    frmMain.BindField txtInventoried, "DateInventoried", rsMain

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
    
    If Not rsMain.EOF Then
        If rsMain.EditMode <> adEditNone Then rsMain.CancelUpdate
    End If
    If (rsMain.State And adStateOpen) = adStateOpen Then rsMain.Close
    Set rsMain = Nothing
    rsManufacturers.Close
    Set rsManufacturers = Nothing
    rsCatalogs.Close
    Set rsCatalogs = Nothing
    
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
    
    Set frmList.rsList = rsMain
    
    adoConn.BeginTrans
    fTransaction = True
    frmList.Show vbModal
    If rsMain.Filter <> vbNullString And rsMain.Filter <> 0 Then
        sbStatus.Panels("Message").Text = "Filter: " & rsMain.Filter
    End If
    adoConn.CommitTrans
    fTransaction = False
End Sub
Private Sub mnuActionRefresh_Click()
    Dim SaveBookmark As String
    
    On Error Resume Next
    SaveBookmark = rsMain("Reference")
    rsMain.Requery
    rsMain.Find "Reference='" & SQLQuote(SaveBookmark) & "'"
    rsManufacturers.Requery
    rsCatalogs.Requery
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
    
    Set frmFilter.RS = rsMain
    frmFilter.Show vbModal
    If rsMain.Filter <> vbNullString And rsMain.Filter <> 0 Then
        sbStatus.Panels("Message").Text = "Filter: " & rsMain.Filter
    End If
End Sub
Private Sub mnuActionNew_Click()
    mode = modeAdd
    frmMain.OpenFields Me
    adodcHobby.Enabled = False
    rsMain.AddNew
    adoConn.BeginTrans
    fTransaction = True
    
    txtInventoried.Text = Format(Now(), "mm/dd/yyyy hh:nn AMPM")
    txtName.SetFocus
End Sub
Private Sub mnuActionDelete_Click()
    mode = modeDelete
    If MsgBox("Are you sure you want to permanently delete this record...?", vbYesNo, Me.Caption) = vbYes Then
        rsMain.Delete
        rsMain.MoveNext
        If rsMain.EOF Then rsMain.MoveLast
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
    Dim frm As Form
    Dim Report As New scrToolsReport
    Dim vRS As ADODB.Recordset
    
    MakeVirtualRecordset adoConn, rsMain, vRS
    
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
Private Sub rsMain_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Dim Caption As String
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    If rsMain.BOF And rsMain.EOF Then
        Caption = "No Records"
    ElseIf rsMain.EOF Then
        Caption = "EOF"
    ElseIf rsMain.BOF Then
        Caption = "BOF"
    Else
        Caption = "Reference #" & rsMain.Bookmark & ": " & rsMain("Manufacturer") & " " & rsMain("Name")
        
        i = InStr(Caption, "&")
        If i > 0 Then Caption = Left(Caption, i) & "&" & Mid(Caption, i + 1)
        If rsMain.Filter <> vbNullString And rsMain.Filter <> 0 Then
            sbStatus.Panels("Message").Text = "Filter: " & rsMain.Filter
        End If
        sbStatus.Panels("Position").Text = "Record " & rsMain.Bookmark & " of " & rsMain.RecordCount
    End If
    
    adodcHobby.Caption = Caption
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
Private Sub txtPrice_Validate(Cancel As Boolean)
    ValidateCurrency txtPrice.Text, Cancel
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
