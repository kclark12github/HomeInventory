VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSoftware 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Software Inventory"
   ClientHeight    =   3360
   ClientLeft      =   36
   ClientTop       =   492
   ClientWidth     =   7524
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   7524
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtISBN 
      Height          =   288
      Left            =   4800
      TabIndex        =   21
      Text            =   "ISBN"
      Top             =   660
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
      TabIndex        =   19
      Top             =   960
      Width           =   972
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   372
      Left            =   6420
      TabIndex        =   18
      Top             =   2880
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   5400
      TabIndex        =   17
      Top             =   2880
      Width           =   972
   End
   Begin MSAdodcLib.Adodc adodcSoftware 
      Height          =   312
      Left            =   204
      Top             =   2400
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
      TabIndex        =   6
      Text            =   "Inventoried"
      Top             =   1932
      Width           =   1812
   End
   Begin MSDataListLib.DataCombo dbcType 
      Height          =   288
      Left            =   1464
      TabIndex        =   4
      Top             =   1272
      Width           =   2292
      _ExtentX        =   4043
      _ExtentY        =   508
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   "Type"
   End
   Begin MSDataListLib.DataCombo dbcPublisher 
      Height          =   288
      Left            =   1458
      TabIndex        =   0
      Top             =   72
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
      TabIndex        =   5
      Text            =   "CDkey"
      Top             =   1596
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
      TabIndex        =   3
      Top             =   972
      Width           =   972
   End
   Begin VB.TextBox txtVersion 
      Height          =   288
      Left            =   1458
      TabIndex        =   2
      Text            =   "Version"
      Top             =   672
      Width           =   1512
   End
   Begin VB.CheckBox chkCataloged 
      Alignment       =   1  'Right Justify
      Caption         =   "Cataloged"
      Height          =   192
      Left            =   3558
      TabIndex        =   7
      Top             =   1980
      Width           =   1152
   End
   Begin VB.TextBox txtTitle 
      Height          =   288
      Left            =   1458
      TabIndex        =   1
      Text            =   "Title"
      Top             =   372
      Width           =   5892
   End
   Begin MSDataListLib.DataCombo dbcPlatform 
      Height          =   288
      Left            =   4800
      TabIndex        =   23
      Top             =   1260
      Width           =   2292
      _ExtentX        =   4043
      _ExtentY        =   508
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   "Platform"
   End
   Begin VB.Label lblPlatform 
      AutoSize        =   -1  'True
      Caption         =   "Platform:"
      Height          =   192
      Left            =   4104
      TabIndex        =   24
      Top             =   1308
      Width           =   624
   End
   Begin VB.Label lblISBN 
      AutoSize        =   -1  'True
      Caption         =   "ISBN:"
      Height          =   192
      Left            =   4320
      TabIndex        =   22
      Top             =   708
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
      TabIndex        =   20
      Top             =   1008
      Width           =   360
   End
   Begin VB.Label lblIDlabel 
      AutoSize        =   -1  'True
      Caption         =   "ID:"
      Height          =   192
      Left            =   6744
      TabIndex        =   16
      Top             =   1980
      Width           =   192
   End
   Begin VB.Label lblDateInventoried 
      AutoSize        =   -1  'True
      Caption         =   "Date Inventoried:"
      Height          =   192
      Left            =   144
      TabIndex        =   15
      Top             =   1980
      Width           =   1212
   End
   Begin VB.Label lblCDkey 
      AutoSize        =   -1  'True
      Caption         =   "CD Key:"
      Height          =   192
      Left            =   828
      TabIndex        =   14
      Top             =   1620
      Width           =   576
   End
   Begin VB.Label lblType 
      AutoSize        =   -1  'True
      Caption         =   "Type:"
      Height          =   192
      Left            =   984
      TabIndex        =   13
      Top             =   1320
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
      TabIndex        =   12
      Top             =   1020
      Width           =   456
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      Caption         =   "Version:"
      Height          =   192
      Left            =   762
      TabIndex        =   11
      Top             =   720
      Width           =   588
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Title:"
      Height          =   192
      Left            =   1002
      TabIndex        =   10
      Top             =   420
      Width           =   348
   End
   Begin VB.Label lblPublisher 
      AutoSize        =   -1  'True
      Caption         =   "Publisher:"
      Height          =   192
      Left            =   642
      TabIndex        =   9
      Top             =   120
      Width           =   708
   End
   Begin VB.Label lblID 
      AutoSize        =   -1  'True
      Caption         =   "lblID"
      Height          =   192
      Left            =   7032
      TabIndex        =   8
      Top             =   1980
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
Attribute VB_Name = "frmSoftware"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoConn As adodb.Connection
Dim WithEvents rsSoftware As adodb.Recordset
Attribute rsSoftware.VB_VarHelpID = -1
Dim rsPublishers As New adodb.Recordset
Dim rsTypes As New adodb.Recordset
Dim rsPlatforms As New adodb.Recordset
Dim mode As ActionMode
Dim fTransaction As Boolean
Dim DBinfo As DataBaseInfo
Private Sub cmdCancel_Click()
    Select Case mode
        Case modeDisplay
            Unload Me
        Case modeAdd, modeModify
            rsSoftware.CancelUpdate
            adoConn.RollbackTrans
            fTransaction = False
            frmMain.ProtectFields Me
            mode = modeDisplay
            adodcSoftware.Enabled = True
    End Select
End Sub
Private Sub cmdOK_Click()
    Select Case mode
        Case modeDisplay
            Unload Me
        Case modeAdd, modeModify
            rsSoftware.UpdateBatch
            adoConn.CommitTrans
            fTransaction = False
            frmMain.ProtectFields Me
            mode = modeDisplay
            adodcSoftware.Enabled = True
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
    Set adoConn = New adodb.Connection
    Set rsSoftware = New adodb.Recordset
    Set DBinfo = frmMain.DBcollection("Software")
    With DBinfo
        adoConn.Provider = .Provider
        adoConn.CommandTimeout = 60
        adoConn.ConnectionTimeout = 60
        adoConn.Open .PathName, .UserName, .Password
    End With
    rsSoftware.CursorLocation = adUseClient
    rsSoftware.Open "select ID,Type,Publisher,Title,Version,ISBN,Platform,Value,Cost,CDkey,DateInventoried,Cataloged from [Software] order by Type,Title", adoConn, adOpenKeyset, adLockBatchOptimistic

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
    
    If rsSoftware.EditMode <> adEditNone Then rsSoftware.CancelUpdate
    If rsSoftware.State = adStateOpen Then rsSoftware.Close
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
    Set frmList.mnuList = mnuAction
    Set frmList.dgdList.DataSource = frmList.rsList
    Set frmList.dgdList.Columns("Value").DataFormat = CurrencyFormat
    Set frmList.dgdList.Columns("Cost").DataFormat = CurrencyFormat
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
    adodcSoftware.Enabled = False
    rsSoftware.AddNew
    adoConn.BeginTrans
    fTransaction = True
    
    txtInventoried.Text = Format(Now(), "mm/dd/yyyy hh:nn AMPM")
    chkCataloged.Value = vbChecked
    strDefaultAlphaSort = ""
    dbcAuthor.SetFocus
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
    
    dbcAuthor.SetFocus
End Sub
Private Sub mnuActionReport_Click()
    Dim Report As New scrBooksReport
    
    Report.Database.SetDataSource rsSoftware, 3, 1
    Set frmMain.rdcReport = Report
    Set frmMain.frmReport = Me
    
    frmViewReport.Show vbModal
    
    Set Report = Nothing
End Sub
Private Sub rsSoftware_MoveComplete(ByVal adReason As adodb.EventReasonEnum, ByVal pError As adodb.Error, adStatus As adodb.EventStatusEnum, ByVal pRecordset As adodb.Recordset)
    On Error GoTo ErrorHandler
    If rsSoftware.EOF Then rsSoftware.MoveLast
    If rsSoftware.BOF Then rsSoftware.MoveFirst
    If rsSoftware.BOF And rsSoftware.EOF Then adodcSoftware.Caption = "No Records"
    'adodcSoftware.Caption = "Books ID #" & rsSoftware("ID")
    adodcSoftware.Caption = "Reference #" & rsSoftware.Bookmark & ": " & rsSoftware("Type") & "; " & rsSoftware("Title")
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description & " (Error " & Err.Number & ")", vbExclamation, Me.Caption
    Resume Next
End Sub
Private Sub txtCDkey_GotFocus()
    TextSelected
End Sub
Private Sub txtCDkey_Validate(Cancel As Boolean)
    If txtCDkey.Text = "" Then
        MsgBox "CDkey must be specified!", vbExclamation, Me.Caption
        txtCDkey.SetFocus
        Cancel = True
    End If
End Sub
Private Sub txtCost_GotFocus()
    TextSelected
End Sub
Private Sub txtCost_Validate(Cancel As Boolean)
    If txtCost.Text = "" Then
        MsgBox "Cost must be specified!", vbExclamation, Me.Caption
        txtCost.SetFocus
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
Private Sub txtValue_Validate(Cancel As Boolean)
    If txtValue.Text = "" Then
        MsgBox "Value must be specified!", vbExclamation, Me.Caption
        txtValue.SetFocus
        Cancel = True
    End If
End Sub
Private Sub txtVersion_GotFocus()
    TextSelected
End Sub

