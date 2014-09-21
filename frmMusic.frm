VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmMusic 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Music Inventory"
   ClientHeight    =   3048
   ClientLeft      =   36
   ClientTop       =   492
   ClientWidth     =   7524
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3048
   ScaleWidth      =   7524
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkLP 
      Alignment       =   1  'Right Justify
      Caption         =   "LP:"
      Height          =   192
      Left            =   5400
      TabIndex        =   21
      Top             =   1020
      Width           =   492
   End
   Begin VB.CheckBox chkCS 
      Alignment       =   1  'Right Justify
      Caption         =   "Cassette:"
      Height          =   192
      Left            =   4320
      TabIndex        =   20
      Top             =   1020
      Width           =   972
   End
   Begin VB.CheckBox chkCD 
      Alignment       =   1  'Right Justify
      Caption         =   "CD:"
      Height          =   192
      Left            =   3660
      TabIndex        =   19
      Top             =   1020
      Width           =   552
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   372
      Left            =   6420
      TabIndex        =   18
      Top             =   2580
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   5400
      TabIndex        =   17
      Top             =   2580
      Width           =   972
   End
   Begin MSAdodcLib.Adodc adodcMusic 
      Height          =   312
      Left            =   204
      Top             =   2100
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
   Begin VB.TextBox txtAlphaSort 
      Height          =   288
      Left            =   1458
      TabIndex        =   5
      Text            =   "AlphaSort"
      Top             =   1296
      Width           =   5832
   End
   Begin VB.TextBox txtInventoried 
      Height          =   288
      Left            =   1458
      TabIndex        =   6
      Text            =   "Inventoried"
      Top             =   1632
      Width           =   1812
   End
   Begin MSDataListLib.DataCombo dbcType 
      Height          =   288
      Left            =   1464
      TabIndex        =   4
      Top             =   972
      Width           =   2052
      _ExtentX        =   3620
      _ExtentY        =   508
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   "Type"
   End
   Begin MSDataListLib.DataCombo dbcArtist 
      Height          =   288
      Left            =   1458
      TabIndex        =   0
      Top             =   72
      Width           =   5892
      _ExtentX        =   10393
      _ExtentY        =   508
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   "Artist"
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
      Left            =   5244
      TabIndex        =   3
      Top             =   672
      Width           =   972
   End
   Begin VB.TextBox txtYear 
      Height          =   288
      Left            =   1458
      TabIndex        =   2
      Text            =   "Year"
      Top             =   672
      Width           =   972
   End
   Begin VB.CheckBox chkInventoried 
      Alignment       =   1  'Right Justify
      Caption         =   "Inventoried"
      Height          =   192
      Left            =   3558
      TabIndex        =   7
      Top             =   1680
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
   Begin VB.Label lblIDlabel 
      AutoSize        =   -1  'True
      Caption         =   "ID:"
      Height          =   192
      Left            =   6744
      TabIndex        =   16
      Top             =   1680
      Width           =   192
   End
   Begin VB.Label lblDateInventoried 
      AutoSize        =   -1  'True
      Caption         =   "Date Inventoried:"
      Height          =   192
      Left            =   144
      TabIndex        =   15
      Top             =   1680
      Width           =   1212
   End
   Begin VB.Label lblAlphaSort 
      AutoSize        =   -1  'True
      Caption         =   "AlphaSort:"
      Height          =   192
      Left            =   612
      TabIndex        =   14
      Top             =   1320
      Width           =   744
   End
   Begin VB.Label lblType 
      AutoSize        =   -1  'True
      Caption         =   "Type:"
      Height          =   192
      Left            =   936
      TabIndex        =   13
      Top             =   1020
      Width           =   420
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
      Left            =   4728
      TabIndex        =   12
      Top             =   720
      Width           =   408
   End
   Begin VB.Label lblYear 
      AutoSize        =   -1  'True
      Caption         =   "Year:"
      Height          =   192
      Left            =   966
      TabIndex        =   11
      Top             =   720
      Width           =   384
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
   Begin VB.Label lblArtist 
      AutoSize        =   -1  'True
      Caption         =   "Artist:"
      Height          =   192
      Left            =   864
      TabIndex        =   9
      Top             =   120
      Width           =   384
   End
   Begin VB.Label lblID 
      AutoSize        =   -1  'True
      Caption         =   "lblID"
      Height          =   192
      Left            =   7032
      TabIndex        =   8
      Top             =   1680
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
Attribute VB_Name = "frmMusic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoConn As ADODB.Connection
Dim WithEvents rsMusic As ADODB.Recordset
Attribute rsMusic.VB_VarHelpID = -1
Dim rsArtists As New ADODB.Recordset
Dim rsTypes As New ADODB.Recordset
Dim strDefaultAlphaSort As String
Dim mode As ActionMode
Dim fTransaction As Boolean
Dim DBinfo As DataBaseInfo
Private Sub cmdCancel_Click()
    Select Case mode
        Case modeDisplay
            Unload Me
        Case modeAdd, modeModify
            rsMusic.CancelUpdate
            adoConn.RollbackTrans
            fTransaction = False
            frmMain.ProtectFields Me
            mode = modeDisplay
            adodcMusic.Enabled = True
    End Select
End Sub
Private Sub cmdOK_Click()
    Select Case mode
        Case modeDisplay
            Unload Me
        Case modeAdd, modeModify
            rsMusic.UpdateBatch
            adoConn.CommitTrans
            fTransaction = False
            frmMain.ProtectFields Me
            mode = modeDisplay
            adodcMusic.Enabled = True
    End Select
End Sub
Private Sub dbcArtist_GotFocus()
    TextSelected
End Sub
Private Sub dbcArtist_Validate(Cancel As Boolean)
    If dbcArtist.Text = "" Then
        MsgBox "Artist must be specified!", vbExclamation, Me.Caption
        dbcArtist.SetFocus
        Cancel = True
    End If
End Sub
Private Sub dbcType_GotFocus()
    TextSelected
End Sub
Private Function DefaultAlphaSort() As String
    Dim LastName As String
    Dim Title As String
    Dim iAnd As Integer
    Dim iAmpersand As Integer
    Dim iComma As Integer
    Dim iSemiColon As Integer
    Dim iSeparator As Integer
    
    'Start with the Artist's last name...
    LastName = dbcArtist.Text
    iAnd = InStr(dbcArtist.Text, " and ")
    iAmpersand = InStr(dbcArtist.Text, " & ")
    iComma = InStr(dbcArtist.Text, ",")
    iSemiColon = InStr(dbcArtist.Text, ";")
    
    If iComma > 0 Then
        'Assume the comma separates Artists, and...
        iSeparator = iComma
    ElseIf iSemiColon > 0 Then
        'Assume the semicolon separates Artists, and...
        iSeparator = iSemiColon
    ElseIf iAnd > 0 Then
        'Assume the "and" separates Artists, and...
        iSeparator = iAnd
    ElseIf iAmpersand > 0 Then
        'Assume the "&" separates Artists, and...
        iSeparator = iAmpersand
    End If
    
    If iSeparator > 0 Then
        '...take the first Artist...
        LastName = Mid(LastName, 1, iSeparator - 1)
    End If
        
    'OK, we have a single person's name (theoretically)...
    'Grab the last word on the line and assume it's his last name...
    If InStr(LastName, " ") Then
        iSeparator = InStrRev(LastName, " ", Len(LastName))
        LastName = Mid(LastName, 5)
    End If
    
    'Check for "The" at the beginning of the title...
    Title = txtTitle.Text
    If Mid(UCase(Title), 1, 4) = "THE " Then
        Title = Mid(Title, 5) & ", " & Mid(Title, 1, 3)
    End If
    
    DefaultAlphaSort = dbcType.Text & ": " & UCase(LastName) & ": " & txtYear.Text & "; " & Title
End Function
Private Sub Form_Load()
    Set adoConn = New ADODB.Connection
    Set rsMusic = New ADODB.Recordset
    Set DBinfo = frmMain.DBcollection("Music")
    With DBinfo
        adoConn.Provider = .Provider
        adoConn.CommandTimeout = 60
        adoConn.ConnectionTimeout = 60
        adoConn.Open .PathName, .UserName, .Password
    End With
    rsMusic.CursorLocation = adUseClient
    rsMusic.Open "select * from [Music] order by AlphaSort", adoConn, adOpenKeyset, adLockBatchOptimistic
    
    rsArtists.CursorLocation = adUseClient
    rsArtists.Open "select distinct Artist from [Music] order by Artist", adoConn, adOpenStatic, adLockReadOnly
    
    rsTypes.CursorLocation = adUseClient
    rsTypes.Open "select distinct Type from [Music] order by Type", adoConn, adOpenStatic, adLockReadOnly
    
    Set adodcMusic.Recordset = rsMusic
    frmMain.BindField lblID, "ID", rsMusic
    frmMain.BindField dbcArtist, "Artist", rsMusic, rsArtists, "Artist", "Artist"
    frmMain.BindField txtTitle, "Title", rsMusic
    frmMain.BindField txtYear, "Year", rsMusic
    frmMain.BindField txtPrice, "Price", rsMusic
    frmMain.BindField chkCD, "CD", rsMusic
    frmMain.BindField chkCS, "CS", rsMusic
    frmMain.BindField chkLP, "LP", rsMusic
    frmMain.BindField txtAlphaSort, "AlphaSort", rsMusic
    frmMain.BindField dbcType, "Type", rsMusic, rsTypes, "Type", "Type"
    frmMain.BindField chkInventoried, "Inventoried", rsMusic
    frmMain.BindField txtInventoried, "DateInventoried", rsMusic

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
    
    If rsMusic.EditMode <> adEditNone Then rsMusic.CancelUpdate
    If rsMusic.State = adStateOpen Then rsMusic.Close
    Set rsMusic = Nothing
    rsArtists.Close
    Set rsArtists = Nothing
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
    
    Set frmList.rsList = rsMusic
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
    adodcMusic.Enabled = False
    rsMusic.AddNew
    adoConn.BeginTrans
    fTransaction = True
    
    txtInventoried.Text = Format(Now(), "mm/dd/yyyy hh:nn AMPM")
    chkCataloged.Value = vbChecked
    strDefaultAlphaSort = ""
    dbcArtist.SetFocus
End Sub
Private Sub mnuActionDelete_Click()
    mode = modeDelete
    If MsgBox("Are you sure you want to permanently delete this record...?", vbYesNo, Me.Caption) = vbYes Then
        rsMusic.Delete
        rsMusic.MoveNext
        If rsMusic.EOF Then rsMusic.MoveLast
    End If
    mode = modeDisplay
End Sub
Private Sub mnuActionModify_Click()
    mode = modeModify
    frmMain.OpenFields Me
    adodcMusic.Enabled = False
    adoConn.BeginTrans
    fTransaction = True
    
    dbcArtist.SetFocus
End Sub
Private Sub mnuActionReport_Click()
    Dim Report As New scrMusicReport
    
    Report.Database.SetDataSource rsMusic, 3, 1
    Set frmMain.rdcReport = Report
    Set frmMain.frmReport = Me
    
    frmViewReport.Show vbModal
    
    Set Report = Nothing
End Sub
Private Sub rsMusic_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    On Error GoTo ErrorHandler
    If rsMusic.EOF Then rsMusic.MoveLast
    If rsMusic.BOF Then rsMusic.MoveFirst
    If rsMusic.BOF And rsMusic.EOF Then adodcMusic.Caption = "No Records"
    'adodcMusic.Caption = "Music ID #" & rsMusic("ID")
    adodcMusic.Caption = "Reference #" & rsMusic.Bookmark & ": " & rsMusic("ALPHASORT")
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description & " (Error " & Err.Number & ")", vbExclamation, Me.Caption
    Resume Next
End Sub
Private Sub txtAlphaSort_GotFocus()
    If txtAlphaSort.Text = "" Then txtAlphaSort.Text = DefaultAlphaSort
    TextSelected
End Sub
Private Sub txtAlphaSort_Validate(Cancel As Boolean)
    If txtAlphaSort.Text = "" Then
        MsgBox "AlphaSort must be specified!", vbExclamation, Me.Caption
        txtAlphaSort.SetFocus
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
Private Sub txtMisc_GotFocus()
    TextSelected
End Sub
Private Sub txtPrice_GotFocus()
    TextSelected
End Sub
Private Sub txtPrice_Validate(Cancel As Boolean)
    If txtPrice.Text = "" Then
        MsgBox "Price must be specified!", vbExclamation, Me.Caption
        txtPrice.SetFocus
        Cancel = True
    End If
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
Private Sub txtYear_GotFocus()
    TextSelected
End Sub
Private Sub txtYear_Validate(Cancel As Boolean)
    If txtYear.Text = "" Then
        MsgBox "Year must be specified!", vbExclamation, Me.Caption
        txtYear.SetFocus
        Cancel = True
    End If
End Sub

