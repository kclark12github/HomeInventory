VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmBooks 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Book Inventory"
   ClientHeight    =   3696
   ClientLeft      =   36
   ClientTop       =   492
   ClientWidth     =   7524
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3696
   ScaleWidth      =   7524
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   372
      Left            =   6420
      TabIndex        =   20
      Top             =   3240
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   5400
      TabIndex        =   19
      Top             =   3240
      Width           =   972
   End
   Begin MSAdodcLib.Adodc adodcBooks 
      Height          =   312
      Left            =   204
      Top             =   2760
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
      TabIndex        =   6
      Text            =   "AlphaSort"
      Top             =   1956
      Width           =   5832
   End
   Begin VB.TextBox txtInventoried 
      Height          =   288
      Left            =   1458
      TabIndex        =   7
      Text            =   "Inventoried"
      Top             =   2292
      Width           =   1812
   End
   Begin MSDataListLib.DataCombo dbcSubject 
      Height          =   288
      Left            =   1464
      TabIndex        =   4
      Top             =   1272
      Width           =   5892
      _ExtentX        =   10393
      _ExtentY        =   508
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   "Subject"
   End
   Begin MSDataListLib.DataCombo dbcAuthor 
      Height          =   288
      Left            =   1458
      TabIndex        =   0
      Top             =   72
      Width           =   5892
      _ExtentX        =   10393
      _ExtentY        =   508
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   "Author"
   End
   Begin VB.TextBox txtMisc 
      Height          =   288
      Left            =   1458
      TabIndex        =   5
      Text            =   "Misc"
      Top             =   1596
      Width           =   3792
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
      Left            =   1458
      TabIndex        =   3
      Top             =   972
      Width           =   972
   End
   Begin VB.TextBox txtISBN 
      Height          =   288
      Left            =   1458
      TabIndex        =   2
      Text            =   "ISBN"
      Top             =   672
      Width           =   1512
   End
   Begin VB.CheckBox chkCataloged 
      Alignment       =   1  'Right Justify
      Caption         =   "Cataloged"
      Height          =   192
      Left            =   3558
      TabIndex        =   8
      Top             =   2340
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
      TabIndex        =   18
      Top             =   2340
      Width           =   192
   End
   Begin VB.Label lblDateInventoried 
      AutoSize        =   -1  'True
      Caption         =   "Date Inventoried:"
      Height          =   192
      Left            =   144
      TabIndex        =   17
      Top             =   2340
      Width           =   1212
   End
   Begin VB.Label lblAlphaSort 
      AutoSize        =   -1  'True
      Caption         =   "AlphaSort:"
      Height          =   192
      Left            =   612
      TabIndex        =   16
      Top             =   1980
      Width           =   744
   End
   Begin VB.Label lblMisc 
      AutoSize        =   -1  'True
      Caption         =   "Miscellaneous:"
      Height          =   192
      Left            =   276
      TabIndex        =   15
      Top             =   1620
      Width           =   1080
   End
   Begin VB.Label lblSubject 
      AutoSize        =   -1  'True
      Caption         =   "Subject:"
      Height          =   192
      Left            =   780
      TabIndex        =   14
      Top             =   1320
      Width           =   576
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
      Left            =   942
      TabIndex        =   13
      Top             =   1020
      Width           =   408
   End
   Begin VB.Label lblISBN 
      AutoSize        =   -1  'True
      Caption         =   "ISBN:"
      Height          =   192
      Left            =   942
      TabIndex        =   12
      Top             =   720
      Width           =   408
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Title:"
      Height          =   192
      Left            =   1002
      TabIndex        =   11
      Top             =   420
      Width           =   348
   End
   Begin VB.Label lblAuthor 
      AutoSize        =   -1  'True
      Caption         =   "Author:"
      Height          =   192
      Left            =   858
      TabIndex        =   10
      Top             =   120
      Width           =   492
   End
   Begin VB.Label lblID 
      AutoSize        =   -1  'True
      Caption         =   "lblID"
      Height          =   192
      Left            =   7032
      TabIndex        =   9
      Top             =   2340
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
Attribute VB_Name = "frmBooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoConn As ADODB.Connection
Dim WithEvents rsBooks As ADODB.Recordset
Attribute rsBooks.VB_VarHelpID = -1
Dim rsAuthors As New ADODB.Recordset
Dim rsSubjects As New ADODB.Recordset
Dim strDefaultAlphaSort As String
Dim mode As ActionMode
Dim fTransaction As Boolean
Dim DBinfo As DataBaseInfo
Private Sub cmdCancel_Click()
    Select Case mode
        Case modeDisplay
            Unload Me
        Case modeAdd, modeModify
            rsBooks.CancelUpdate
            adoConn.RollbackTrans
            fTransaction = False
            frmMain.ProtectFields Me
            mode = modeDisplay
            adodcBooks.Enabled = True
    End Select
End Sub
Private Sub cmdOK_Click()
    Select Case mode
        Case modeDisplay
            Unload Me
        Case modeAdd, modeModify
            rsBooks.UpdateBatch
            adoConn.CommitTrans
            fTransaction = False
            frmMain.ProtectFields Me
            mode = modeDisplay
            adodcBooks.Enabled = True
    End Select
End Sub
Private Sub dbcAuthor_GotFocus()
    TextSelected
End Sub
Private Sub dbcAuthor_Validate(Cancel As Boolean)
    If dbcAuthor.Text = "" Then
        MsgBox "Author must be specified!", vbExclamation, Me.Caption
        dbcAuthor.SetFocus
        Cancel = True
    End If
End Sub
Private Sub dbcSubject_GotFocus()
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
    
    'Start with the Author's last name...
    LastName = dbcAuthor.Text
    iAnd = InStr(dbcAuthor.Text, " and ")
    iAmpersand = InStr(dbcAuthor.Text, " & ")
    iComma = InStr(dbcAuthor.Text, ",")
    iSemiColon = InStr(dbcAuthor.Text, ";")
    
    If iComma > 0 Then
        'Assume the comma separates authors, and...
        iSeparator = iComma
    ElseIf iSemiColon > 0 Then
        'Assume the semicolon separates authors, and...
        iSeparator = iSemiColon
    ElseIf iAnd > 0 Then
        'Assume the "and" separates authors, and...
        iSeparator = iAnd
    ElseIf iAmpersand > 0 Then
        'Assume the "&" separates authors, and...
        iSeparator = iAmpersand
    End If
    
    If iSeparator > 0 Then
        '...take the first Author...
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
    
    DefaultAlphaSort = UCase(LastName) & ": " & Title
End Function
Private Sub Form_Load()
    Set adoConn = New ADODB.Connection
    Set rsBooks = New ADODB.Recordset
    Set DBinfo = frmMain.DBcollection("Books")
    With DBinfo
        adoConn.Provider = .Provider
        adoConn.CommandTimeout = 60
        adoConn.ConnectionTimeout = 60
        adoConn.Open .PathName, .UserName, .Password
    End With
    rsBooks.CursorLocation = adUseClient
    rsBooks.Open "select * from [Master Book List] order by AlphaSort", adoConn, adOpenKeyset, adLockBatchOptimistic

    rsAuthors.CursorLocation = adUseClient
    rsAuthors.Open "select distinct Author from [Master Book List] order by Author", adoConn, adOpenStatic, adLockReadOnly
    
    rsSubjects.CursorLocation = adUseClient
    rsSubjects.Open "select distinct Subject from [Master Book List] order by Subject", adoConn, adOpenStatic, adLockReadOnly
    
    Set adodcBooks.Recordset = rsBooks
    Set lblID.DataSource = rsBooks
    lblID.DataField = "ID"
    Set dbcAuthor.DataSource = rsBooks
    dbcAuthor.DataField = "Author"
    Set dbcAuthor.RowSource = rsAuthors
    dbcAuthor.BoundColumn = "Author"
    dbcAuthor.ListField = "Author"
    Set txtTitle.DataSource = rsBooks
    txtTitle.DataField = "Title"
    Set txtISBN.DataSource = rsBooks
    txtISBN.DataField = "ISBN"
    Set txtPrice.DataSource = rsBooks
    txtPrice.DataField = "Price"
    Set txtAlphaSort.DataSource = rsBooks
    txtAlphaSort.DataField = "Alphasort"
    Set dbcSubject.DataSource = rsBooks
    dbcSubject.DataField = "Subject"
    Set dbcSubject.RowSource = rsSubjects
    dbcSubject.BoundColumn = "Subject"
    dbcSubject.ListField = "Subject"
    Set txtMisc.DataSource = rsBooks
    txtMisc.DataField = "Misc"
    Set chkCataloged.DataSource = rsBooks
    chkCataloged.DataField = "Cataloged"
    Set txtInventoried.DataSource = rsBooks
    txtInventoried.DataField = "Inventoried"

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
    
    If rsBooks.EditMode <> adEditNone Then rsBooks.CancelUpdate
    If rsBooks.State = adStateOpen Then rsBooks.Close
    Set rsBooks = Nothing
    rsAuthors.Close
    Set rsAuthors = Nothing
    rsSubjects.Close
    Set rsSubjects = Nothing
    
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
    
    Set frmList.rsList = rsBooks
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
    adodcBooks.Enabled = False
    rsBooks.AddNew
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
        rsBooks.Delete
        rsBooks.MoveNext
        If rsBooks.EOF Then rsBooks.MoveLast
    End If
    mode = modeDisplay
End Sub
Private Sub mnuActionModify_Click()
    mode = modeModify
    frmMain.OpenFields Me
    adodcBooks.Enabled = False
    adoConn.BeginTrans
    fTransaction = True
    
    dbcAuthor.SetFocus
End Sub
Private Sub mnuActionReport_Click()
    Dim Report As New scrBooksReport
    
    Report.Database.SetDataSource rsBooks, 3, 1
    Set frmMain.rdcReport = Report
    Set frmMain.frmReport = Me
    
    frmViewReport.Show vbModal
    
    Set Report = Nothing
End Sub
Private Sub rsBooks_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    On Error GoTo ErrorHandler
    If rsBooks.EOF Then rsBooks.MoveLast
    If rsBooks.BOF Then rsBooks.MoveFirst
    If rsBooks.BOF And rsBooks.EOF Then adodcBooks.Caption = "No Records"
    'adodcBooks.Caption = "Books ID #" & rsBooks("ID")
    adodcBooks.Caption = "Books #" & rsBooks.Bookmark & ": " & rsBooks("ALPHASORT")
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description & " (Error " & Err.Number & ")", vbExclamation, Me.Caption
    Resume Next
End Sub
Private Sub txtAlphaSort_GotFocus()
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
Private Sub txtISBN_GotFocus()
    TextSelected
End Sub
Private Sub txtISBN_Validate(Cancel As Boolean)
    If txtISBN.Text = "" Then
        MsgBox "ISBN must be specified!", vbExclamation, Me.Caption
        txtISBN.SetFocus
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
    If txtAlphaSort.Text = "" Then txtAlphaSort.Text = DefaultAlphaSort
End Sub
