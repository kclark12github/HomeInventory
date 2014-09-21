VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmMusic 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Music Inventory"
   ClientHeight    =   3528
   ClientLeft      =   36
   ClientTop       =   492
   ClientWidth     =   7524
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3528
   ScaleWidth      =   7524
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   22
      Top             =   3276
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
            TextSave        =   "4:12 PM"
            Key             =   "Time"
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox chkLP 
      Alignment       =   1  'Right Justify
      Caption         =   "LP:"
      Height          =   192
      Left            =   5460
      TabIndex        =   7
      Top             =   1320
      Width           =   492
   End
   Begin VB.CheckBox chkCS 
      Alignment       =   1  'Right Justify
      Caption         =   "Cassette:"
      Height          =   192
      Left            =   4380
      TabIndex        =   6
      Top             =   1320
      Width           =   972
   End
   Begin VB.CheckBox chkCD 
      Alignment       =   1  'Right Justify
      Caption         =   "CD:"
      Height          =   192
      Left            =   3720
      TabIndex        =   5
      Top             =   1320
      Width           =   552
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   372
      Left            =   6480
      TabIndex        =   12
      Top             =   2880
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   5460
      TabIndex        =   11
      Top             =   2880
      Width           =   972
   End
   Begin MSAdodcLib.Adodc adodcMusic 
      Height          =   312
      Left            =   264
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
   Begin VB.TextBox txtAlphaSort 
      Height          =   288
      Left            =   1524
      TabIndex        =   8
      Text            =   "AlphaSort"
      Top             =   1596
      Width           =   5832
   End
   Begin VB.TextBox txtInventoried 
      Height          =   288
      Left            =   1524
      TabIndex        =   9
      Text            =   "Inventoried"
      Top             =   1932
      Width           =   1812
   End
   Begin MSDataListLib.DataCombo dbcType 
      Height          =   288
      Left            =   1524
      TabIndex        =   4
      Top             =   1260
      Width           =   2052
      _ExtentX        =   3620
      _ExtentY        =   508
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   "Type"
   End
   Begin MSDataListLib.DataCombo dbcArtist 
      Height          =   288
      Left            =   1524
      TabIndex        =   0
      Top             =   372
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
      Left            =   5304
      TabIndex        =   3
      Top             =   972
      Width           =   972
   End
   Begin VB.TextBox txtYear 
      Height          =   288
      Left            =   1524
      TabIndex        =   2
      Text            =   "Year"
      Top             =   972
      Width           =   972
   End
   Begin VB.CheckBox chkInventoried 
      Alignment       =   1  'Right Justify
      Caption         =   "Inventoried"
      Height          =   192
      Left            =   3624
      TabIndex        =   10
      Top             =   1980
      Width           =   1152
   End
   Begin VB.TextBox txtTitle 
      Height          =   288
      Left            =   1524
      TabIndex        =   1
      Text            =   "Title"
      Top             =   672
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
            Picture         =   "frmMusic.frx":0000
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMusic.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMusic.frx":0644
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMusic.frx":096C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMusic.frx":3120
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMusic.frx":3574
            Key             =   "List"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMusic.frx":4040
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMusic.frx":4494
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMusic.frx":4F60
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMusic.frx":5288
            Key             =   "Sort"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMusic.frx":56DC
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMusic.frx":5B30
            Key             =   "SQL"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMusic.frx":5F84
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
            Picture         =   "frmMusic.frx":63D8
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMusic.frx":682C
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMusic.frx":72F8
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMusic.frx":7614
            Key             =   "List"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMusic.frx":80E0
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMusic.frx":8534
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMusic.frx":ACE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMusic.frx":B13C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbAction 
      Align           =   1  'Align Top
      Height          =   288
      Left            =   0
      TabIndex        =   23
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
   Begin VB.Label lblIDlabel 
      AutoSize        =   -1  'True
      Caption         =   "ID:"
      Height          =   192
      Left            =   6804
      TabIndex        =   21
      Top             =   1980
      Width           =   192
   End
   Begin VB.Label lblDateInventoried 
      AutoSize        =   -1  'True
      Caption         =   "Date Inventoried:"
      Height          =   192
      Left            =   204
      TabIndex        =   20
      Top             =   1980
      Width           =   1212
   End
   Begin VB.Label lblAlphaSort 
      AutoSize        =   -1  'True
      Caption         =   "AlphaSort:"
      Height          =   192
      Left            =   672
      TabIndex        =   19
      Top             =   1620
      Width           =   744
   End
   Begin VB.Label lblType 
      AutoSize        =   -1  'True
      Caption         =   "Type:"
      Height          =   192
      Left            =   996
      TabIndex        =   18
      Top             =   1320
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
      Left            =   4788
      TabIndex        =   17
      Top             =   1020
      Width           =   408
   End
   Begin VB.Label lblYear 
      AutoSize        =   -1  'True
      Caption         =   "Year:"
      Height          =   192
      Left            =   1032
      TabIndex        =   16
      Top             =   1020
      Width           =   384
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Title:"
      Height          =   192
      Left            =   1068
      TabIndex        =   15
      Top             =   720
      Width           =   348
   End
   Begin VB.Label lblArtist 
      AutoSize        =   -1  'True
      Caption         =   "Artist:"
      Height          =   192
      Left            =   1032
      TabIndex        =   14
      Top             =   420
      Width           =   384
   End
   Begin VB.Label lblID 
      AutoSize        =   -1  'True
      Caption         =   "lblID"
      Height          =   192
      Left            =   7092
      TabIndex        =   13
      Top             =   1980
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
Attribute VB_Name = "frmMusic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoConn As ADODB.Connection
Dim WithEvents rsMain As ADODB.Recordset
Attribute rsMain.VB_VarHelpID = -1
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
            rsMain.CancelUpdate
            If mode = modeAdd And Not rsMain.EOF Then rsMain.MoveLast
            adoConn.RollbackTrans
            fTransaction = False
            frmMain.ProtectFields Me
            mode = modeDisplay
            adodcMusic.Enabled = True
    End Select
End Sub
Private Sub cmdOK_Click()
    Dim SaveBookmark As String
    
    Select Case mode
        Case modeDisplay
            Unload Me
        Case modeAdd, modeModify
            'Why we need to do this is buggy...
            rsMain("Artist") = dbcArtist.Text
            rsMain("Type") = dbcType.Text
            rsMain.UpdateBatch
            adoConn.CommitTrans
            fTransaction = False
            frmMain.ProtectFields Me
            mode = modeDisplay
            adodcMusic.Enabled = True
            
            mnuActionRefresh_Click
    End Select
End Sub
Private Sub dbcArtist_GotFocus()
    TextSelected
End Sub
Private Sub dbcArtist_Validate(Cancel As Boolean)
    If Not dbcArtist.Enabled Then Exit Sub
    If dbcArtist.Text = "" Then
        MsgBox "Artist must be specified!", vbExclamation, Me.Caption
        dbcArtist.SetFocus
        Cancel = True
    End If
    If rsArtists.Bookmark <> dbcArtist.SelectedItem Then rsArtists.Bookmark = dbcArtist.SelectedItem
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
        LastName = Mid(LastName, iSeparator + 1)
    End If
    
    'Check for "The" at the beginning of the title...
    Title = txtTitle.Text
    If Mid(UCase(Title), 1, 4) = "THE " Then
        Title = Mid(Title, 5) & ", " & Mid(Title, 1, 3)
    End If
    
    If UCase(dbcType.Text) = "SOUNDTRACK" Then
        DefaultAlphaSort = UCase(dbcType.Text & ": " & txtYear.Text & "; " & Title)
    ElseIf InStr(UCase(dbcArtist.Text), "TIME LIFE") > 0 Then
        DefaultAlphaSort = UCase(dbcType.Text & ": TIME LIFE: " & txtYear.Text & "; " & Title)
    Else
        DefaultAlphaSort = UCase(dbcType.Text & ": " & UCase(LastName) & ": " & txtYear.Text & "; " & Title)
    End If
End Function
Private Sub dbcType_Validate(Cancel As Boolean)
    If rsTypes.Bookmark <> dbcType.SelectedItem Then rsTypes.Bookmark = dbcType.SelectedItem
End Sub
Private Sub Form_Load()
    Set adoConn = New ADODB.Connection
    Set rsMain = New ADODB.Recordset
    Set DBinfo = frmMain.DBcollection("Music")
    With DBinfo
        adoConn.Provider = .Provider
        adoConn.CommandTimeout = 60
        adoConn.ConnectionTimeout = 60
        adoConn.Open .PathName, .UserName, .Password
    End With
    rsMain.CursorLocation = adUseClient
    rsMain.Open "select * from [Music] order by AlphaSort", adoConn, adOpenKeyset, adLockBatchOptimistic
    
    rsArtists.CursorLocation = adUseClient
    rsArtists.Open "select distinct Artist from [Music] order by Artist", adoConn, adOpenStatic, adLockReadOnly
    
    rsTypes.CursorLocation = adUseClient
    rsTypes.Open "select distinct Type from [Music] order by Type", adoConn, adOpenStatic, adLockReadOnly
    
    Set adodcMusic.Recordset = rsMain
    frmMain.BindField lblID, "ID", rsMain
    frmMain.BindField dbcArtist, "Artist", rsMain, rsArtists, "Artist", "Artist"
    frmMain.BindField txtTitle, "Title", rsMain
    frmMain.BindField txtYear, "Year", rsMain
    frmMain.BindField txtPrice, "Price", rsMain
    frmMain.BindField chkCD, "CD", rsMain
    frmMain.BindField chkCS, "CS", rsMain
    frmMain.BindField chkLP, "LP", rsMain
    frmMain.BindField txtAlphaSort, "AlphaSort", rsMain
    frmMain.BindField dbcType, "Type", rsMain, rsTypes, "Type", "Type"
    frmMain.BindField chkInventoried, "Inventoried", rsMain
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
    Else
        sbStatus.Panels("Message").Text = vbNullString
    End If
    adoConn.CommitTrans
    fTransaction = False
End Sub
Private Sub mnuActionRefresh_Click()
    Dim SaveBookmark As String
    
    On Error Resume Next
    SaveBookmark = rsMain("AlphaSort")
    rsMain.Requery
    rsMain.Find "AlphaSort='" & SQLQuote(SaveBookmark) & "'"
    rsArtists.Requery
    rsTypes.Requery
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
    Else
        sbStatus.Panels("Message").Text = vbNullString
    End If
End Sub
Private Sub mnuActionNew_Click()
    mode = modeAdd
    frmMain.OpenFields Me
    adodcMusic.Enabled = False
    rsMain.AddNew
    adoConn.BeginTrans
    fTransaction = True
    
    txtInventoried.Text = Format(Now(), "mm/dd/yyyy hh:nn AMPM")
    chkInventoried.Value = vbChecked
    chkCD.Value = vbUnchecked
    chkCS.Value = vbUnchecked
    chkLP.Value = vbUnchecked
    strDefaultAlphaSort = ""
    dbcArtist.SetFocus
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
    adodcMusic.Enabled = False
    adoConn.BeginTrans
    fTransaction = True
    
    dbcArtist.SetFocus
End Sub
Private Sub mnuActionReport_Click()
    Dim frm As Form
    Dim scrApplication As New CRAXDRT.Application
    Dim Report As New CRAXDRT.Report
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
    
    Set Report = scrApplication.OpenReport(App.Path & "\Reports\Music.rpt", crOpenReportByTempCopy)
    Report.Database.SetDataSource vRS, 3, 1
    Report.ReadRecords
    
    frmViewReport.scrViewer.ReportSource = Report
    frmViewReport.Show vbModal
    
    Set scrApplication = Nothing
    Set Report = Nothing
    vRS.Close
    Set vRS = Nothing
End Sub
Private Sub mnuActionSQL_Click()
    Load frmSQL
    Set frmSQL.cnSQL = adoConn
    frmSQL.sbStatus.Panels("DB").Text = "Music.mdb - [Music]"
    frmSQL.Show vbModal
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
        Caption = "Reference #" & rsMain.Bookmark & ": " & rsMain("ALPHASORT")
        
        i = InStr(Caption, "&")
        If i > 0 Then Caption = Left(Caption, i) & "&" & Mid(Caption, i + 1)
        If rsMain.Filter <> vbNullString And rsMain.Filter <> 0 Then
            sbStatus.Panels("Message").Text = "Filter: " & rsMain.Filter
        Else
            sbStatus.Panels("Message").Text = vbNullString
        End If
        sbStatus.Panels("Position").Text = "Record " & rsMain.Bookmark & " of " & rsMain.RecordCount
    End If
    
    adodcMusic.Caption = Caption
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
        Case "SQL"
            mnuActionSQL_Click
    End Select
End Sub
Private Sub txtAlphaSort_GotFocus()
    If txtAlphaSort.Text = "" Then
        txtAlphaSort.Text = DefaultAlphaSort
    Else
        txtAlphaSort.Text = UCase(txtAlphaSort.Text)
    End If
    TextSelected
End Sub
Private Sub txtAlphaSort_KeyPress(KeyAscii As Integer)
    KeyPressUcase KeyAscii
End Sub
Private Sub txtAlphaSort_Validate(Cancel As Boolean)
    If Not txtAlphaSort.Enabled Then Exit Sub
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
    If Not txtInventoried.Enabled Then Exit Sub
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
    ValidateCurrency txtPrice.Text, Cancel
End Sub
Private Sub txtTitle_GotFocus()
    TextSelected
End Sub
Private Sub txtTitle_Validate(Cancel As Boolean)
    If Not txtTitle.Enabled Then Exit Sub
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
    If Not txtYear.Enabled Then Exit Sub
    If txtYear.Text = "" Then
        MsgBox "Year must be specified!", vbExclamation, Me.Caption
        txtYear.SetFocus
        Cancel = True
    End If
End Sub

