VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmBlueAngelsHistory 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Blue Angels History"
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
            Picture         =   "frmBlueAngelsHistory.frx":0010
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlueAngelsHistory.frx":032C
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlueAngelsHistory.frx":0648
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlueAngelsHistory.frx":0A9C
            Key             =   "List"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlueAngelsHistory.frx":1568
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlueAngelsHistory.frx":2234
            Key             =   "New2"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlueAngelsHistory.frx":2D00
            Key             =   "Delete2"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlueAngelsHistory.frx":37CC
            Key             =   "Modify2"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlueAngelsHistory.frx":4298
            Key             =   "New"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlueAngelsHistory.frx":4D64
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
            Picture         =   "frmBlueAngelsHistory.frx":51B8
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlueAngelsHistory.frx":560C
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlueAngelsHistory.frx":5928
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlueAngelsHistory.frx":5C44
            Key             =   "New2"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlueAngelsHistory.frx":6710
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlueAngelsHistory.frx":73DC
            Key             =   "List"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlueAngelsHistory.frx":7EA8
            Key             =   "Delete2"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlueAngelsHistory.frx":8974
            Key             =   "Modify2"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlueAngelsHistory.frx":9440
            Key             =   "New"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlueAngelsHistory.frx":9F0C
            Key             =   "Delete"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbHobby 
      Align           =   1  'Align Top
      Height          =   288
      Left            =   0
      TabIndex        =   12
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
Attribute VB_Name = "frmBlueAngelsHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoConn As ADODB.Connection
Dim WithEvents rsBlueAngelsHistory As ADODB.Recordset
Attribute rsBlueAngelsHistory.VB_VarHelpID = -1
Dim mode As ActionMode
Dim fTransaction As Boolean
Dim DBinfo As DataBaseInfo
Private Sub cmdCancel_Click()
    Select Case mode
        Case modeDisplay
            Unload Me
        Case modeAdd, modeModify
            rsBlueAngelsHistory.CancelUpdate
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
            rsBlueAngelsHistory.UpdateBatch
            adoConn.CommitTrans
            fTransaction = False
            frmMain.ProtectFields Me
            mode = modeDisplay
            adodcHobby.Enabled = True
            
            SaveBookmark = rsBlueAngelsHistory("Aircraft Type")
            rsBlueAngelsHistory.Requery
            rsBlueAngelsHistory.Find "Aircraft Type='" & SaveBookmark & "'"
    End Select
End Sub
Private Sub Form_Load()
    Set adoConn = New ADODB.Connection
    Set rsBlueAngelsHistory = New ADODB.Recordset
    Set DBinfo = frmMain.DBcollection("Hobby")
    With DBinfo
        adoConn.Provider = .Provider
        adoConn.CommandTimeout = 60
        adoConn.ConnectionTimeout = 60
        adoConn.Open .PathName, .UserName, .Password
    End With
    rsBlueAngelsHistory.CursorLocation = adUseClient
    rsBlueAngelsHistory.Open "select * from [Blue Angels History] order by Dates", adoConn, adOpenKeyset, adLockBatchOptimistic
    
    Set adodcHobby.Recordset = rsBlueAngelsHistory
    frmMain.BindField lblID, "ID", rsBlueAngelsHistory
    frmMain.BindField txtAircraftType, "Aircraft Type", rsBlueAngelsHistory
    frmMain.BindField txtDates, "Dates", rsBlueAngelsHistory
    frmMain.BindField txtKits, "Kits", rsBlueAngelsHistory
    frmMain.BindField txtDecalSets, "Decal Sets", rsBlueAngelsHistory

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
    
    If Not rsBlueAngelsHistory.EOF Then
        If rsBlueAngelsHistory.EditMode <> adEditNone Then rsBlueAngelsHistory.CancelUpdate
    End If
    If (rsBlueAngelsHistory.State And adStateOpen) = adStateOpen Then rsBlueAngelsHistory.Close
    Set rsBlueAngelsHistory = Nothing
    
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
    
    Set frmList.rsList = rsBlueAngelsHistory
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
    adodcHobby.Enabled = False
    rsBlueAngelsHistory.AddNew
    adoConn.BeginTrans
    fTransaction = True
    
    txtAircraftType.SetFocus
End Sub
Private Sub mnuActionDelete_Click()
    mode = modeDelete
    If MsgBox("Are you sure you want to permanently delete this record...?", vbYesNo, Me.Caption) = vbYes Then
        rsBlueAngelsHistory.Delete
        rsBlueAngelsHistory.MoveNext
        If rsBlueAngelsHistory.EOF Then rsBlueAngelsHistory.MoveLast
    End If
    mode = modeDisplay
End Sub
Private Sub mnuActionModify_Click()
    mode = modeModify
    frmMain.OpenFields Me
    adodcHobby.Enabled = False
    adoConn.BeginTrans
    fTransaction = True
    
    txtAircraftType.SetFocus
End Sub
Private Sub mnuActionReport_Click()
    'Dim Report As New scrHobbyReport
    
    'Report.Database.SetDataSource rsBlueAngelsHistory, 3, 1
    'Set frmMain.rdcReport = Report
    'Set frmMain.frmReport = Me
    
    'frmViewReport.Show vbModal
    
    'Set Report = Nothing
End Sub
Private Sub rsBlueAngelsHistory_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Dim Caption As String
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    If rsBlueAngelsHistory.BOF And rsBlueAngelsHistory.EOF Then
        Caption = "No Records"
    ElseIf rsBlueAngelsHistory.EOF Then
        Caption = "EOF"
    ElseIf rsBlueAngelsHistory.BOF Then
        Caption = "BOF"
    Else
        Caption = "Reference #" & rsBlueAngelsHistory.Bookmark & ": " & rsBlueAngelsHistory("Dates") & ": " & rsBlueAngelsHistory("Aircraft Type")
        
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
Private Sub txtAircraftType_GotFocus()
    TextSelected
End Sub
Private Sub txtAircraftType_Validate(Cancel As Boolean)
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

