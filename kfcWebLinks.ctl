VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.UserControl kfcWebLinks 
   ClientHeight    =   4668
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7692
   ControlContainer=   -1  'True
   DefaultCancel   =   -1  'True
   ScaleHeight     =   4668
   ScaleWidth      =   7692
   ToolboxBitmap   =   "kfcWebLinks.ctx":0000
   Begin VB.Frame frameLayout 
      Caption         =   "Button Layout"
      Height          =   3372
      Left            =   180
      TabIndex        =   22
      Top             =   420
      Width           =   3252
      Begin ComctlLib.TreeView tvwDB 
         Height          =   3072
         Left            =   60
         TabIndex        =   0
         Top             =   240
         Width           =   3132
         _ExtentX        =   5525
         _ExtentY        =   5419
         _Version        =   327682
         HideSelection   =   0   'False
         Indentation     =   441
         Sorted          =   -1  'True
         Style           =   7
         ImageList       =   "imlIcons"
         Appearance      =   1
         OLEDragMode     =   1
         OLEDropMode     =   1
      End
   End
   Begin VB.Frame frameDetail 
      Caption         =   "Detail"
      Height          =   3372
      Left            =   3480
      TabIndex        =   10
      Top             =   420
      Width           =   4092
      Begin VB.CommandButton cmdHyperlink 
         Caption         =   "&Hyperlink"
         Height          =   372
         Left            =   1560
         TabIndex        =   7
         Top             =   2760
         Visible         =   0   'False
         Width           =   1152
      End
      Begin VB.TextBox txtLabel 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   288
         Left            =   1080
         TabIndex        =   1
         Top             =   264
         Width           =   2892
      End
      Begin VB.TextBox txtURL 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   264
         Left            =   1080
         TabIndex        =   2
         Top             =   600
         Width           =   2892
      End
      Begin VB.TextBox txtParentID 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   264
         Left            =   1080
         TabIndex        =   14
         Top             =   1860
         Visible         =   0   'False
         Width           =   2892
      End
      Begin VB.TextBox txtButtonLabel 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   264
         Left            =   1080
         TabIndex        =   13
         Top             =   1560
         Visible         =   0   'False
         Width           =   2892
      End
      Begin VB.CheckBox chkHasMembers 
         Alignment       =   1  'Right Justify
         Caption         =   "Has Members?"
         Enabled         =   0   'False
         Height          =   192
         Left            =   1080
         TabIndex        =   4
         Top             =   1260
         Width           =   1452
      End
      Begin VB.TextBox txtTargetFrame 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   264
         Left            =   1320
         TabIndex        =   3
         Top             =   900
         Width           =   2652
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   372
         Left            =   1560
         TabIndex        =   5
         Top             =   2760
         Visible         =   0   'False
         Width           =   1152
      End
      Begin VB.PictureBox pboxInvalid 
         Height          =   432
         Left            =   780
         Picture         =   "kfcWebLinks.ctx":0312
         ScaleHeight     =   384
         ScaleWidth      =   384
         TabIndex        =   11
         Top             =   2760
         Visible         =   0   'False
         Width           =   432
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Default         =   -1  'True
         Height          =   372
         Left            =   2820
         TabIndex        =   8
         Top             =   2760
         Visible         =   0   'False
         Width           =   1152
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   372
         Left            =   2820
         TabIndex        =   6
         Top             =   2760
         Visible         =   0   'False
         Width           =   1152
      End
      Begin ComctlLib.ProgressBar prgLoad 
         Height          =   132
         Left            =   120
         TabIndex        =   12
         Top             =   2880
         Width           =   3852
         _ExtentX        =   6795
         _ExtentY        =   233
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label lblID 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ID"
         Height          =   192
         Left            =   3816
         TabIndex        =   21
         Top             =   1260
         Width           =   156
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Label:"
         Enabled         =   0   'False
         Height          =   192
         Left            =   600
         TabIndex        =   20
         Top             =   300
         Width           =   444
      End
      Begin VB.Label lblURL 
         AutoSize        =   -1  'True
         Caption         =   "URL:"
         Enabled         =   0   'False
         Height          =   192
         Left            =   660
         TabIndex        =   19
         Top             =   636
         Width           =   360
      End
      Begin VB.Label lblParentID 
         AutoSize        =   -1  'True
         Caption         =   "Parent ID:"
         Enabled         =   0   'False
         Height          =   192
         Left            =   324
         TabIndex        =   18
         Top             =   1896
         Visible         =   0   'False
         Width           =   696
      End
      Begin VB.Label lblButtonLabel 
         AutoSize        =   -1  'True
         Caption         =   "Button:"
         Enabled         =   0   'False
         Height          =   192
         Left            =   540
         TabIndex        =   17
         Top             =   1596
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblTargetFrame 
         AutoSize        =   -1  'True
         Caption         =   "Target Frame:"
         Enabled         =   0   'False
         Height          =   192
         Left            =   240
         TabIndex        =   16
         Top             =   936
         Width           =   1020
      End
      Begin ComctlLib.ImageList imlIcons 
         Left            =   300
         Top             =   2760
         _ExtentX        =   804
         _ExtentY        =   804
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   11
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "kfcWebLinks.ctx":061C
               Key             =   "Buttons32"
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "kfcWebLinks.ctx":0936
               Key             =   "Open"
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "kfcWebLinks.ctx":0C50
               Key             =   "Open32"
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "kfcWebLinks.ctx":0F6A
               Key             =   "Closed"
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "kfcWebLinks.ctx":1284
               Key             =   "EntireNet"
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "kfcWebLinks.ctx":159E
               Key             =   "Button32"
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "kfcWebLinks.ctx":18B8
               Key             =   "IE Document"
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "kfcWebLinks.ctx":1BD2
               Key             =   "Buttons"
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "kfcWebLinks.ctx":1EEC
               Key             =   "IE Shortcut"
            EndProperty
            BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "kfcWebLinks.ctx":2206
               Key             =   "Button"
            EndProperty
            BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "kfcWebLinks.ctx":2520
               Key             =   "Closed32"
            EndProperty
         EndProperty
      End
      Begin VB.Label lblLoad 
         AutoSize        =   -1  'True
         Caption         =   "lblLoad"
         Height          =   192
         Left            =   120
         TabIndex        =   15
         Top             =   2640
         Width           =   540
      End
   End
   Begin ComctlLib.TabStrip tsUpdate 
      Height          =   3792
      Left            =   60
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   60
      Width           =   7572
      _ExtentX        =   13356
      _ExtentY        =   6689
      MultiRow        =   -1  'True
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Web Links"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuContext 
      Caption         =   "Context"
      Visible         =   0   'False
      Begin VB.Menu mnuContextHyperlink 
         Caption         =   "&Hyperlink"
      End
      Begin VB.Menu mnuContextSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextNew 
         Caption         =   "&New..."
         Begin VB.Menu mnuContextNewGroup 
            Caption         =   "&Group"
         End
         Begin VB.Menu mnuContextNewLink 
            Caption         =   "&Link"
         End
      End
      Begin VB.Menu mnuContextRename 
         Caption         =   "&Rename"
      End
      Begin VB.Menu mnuContextUpdate 
         Caption         =   "&Update"
      End
      Begin VB.Menu mnuContextSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextDelete 
         Caption         =   "&Delete"
      End
   End
End
Attribute VB_Name = "kfcWebLinks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private cnKFC As New ADODB.Connection
Private cmdKFC As New ADODB.Command
Private RootIndex As Integer
Private gfDragMode As Boolean
Private DragNode As ComctlLib.Node
Private fUpdateInProgress As Boolean
Private fAdding As Boolean
Private LocalSite As String
Private WithEvents IE As SHDocVw.InternetExplorer
Attribute IE.VB_VarHelpID = -1
'Const gstrConnectionString = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=KFC.mdb;DefaultDir=E:\WebShare\wwwroot\Access;"
Const gstrProvider = "Microsoft.Jet.OLEDB.3.51"
Const gstrConnectionString = "E:\WebShare\wwwroot\Access\KFC.mdb"
Const gstrRunTimeUserName = "admin"
Const gstrRunTimePassword = ""
Event BeginEditMode()
Event EndEditMode()
Private Sub AddEntry(ByRef strID As String, strLabel As String, strParentID As String, strTargetFrame As String, strButton As String, strURL As String, fHasMembers As Boolean)
    Dim rsEntry As New ADODB.Recordset
    Dim intID As Integer
    
    Trace trcEnter, "AddEntry()"
    rsEntry.MaxRecords = 1
    Trace trcBody, "rsEntry.Open ""SELECT * from MenuEntries"", cnKFC, adOpenKeyset, adLockReadOnly"
    rsEntry.Open "SELECT * from MenuEntries", cnKFC, adOpenForwardOnly, adLockOptimistic
    Trace trcBody, "rsEntry.AddNew"
    rsEntry.AddNew
    strID = rsEntry("ID")
    If Len(VBencode(strLabel)) > rsEntry("Label").DefinedSize Then
        rsEntry("Label") = Mid(VBencode(strLabel), 1, rsEntry("Label").DefinedSize)
    Else
        rsEntry("Label") = VBencode(strLabel)
    End If
    rsEntry("ParentID") = VBencode(strParentID)
    rsEntry("TargetFrame") = VBencode(strTargetFrame)
    rsEntry("ButtonLabel") = VBencode(strButton)
    rsEntry("URL") = URLencode(strURL)
    rsEntry("HasMembers") = fHasMembers
    rsEntry.Update
    rsEntry.Close
    Set rsEntry = Nothing

    Trace trcExit, "AddEntry()"
End Sub
Private Function AddNode(ParentIndex As Integer, strID As String, strLabel As String, strButton As String, strParentID As String, fHasMembers As Boolean, fSelectNode As Boolean) As Integer
    Dim mNode As ComctlLib.Node
    Dim xNode As ComctlLib.Node
        
    Trace trcEnter, "AddNode()"
    If ParentIndex = 0 Then
        Set mNode = tvwDB.Nodes.Add()
    Else
        Set mNode = tvwDB.Nodes.Add(ParentIndex, tvwChild)
    End If
    mNode.Text = strLabel
    mNode.Sorted = True
    If ParentIndex = 0 Then 'It's the Root...
        mNode.Key = strLabel
        mNode.Tag = strLabel
        mNode.Image = "Buttons"
        ForceNodeSort mNode.Index
    Else
        If strID = "0" Then     'It's a button...
            mNode.Key = "Button: " & strButton
            mNode.Tag = "Button: " & strButton
            mNode.Image = "Button"
        Else
            mNode.Key = strButton & strParentID & strID
            If fHasMembers Then
                mNode.Tag = "Group: " & strID
                mNode.Image = "Closed"
            Else
                mNode.Tag = "Link: " & strID
                mNode.Image = "IE Shortcut"
            End If
        End If
        ForceNodeSort ParentIndex
    End If
    AddNode = mNode.Index
    
    If fSelectNode Then
        mNode.EnsureVisible
        mNode.Selected = True
        PopulateDetail mNode
    End If
    Trace trcExit, "AddNode()"
End Function
Private Sub ClearDetail()
    Dim Control As Control
   
    Trace trcEnter, "ClearDetail()"
    For Each Control In Controls
        If (TypeOf Control Is Label) Or (TypeOf Control Is Frame) Then Control.Caption = ""
        If (TypeOf Control Is TextBox) Then Control.Text = ""
        If (TypeOf Control Is CheckBox) Then Control.Value = 0
    Next
    
    frameDetail.Caption = "Detail"
    frameLayout.Caption = "Button Layout"
    
    lblLabel.Caption = "Label:"
    lblParentID.Caption = "Parent ID:"
    lblTargetFrame.Caption = "Target Frame:"
    lblButtonLabel.Caption = "Button Label:"
    lblURL.Caption = "URL:"
    
    lblLabel.Enabled = False
    lblParentID.Enabled = False
    lblTargetFrame.Enabled = False
    lblButtonLabel.Enabled = False
    lblURL.Enabled = False
    lblID.Enabled = False
    lblID.Visible = False
    
    txtLabel.Enabled = False
    txtParentID.Enabled = False
    txtTargetFrame.Enabled = False
    txtButtonLabel.Enabled = False
    txtURL.Enabled = False
    chkHasMembers.Enabled = False
    
    txtParentID.Visible = False
    txtButtonLabel.Visible = False
    
    lblLoad.Visible = False
    prgLoad.Visible = False
    
    cmdOK.Visible = False
    cmdCancel.Visible = False
    cmdUpdate.Visible = False
    cmdHyperlink.Visible = False
    Trace trcExit, "ClearDetail()"
End Sub
Private Sub cmdCancel_Click()
    DisableFields
    'tvwDB.Enabled = True
    fUpdateInProgress = False
End Sub
Private Sub cmdHyperlink_Click()
    Dim mNode As ComctlLib.Node
    
    Set mNode = tvwDB.SelectedItem
    If mNode Is tvwDB.Nodes(RootIndex) Then             'Root Level...
        IEhyperlink LocalSite, "_top"
    ElseIf IsLink(mNode) Then                           'Links...
        IEhyperlink txtURL.Text, txtTargetFrame.Text
    End If
End Sub
Private Sub cmdOK_Click()
    Dim mNode As ComctlLib.Node
    Dim rsEntry As New ADODB.Recordset
    Dim intID As Integer
    Dim fHasMembers As Boolean
    Dim strID As String
    
    Trace trcEnter, "cmdOK_Click()"
    
    If fAdding Then
        If chkHasMembers.Value = 1 Then
            fHasMembers = True
        Else
            fHasMembers = False
        End If
        AddEntry strID, txtLabel.Text, txtParentID.Text, txtTargetFrame.Text, txtButtonLabel.Text, txtURL.Text, fHasMembers
        AddNode tvwDB.SelectedItem.Index, strID, txtLabel.Text, txtButtonLabel.Text, txtParentID.Text, fHasMembers, True
        fAdding = False
    Else
        Set mNode = tvwDB.SelectedItem
        intID = Mid(lblID.Caption, 5)
        rsEntry.Open "SELECT * from MenuEntries where ID=" & intID, cnKFC, adOpenForwardOnly, adLockOptimistic
        rsEntry("Label") = VBencode(txtLabel.Text)
        rsEntry("ParentID") = VBencode(txtParentID.Text)      'Set behind the scenes...
        rsEntry("TargetFrame") = VBencode(txtTargetFrame.Text)
        rsEntry("ButtonLabel") = VBencode(txtButtonLabel.Text)
        rsEntry("URL") = URLencode(txtURL.Text)
        If chkHasMembers.Value = 1 Then
            rsEntry("HasMembers") = True
        Else
            rsEntry("HasMembers") = False
        End If
        rsEntry.Update
        rsEntry.Close
        mNode.Text = txtLabel.Text
        ForceNodeSort mNode.Index
    End If
    Set rsEntry = Nothing
    
    DisableFields
    tvwDB.SetFocus
    'tvwDB.Enabled = True
    fUpdateInProgress = False
    Trace trcExit, "cmdOK_Click()"
End Sub
Private Sub cmdUpdate_Click()
    EnableFields
    'tvwDB.Enabled = False
    fUpdateInProgress = True
End Sub
Private Sub CopyNode(TargetNode As ComctlLib.Node, SourceID As String)
    Dim mNode As ComctlLib.Node
    Dim rsSourceEntry As New ADODB.Recordset
    Dim rsTargetEntry As New ADODB.Recordset
    Dim TargetID As String
    Dim strButton As String
    Dim strID As String
    Dim SourceLabel As String
    Dim SourceURL As String
    Dim SourceTargetFrame As String
    Dim SourceHasMembers As Boolean
    
    If IsLink(TargetNode) Then
        TargetID = Trim(Mid(TargetNode.Tag, 6))
    ElseIf IsGroup(TargetNode) Then
        TargetID = Trim(Mid(TargetNode.Tag, 7))
    Else
        'Button Level...
        strButton = Trim(Mid(TargetNode.Tag, 9))
    End If
    
    If TargetID <> "" Then
        'First get the particulars of the new target Node...
        rsTargetEntry.Open "SELECT * from MenuEntries where ID=" & TargetID, cnKFC, adOpenForwardOnly, adLockReadOnly
        strButton = VBdecode(rsTargetEntry("ButtonLabel"))
        rsTargetEntry.Close
        Set rsTargetEntry = Nothing
    End If
    
    'Copy this guy to the new tree...
    rsSourceEntry.Open "SELECT * from MenuEntries where ID=" & SourceID, cnKFC, adOpenForwardOnly, adLockReadOnly
    SourceLabel = rsSourceEntry("Label")
    SourceTargetFrame = VBdecode(rsSourceEntry("TargetFrame"))
    SourceURL = VBdecode(rsSourceEntry("URL"))
    SourceHasMembers = rsSourceEntry("HasMembers")
    rsSourceEntry.Close
    AddEntry strID, SourceLabel, VBdecode(TargetID), SourceTargetFrame, strButton, SourceURL, SourceHasMembers
    Set mNode = tvwDB.Nodes(AddNode(TargetNode.Index, strID, SourceLabel, strButton, TargetID, SourceHasMembers, False))
    
    If SourceHasMembers Then
        'Now copy his children to the new tree...
        rsSourceEntry.Open "SELECT * from MenuEntries where ParentID=" & SourceID, cnKFC, adOpenForwardOnly, adLockOptimistic
        While Not rsSourceEntry.EOF
            CopyNode mNode, rsSourceEntry("ID")
            rsSourceEntry.MoveNext
        Wend
        rsSourceEntry.Close
    End If
    Set rsSourceEntry = Nothing
End Sub
Private Sub DeleteByParent(ParentID As Long)
    Dim rsEntry As New ADODB.Recordset
    
    'Note: Working off ParentCode (the way it's currently defined) is flawed...
    '      It should probably be changed to the ID field of the parent record
    '      to avoid ambiguity between different records with the same ParentCode...
    rsEntry.Open "SELECT * from MenuEntries where ParentID=" & ParentID, cnKFC, adOpenForwardOnly, adLockOptimistic
    While Not rsEntry.EOF
        DeleteByParent rsEntry("ID")
        rsEntry.Delete adAffectCurrent
        rsEntry.MoveNext
    Wend
    rsEntry.Close
    Set rsEntry = Nothing
End Sub
Public Sub DisableFields()
    Trace trcEnter, "DisableFields()"
    lblLabel.Enabled = False
    lblTargetFrame.Enabled = False
    lblURL.Enabled = False
    lblID.Enabled = False
    
    txtLabel.Enabled = False
    txtTargetFrame.Enabled = False
    txtURL.Enabled = False
    
    txtLabel.BackColor = vbButtonFace
    txtTargetFrame.BackColor = vbButtonFace
    txtURL.BackColor = vbButtonFace
    
    chkHasMembers.Enabled = False

    cmdOK.Visible = False
    cmdCancel.Visible = False
    cmdUpdate.Visible = True
    cmdHyperlink.Visible = True
    tsUpdate.TabStop = True
    tvwDB.TabStop = True
    RaiseEvent EndEditMode
    Trace trcExit, "DisableFields()"
End Sub
Public Sub EnableFields()
    Trace trcEnter, "EnableFields()"
    If frameDetail.Caption = "Link Detail" Then
        lblLabel.Enabled = True
        lblTargetFrame.Enabled = True
        lblURL.Enabled = True
        lblID.Enabled = True
        
        txtLabel.Enabled = True
        txtTargetFrame.Enabled = True
        txtURL.Enabled = True
        
        txtLabel.BackColor = vbWindowBackground
        txtTargetFrame.BackColor = vbWindowBackground
        txtURL.BackColor = vbWindowBackground
        
        chkHasMembers.Enabled = False
    Else
        lblLabel.Enabled = True
        lblID.Enabled = True
        
        txtLabel.Enabled = True
        
        txtLabel.BackColor = vbWindowBackground
    End If
    txtLabel.SetFocus

    cmdOK.Visible = True
    cmdOK.Default = True
    cmdCancel.Visible = True
    cmdCancel.Cancel = True
    cmdUpdate.Visible = False
    cmdHyperlink.Visible = False
    tsUpdate.TabStop = False
    tvwDB.TabStop = False
    RaiseEvent BeginEditMode
    Trace trcExit, "EnableFields()"
End Sub
Private Function FilterRecordset(rsTemp As ADODB.Recordset, strFilter As String) As ADODB.Recordset
    rsTemp.Filter = strFilter
    Set FilterRecordset = rsTemp
End Function
Private Function FindEntry(Key As String) As ComctlLib.Node
    Trace trcEnter, "FindEntry()"
    
    On Error Resume Next
    Set FindEntry = tvwDB.Nodes(Key)
    Select Case Err.Number
        Case 0
            On Error GoTo 0
        Case 35601
            Set FindEntry = Nothing
            Err.Clear
        Case Else
            MsgBox Err.Description & vbCr & "( Error #" & Err.Number & ")", vbCritical + vbOKOnly
            Stop
    End Select
    On Error GoTo 0
    Trace trcExit, "FindEntry()"
End Function
Private Sub ForceNodeSort(ParentIndex As Integer)
    Dim xNode As ComctlLib.Node
    
    'Shouldn't need to do this, but sorting doesn't work unless we do...
    Set xNode = tvwDB.Nodes.Add(ParentIndex, tvwChild)
    tvwDB.Nodes.Remove xNode.Index
End Sub
Private Sub IE_OnQuit()
    Set IE = Nothing
End Sub
Private Sub IE_OnVisible(ByVal Visible As Boolean)
    'Debug.Print "IE_OnVisible()"
End Sub
Private Sub mnuContextDelete_Click()
    Dim intID As Integer
    Dim rsEntry As New ADODB.Recordset
    Dim mNode As ComctlLib.Node
    
    If MsgBox("Are you sure you want to delete " & tvwDB.SelectedItem.Text & "?", vbYesNo) = vbNo Then Exit Sub
    
    Set mNode = tvwDB.SelectedItem
    If IsLink(mNode) Then
        intID = Trim(Mid(mNode.Tag, 6))
    ElseIf IsGroup(mNode) Then
        intID = Trim(Mid(mNode.Tag, 7))
    End If
    
    rsEntry.Open "SELECT * from MenuEntries where ID=" & intID, cnKFC, adOpenForwardOnly, adLockOptimistic
    DeleteByParent rsEntry("ID")
    rsEntry.Delete adAffectCurrent
    rsEntry.Close
    Set rsEntry = Nothing
    
    tvwDB.Nodes.Remove mNode.Index
    tvwDB.SetFocus
    tvwDB_NodeClick tvwDB.SelectedItem
End Sub
Private Sub mnuContextHyperlink_Click()
    If cmdHyperlink.Enabled Then cmdHyperlink_Click
End Sub
Private Sub mnuContextNewGroup_Click()
    Dim intID As Integer
    Dim rsEntry As New ADODB.Recordset
    Dim mNode As ComctlLib.Node
    
    ClearDetail
    Set mNode = tvwDB.SelectedItem
    If IsLink(mNode) Then
        intID = Trim(Mid(mNode.Tag, 6))
    ElseIf IsGroup(mNode) Then
        intID = Trim(Mid(mNode.Tag, 7))
    End If
    
    If intID = 0 Then
        txtParentID.Text = 0
        txtButtonLabel.Text = Trim(Mid(mNode.Tag, 9))
    Else
        rsEntry.Open "SELECT * from MenuEntries where ID=" & intID, cnKFC, adOpenForwardOnly, adLockReadOnly
        txtParentID.Text = rsEntry("ID")
        txtButtonLabel.Text = VBdecode(rsEntry("ButtonLabel"))
        rsEntry.Close
    End If
    Set rsEntry = Nothing
    
    frameDetail.Caption = "Group Detail"
    chkHasMembers.Value = vbChecked
    
    fAdding = True
    cmdUpdate_Click
End Sub
Private Sub mnuContextNewLink_Click()
    Dim intID As Integer
    Dim rsEntry As New ADODB.Recordset
    Dim mNode As ComctlLib.Node
    
    ClearDetail
    Set mNode = tvwDB.SelectedItem
    If IsLink(mNode) Then
        intID = Trim(Mid(mNode.Tag, 6))
    ElseIf IsGroup(mNode) Then
        intID = Trim(Mid(mNode.Tag, 7))
    End If
    
    If intID = 0 Then
        txtParentID.Text = 0
        txtButtonLabel.Text = Trim(Mid(mNode.Tag, 9))
    Else
        rsEntry.Open "SELECT * from MenuEntries where ID=" & intID, cnKFC, adOpenForwardOnly, adLockReadOnly
        txtParentID.Text = rsEntry("ID")
        txtButtonLabel.Text = VBdecode(rsEntry("ButtonLabel"))
        rsEntry.Close
    End If
    Set rsEntry = Nothing
    
    frameDetail.Caption = "Link Detail"
    txtTargetFrame.Text = "_top"
    
    fAdding = True
    cmdUpdate_Click
End Sub
Private Sub mnuContextRename_Click()
    tvwDB.StartLabelEdit
End Sub
Private Sub mnuContextUpdate_Click()
    If cmdUpdate.Enabled Then cmdUpdate_Click
End Sub
Private Sub PopulateButton(ButtonLabel As String, ParentID As String, intTreeViewIndex As Integer)
    Dim rsMenuEntries As New ADODB.Recordset
    Dim pNode As ComctlLib.Node
    Dim mNode As ComctlLib.Node
    Dim SQLstatement As String
    Dim NodeIndex As Integer
    
    Trace trcEnter, "PopulateButton()"
    
    SQLstatement = "SELECT * FROM MenuEntries where ButtonLabel='" & ButtonLabel & "' and ParentID=" & ParentID & " order by ButtonLabel,ParentID,Label"
    Trace trcBody, "rsMenuEntries.Open """ & SQLstatement & """, cnKFC, adOpenKeyset, adLockReadOnly"
    rsMenuEntries.Open SQLstatement, cnKFC, adOpenForwardOnly, adLockReadOnly
    Do Until rsMenuEntries.EOF
        prgLoad.Value = prgLoad.Value + 1
        
        Trace trcBody, "Processing Entry: ButtonLabel: " & rsMenuEntries("ButtonLabel") & "; ID: " & rsMenuEntries("ID") & "; Label: " & rsMenuEntries("Label")
        NodeIndex = intTreeViewIndex
        NodeIndex = AddNode(NodeIndex, rsMenuEntries("ID"), VBdecode(rsMenuEntries("Label")), ButtonLabel, VBdecode(rsMenuEntries("ParentID")), rsMenuEntries("HasMembers"), False)
        If rsMenuEntries("HasMembers") Then
            PopulateButton ButtonLabel, rsMenuEntries("ID"), NodeIndex
        End If
        
        Trace trcBody, "rsMenuEntries.MoveNext"
        rsMenuEntries.MoveNext
        DoEvents
    Loop
    
    ForceNodeSort intTreeViewIndex
    
    Trace trcBody, "rsMenuEntries.Close"
    rsMenuEntries.Close
    Set rsMenuEntries = Nothing
    Trace trcExit, "PopulateButton()"
End Sub
Private Sub PopulateDetail(ByVal Node As ComctlLib.Node)
    Dim intID As Integer
    Dim rsEntry As New ADODB.Recordset
    
    Trace trcEnter, "PopulateDetail()"
    If IsLink(Node) Then
        intID = Trim(Mid(Node.Tag, 6))
        frameDetail.Caption = "Link Detail"
        cmdUpdate.Visible = True
        cmdHyperlink.Visible = True
        cmdHyperlink.Default = True
    ElseIf IsGroup(Node) Then
        intID = Trim(Mid(Node.Tag, 7))
        frameDetail.Caption = "Group Detail"
        cmdUpdate.Visible = True
    End If
    
    Trace trcBody, "rsEntry.Open ""SELECT * from MenuEntries where ID=" & intID & """, cnKFC, adOpenForwardOnly, adLockReadOnly"
    rsEntry.Open "SELECT * from MenuEntries where ID=" & intID, cnKFC, adOpenForwardOnly, adLockReadOnly
    lblID.Caption = "ID: " & intID
    lblID.Visible = True
    txtLabel.Text = VBdecode(rsEntry("Label"))
    txtTargetFrame.Text = VBdecode(rsEntry("TargetFrame"))
    txtParentID.Text = VBdecode(rsEntry("ParentID"))
    txtButtonLabel.Text = VBdecode(rsEntry("ButtonLabel"))
    txtURL.Text = URLdecode(rsEntry("URL"))
    If rsEntry("HasMembers") Then
        chkHasMembers.Value = 1
    Else
        chkHasMembers.Value = 0
    End If
    cmdUpdate.Visible = True
    
    rsEntry.Close
    Set rsEntry = Nothing
    Trace trcExit, "PopulateDetail()"
End Sub
Public Sub PopulateMenu()
    Dim intIndex As Integer
    Dim rsTable As New ADODB.Recordset
    Dim rsButtons As New ADODB.Recordset
    Dim SQLstatement As String
    Dim Count As Long
    
    'gfTraceMode = True
    
    Trace trcBody, String(132, "=")
    Trace trcEnter, "PopulateMenu()"
    cnKFC.ConnectionTimeout = 60
    cnKFC.CommandTimeout = 60
    cnKFC.Provider = "Microsoft.Jet.OLEDB.3.51"
    cnKFC.Open gstrConnectionString, gstrRunTimeUserName, gstrRunTimePassword
    
    SQLstatement = "SELECT * FROM MenuEntries"
    rsTable.CacheSize = 2000
    rsTable.Open SQLstatement, cnKFC, adOpenKeyset, adLockReadOnly
    prgLoad.Max = rsTable.RecordCount
    rsTable.Close
    Set rsTable = Nothing
    
    tvwDB.Sorted = True
    RootIndex = AddNode(0, "0", "Web Menu Buttons", "", "0", True, False)
    
    'rsButtons.CacheSize = 100
    lblLoad.Caption = "Loading Links DataBase..."
    
    'SQLstatement = "SELECT Count(*) as Count FROM MenuEntries"
    'Trace trcBody, "rsButtons.Open """ & SQLstatement & """, cnKFC, adOpenKeyset, adLockReadOnly"
    'rsButtons.Open SQLstatement, cnKFC, adOpenKeyset, adLockReadOnly
    'prgLoad.Max = rsButtons!Count
    'rsButtons.Close
    
    prgLoad.Visible = True
    prgLoad.Value = 0
    Count = 0
    lblID.Visible = False
    lblLoad.Visible = True
    
    SQLstatement = "SELECT Distinct ButtonLabel FROM MenuEntries order by ButtonLabel"
'SQLstatement = "SELECT Distinct ButtonLabel FROM MenuEntries where buttonlabel = 'Books' order by ButtonLabel"
    Trace trcBody, "rsButtons.Open """ & SQLstatement & """, cnKFC, adOpenKeyset, adLockReadOnly"
    rsButtons.Open SQLstatement, cnKFC, adOpenKeyset, adLockReadOnly
    
    While Not rsButtons.EOF
        lblLoad.Caption = "Loading " & rsButtons!ButtonLabel & " Entries..."
        intIndex = AddNode(RootIndex, "0", VBdecode(rsButtons("ButtonLabel")), VBdecode(rsButtons("ButtonLabel")), "0", True, False)
        PopulateButton rsButtons("ButtonLabel"), 0, intIndex
        'lblLoad.Caption = rsButtons!ButtonLabel & " Load Complete."
        tvwDB.Refresh
        DoEvents
        rsButtons.MoveNext
    Wend
    Trace trcBody, "rsButtons.Close"
    rsButtons.Close
    Set rsButtons = Nothing
   
    ForceNodeSort RootIndex
    tvwDB.Nodes(RootIndex).Selected = True
    tvwDB.Nodes(RootIndex).Expanded = True
    tvwDB.SetFocus
    lblLoad.Visible = False
    prgLoad.Visible = False
    Trace trcExit, "PopulateMenu()"
End Sub
Private Sub IEhyperlink(strURL, strFrame)
    Dim TwipWidth As Integer
    Dim TwipHeight As Integer
    Dim TwipLeft As Integer
    Dim TwipTop As Integer
    Dim TargetURL As String
    Dim TargetFrame As String
    
    On Error Resume Next
    
    If IE Is Nothing Then
        Set IE = New InternetExplorer
        IE.AddressBar = True
        IE.FullScreen = False
        IE.MenuBar = True
        IE.RegisterAsBrowser = True
        IE.Resizable = True
        IE.StatusBar = True
        IE.TheaterMode = False  'although very cool...
        IE.Visible = True
        
        IE.Width = 875
        IE.Height = 711
        TwipWidth = Screen.TwipsPerPixelX * IE.Width
        TwipHeight = Screen.TwipsPerPixelY * IE.Height
        TwipTop = Screen.TwipsPerPixelX * IE.Top
        TwipLeft = Screen.TwipsPerPixelY * IE.Left
        If TwipWidth > Screen.Width Then IE.Width = Screen.Width / Screen.TwipsPerPixelX
        If TwipHeight > Screen.Height Then IE.Height = Screen.Height / Screen.TwipsPerPixelY
        IE.Top = 0
        IE.Left = (Screen.Width / Screen.TwipsPerPixelX) - IE.Width
    End If
    
    TargetURL = strURL
    TargetFrame = strFrame
    If Left(TargetURL, 1) = "/" Then      'Better be local site...
        TargetURL = Left(LocalSite, Len(LocalSite) - 1) & TargetURL
        'Check to see if we're on the local site's frame page before allowing local frame names...
        If LCase(TargetFrame) = "index" Or LCase(TargetFrame) = "body" Then    '...which are the only ones allowed...
            'If we're not already on the local site, then use "_top"...
            If IE.LocationURL <> LCase(LocalSite) And _
                IE.LocationURL <> LCase(LocalSite & "default.htm") And _
                IE.LocationURL <> LCase(LocalSite & "default.asp") Then
                TargetFrame = "_top"
            End If
        End If
    ElseIf IE.LocationURL = vbNullString And TargetFrame <> "_top" Then
        TargetFrame = "_top"
    End If
    IE.Navigate TargetURL, , TargetFrame
End Sub
Private Sub tvwDB_AfterLabelEdit(Cancel As Integer, NewString As String)
    Dim rsEntry As New ADODB.Recordset
    Dim intID As Integer
    Dim mNode As ComctlLib.Node
    
    Trace trcEnter, "tvwDB_AfterLabelEdit()"
    If Len(NewString) = 0 Then
        Cancel = True
    Else
        Set mNode = tvwDB.SelectedItem
        intID = Mid(mNode.Tag, InStr(mNode.Tag, ": ") + 2)
        
        rsEntry.Open "SELECT Label from MenuEntries where ID=" & intID, cnKFC, adOpenKeyset, adLockOptimistic
        rsEntry("Label") = VBencode(NewString)
        rsEntry.Update
        txtLabel.Text = NewString
        
        rsEntry.Close
        Set rsEntry = Nothing
    
        DisableFields
    End If
    Trace trcExit, "tvwDB_AfterLabelEdit()"
End Sub
Private Sub tvwDB_BeforeLabelEdit(Cancel As Integer)
    Dim mNode As ComctlLib.Node
    
    Trace trcEnter, "tvwDB_BeforeLabelEdit()"
    Set mNode = tvwDB.SelectedItem
    If mNode.Index = RootIndex Then
        Cancel = 1
        Exit Sub
    ElseIf mNode.Parent.Index = RootIndex Then
        Cancel = 1
        Exit Sub
    End If
    Trace trcExit, "tvwDB_BeforeLabelEdit()"
End Sub
Private Sub tvwDB_Collapse(ByVal Node As ComctlLib.Node)
    Trace trcEnter, "tvwDB_Collapse()"
    If IsGroup(Node) Then
        Node.Image = "Closed"
    End If
    Trace trcExit, "tvwDB_Collapse()"
End Sub
Private Sub tvwDB_Expand(ByVal Node As ComctlLib.Node)
    Trace trcEnter, "tvwDB_Expand()"
    If IsGroup(Node) Then
        Node.Image = "Open"
    End If
    Trace trcExit, "tvwDB_Expand()"
End Sub
Private Sub tvwDB_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Trace trcEnter, "tvwDB_MouseDown()"
    If fUpdateInProgress Then cmdCancel_Click
    Set tvwDB.SelectedItem = tvwDB.HitTest(x, y)
    Trace trcExit, "tvwDB_MouseDown()"
End Sub
Private Sub tvwDB_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Trace trcEnter, "tvwDB_MouseMove()"
    If Button = vbLeftButton Then ' Signal a Drag operation.
        If IsGroup(tvwDB.SelectedItem) Or IsLink(tvwDB.SelectedItem) Then
            Set DragNode = tvwDB.SelectedItem
            gfDragMode = True
            
            ' Set the drag icon with the CreateDragImage method.
            tvwDB.DragIcon = tvwDB.SelectedItem.CreateDragImage
            tvwDB.Drag vbBeginDrag
        Else
            'MsgBox "Sorry, but the list of buttons is fixed," & vbcr & _
            '       "buttons cannot be moved.", _
            '        vbExclamation + vbOKOnly
            tvwDB.Drag vbCancel
        End If
    End If
    Trace trcExit, "tvwDB_MouseMove()"
End Sub
Private Sub tvwDB_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim mNode As ComctlLib.Node
    Dim ctl As Control
    
    Trace trcEnter, "tvwDB_MouseUp()"
    If Button = vbKeyRButton Then
        Set mNode = tvwDB.SelectedItem
        'Make everything visible & enabled to start...
        For Each ctl In UserControl.Controls
            If TypeName(ctl) = "Menu" Then
                ctl.Enabled = True
                ctl.Visible = True
            End If
        Next
        
        If mNode Is tvwDB.Nodes(RootIndex) Then             'Root Level...
            mnuContextSep1.Visible = False
            mnuContextNew.Visible = False
            mnuContextRename.Visible = False
            mnuContextUpdate.Visible = False
            mnuContextSep2.Visible = False
            mnuContextDelete.Visible = False
        ElseIf mNode.Parent Is tvwDB.Nodes(RootIndex) Then  'Button level...
            mnuContextHyperlink.Enabled = False
            mnuContextHyperlink.Visible = False
            mnuContextSep1.Visible = False
            mnuContextRename.Visible = False
            mnuContextUpdate.Visible = False
            mnuContextSep2.Visible = False
            mnuContextDelete.Visible = False
        ElseIf IsGroup(mNode) Then                          'Groups...
            mnuContextHyperlink.Visible = False
            mnuContextSep1.Visible = False
        ElseIf IsLink(mNode) Then                           'Links...
            mnuContextNew.Enabled = False
            mnuContextNew.Visible = False
        End If
        PopupMenu mnuContext
    End If
    Trace trcExit, "tvwDB_MouseUp()"
End Sub
Private Sub tvwDB_NodeClick(ByVal Node As ComctlLib.Node)
    Dim intID As Integer
    
    Trace trcEnter, "tvwDB_NodeClick()"
    ClearDetail
    If IsLink(Node) Or IsGroup(Node) Then PopulateDetail Node
    Trace trcExit, "tvwDB_NodeClick()"
End Sub
Private Sub tvwDB_OLECompleteDrag(Effect As Long)
    Dim strID As String
    Dim rsEntry As ADODB.Recordset
    
    If Effect And vbDropEffectMove Then
        If DragNode Is Nothing Then
        Else
            'Remove the node from the TreeView...
            If IsLink(DragNode) Then
                strID = Trim(Mid(DragNode.Tag, 6))
            ElseIf IsGroup(DragNode) Then
                strID = Trim(Mid(DragNode.Tag, 7))
            End If
    
            Set rsEntry = New ADODB.Recordset
            rsEntry.Open "SELECT * from MenuEntries where ID=" & strID, cnKFC, adOpenForwardOnly, adLockOptimistic
            DeleteByParent rsEntry("ID")
            rsEntry.Delete adAffectCurrent
            rsEntry.Close
            Set rsEntry = Nothing
        End If
    End If
    Set DragNode = Nothing
    Screen.MousePointer = vbDefault
End Sub
Private Sub tvwDB_OLEDragDrop(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    Dim rsEntry As ADODB.Recordset
    Dim mNode As ComctlLib.Node
    Dim tNode As ComctlLib.Node
    Dim vFN As String
    Dim strID As String
    Dim strParentID As String
    Dim strLabel As String
    Dim strURL As String
    Dim strButton As String
    
    'By default, dropped items are Moved, Ctrl-Drop Copies...
    If Shift And vbCtrlMask Then
        Effect = Effect And vbDropEffectCopy
    Else
        Effect = Effect And vbDropEffectMove
    End If
    
    'Debug.Print "OLEDragDrop(Data, " & ", " & Effect & ", " & Button & ", " & Shift & ", " & x & ", " & y & ")"
    Set tNode = tvwDB.HitTest(x, y)
    If IsLink(tNode) Then
        If IsGroup(tvwDB.SelectedItem) Then
            MsgBox "Cannot drag folders on top of links." & vbCr & _
                   "Duh, what were you thinking...?.", _
                   vbExclamation + vbOKOnly
        Else
            MsgBox "Can only drag links/shortcuts to folders or buttons." & vbCr & _
                   "If you want to combine multiple links in a " & vbCr & _
                   "new folder, you'll have to create the new " & vbCr & _
                   "folder first (right mouse click).", _
                   vbExclamation + vbOKOnly
        End If
        Effect = Effect And vbDropEffectNone
        GoTo ExitSub
    ElseIf tNode Is tvwDB.Nodes(RootIndex) Then
        Effect = Effect And vbDropEffectNone
        GoTo ExitSub
    End If
        
    If DragNode Is Nothing Then
        Set tvwDB.DropHighlight = tvwDB.HitTest(x, y)
        If Data.GetFormat(vbCFFiles) Then
            For i = 1 To Data.Files.Count
                vFN = Data.Files.Item(i)
                
                'Get Parent information...
                Set mNode = tvwDB.DropHighlight
                If IsLink(mNode) Then
                    strParentID = Trim(Mid(mNode.Tag, 6))
                ElseIf IsGroup(mNode) Then
                    strParentID = Trim(Mid(mNode.Tag, 7))
                Else
                    strParentID = "0"
                End If
                
                If strParentID = "0" Then
                    strButton = Trim(Mid(mNode.Tag, 9))
                Else
                    Set rsEntry = New ADODB.Recordset
                    rsEntry.Open "SELECT * from MenuEntries where ID=" & strParentID, cnKFC, adOpenForwardOnly, adLockReadOnly
                    strParentID = rsEntry("ID")
                    strButton = VBdecode(rsEntry("ButtonLabel"))
                    rsEntry.Close
                    Set rsEntry = Nothing
                End If
                
                strLabel = ParsePath(vFN, FileNameBase)
                strURL = GetINIKey(vFN, "InternetShortcut", "URL", "")
                If LCase(ParsePath(vFN, FileNameExt)) <> ".url" Or strURL = "" Then
                    MsgBox ParsePath(vFN, FileNameBaseExt) & " is not a valid Internet Shortcut.", vbExclamation, "WebLinks Error"
                Else
                    AddEntry strID, strLabel, strParentID, "_top", strButton, strURL, False
                    AddNode tvwDB.DropHighlight.Index, strID, strLabel, strButton, strParentID, False, True
                    If Not SendToRecycleBin(vFN, True, False) Then MsgBox "Cannot delete " & vFN, vbExclamation
                End If
            Next
        Else
            'We'll deal with this if/when we need to...
            Effect = Effect And vbDropEffectNone
            GoTo ExitSub
        End If
    Else
        If DragNode Is tvwDB.DropHighlight Then GoTo ExitSub
        If DragNode.Parent Is tvwDB.DropHighlight Then GoTo ExitSub
        If IsLink(DragNode) Then
            strID = Trim(Mid(DragNode.Tag, 6))
        ElseIf IsGroup(DragNode) Then
            strID = Trim(Mid(DragNode.Tag, 7))
        End If
        CopyNode tvwDB.DropHighlight, strID
    End If
    
ExitSub:
    Set tvwDB.DropHighlight = Nothing
    If Not DragNode Is Nothing Then
        'This is done in the DragComplete Event...
        'Set DragNode = Nothing
        gfDragMode = False
    End If
End Sub
Private Sub tvwDB_OLEDragOver(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    Dim tNode As ComctlLib.Node
    Dim i As Integer
    Dim vFN As String
    
    'Debug.Print "OLEDragOver(Data, " & Effect & ", " & Button & ", " & Shift & ", " & x & ", " & y & ", " & State & ")"
    If Shift And vbCtrlMask Then
        Effect = Effect And vbDropEffectCopy
    Else
        Effect = Effect And vbDropEffectMove
    End If
    
    Set tNode = tvwDB.HitTest(x, y)
    If tNode Is Nothing Then
        Effect = Effect And vbDropEffectNone
    ElseIf IsLink(tNode) Then
        Effect = Effect And vbDropEffectNone
    ElseIf DragNode Is Nothing Then
        If tNode.Index = RootIndex Then
            Effect = Effect And vbDropEffectNone
        Else
            If Data.GetFormat(vbCFFiles) Then
                For i = 1 To Data.Files.Count
                    vFN = Data.Files.Item(i)
                    If LCase(ParsePath(vFN, FileNameExt)) <> ".url" Then
                        Effect = Effect And vbDropEffectNone
                    End If
                Next
            Else
                If Data.GetFormat(vbCFText) Then
                    'If we're dragging from IE... the Data represents the URL, but
                    'we have no way of getting to the page's title...
                End If
                Effect = Effect And vbDropEffectNone
            End If
        End If
    ElseIf DragNode Is tNode Then
        Effect = Effect And vbDropEffectNone
    ElseIf DragNode.Parent Is tNode Then
        Effect = Effect And vbDropEffectNone
    ElseIf tNode.Index = RootIndex Then
        Effect = Effect And vbDropEffectNone
    End If
    If Not (Effect And vbDropEffectNone) Then Set tvwDB.DropHighlight = tNode
End Sub
Private Sub tvwDB_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    'Debug.Print "OLEGiveFeedback(" & Effect & ", " & DefaultCursors & ")"
End Sub
Private Sub tvwDB_OLESetData(Data As ComctlLib.DataObject, DataFormat As Integer)
    'Debug.Print "OLESetData(Data, " & DataFormat & ")"
    If DataFormat = vbCFText Then
        If gfDragMode Then Data.SetData tvwDB.SelectedItem.Text, vbCFText
   End If
End Sub
Private Sub tvwDB_OLEStartDrag(Data As ComctlLib.DataObject, AllowedEffects As Long)
    'Debug.Print "OLEStartDrag(Data, " & AllowedEffects & ")"
    AllowedEffects = vbDropEffectNone
    If IsGroup(tvwDB.SelectedItem) Or IsLink(tvwDB.SelectedItem) Then
        Set DragNode = tvwDB.SelectedItem
        gfDragMode = True
        AllowedEffects = vbDropEffectMove Or vbDropEffectCopy
    End If
End Sub
Private Sub txtLabel_GotFocus()
    txtLabel.SelStart = 0
    txtLabel.SelLength = Len(txtLabel.Text)
End Sub
Private Sub txtParentID_GotFocus()
    txtParentID.SelStart = 0
    txtParentID.SelLength = Len(txtParentID.Text)
End Sub
Private Sub txtTargetFrame_GotFocus()
    txtTargetFrame.SelStart = 0
    txtTargetFrame.SelLength = Len(txtTargetFrame.Text)
End Sub
Private Sub txtTargetFrame_Validate(Cancel As Boolean)
    If txtTargetFrame.Text = "" Then txtTargetFrame.Text = "_top"
    Select Case LCase(txtTargetFrame.Text)
        Case "_top", "_blank"
            'OK...
        Case Else
            If Left(txtURL.Text, 1) = "/" Then
                'Only allow acceptable frame names on local site...
                Select Case LCase(txtTargetFrame.Text)
                    Case "index", "body"
                        'OK...
                    Case Else
                        MsgBox "Invalid frame for local site." & vbCr & "Supported frames: ""Index"" or ""Body""", vbExclamation
                        Cancel = True
                End Select
            Else
                MsgBox "Invalid frame specified." & vbCr & "Supported frames: ""_top"" or ""_blank""", vbExclamation
                Cancel = True
            End If
    End Select
End Sub
Private Sub txtURL_GotFocus()
    txtURL.SelStart = 0
    txtURL.SelLength = Len(txtURL.Text)
End Sub
Private Sub UserControl_Resize()
    tsUpdate.Width = UserControl.ScaleWidth - (2 * 60)
    tsUpdate.Height = UserControl.ScaleHeight - (2 * 60)
    frameDetail.Left = tsUpdate.Width - frameDetail.Width - 60
    frameDetail.Top = 420
    frameDetail.Height = tsUpdate.Height - frameDetail.Top - 60
    frameLayout.Left = 180
    frameLayout.Top = 420
    frameLayout.Width = frameDetail.Left - (3 * 60)
    frameLayout.Height = frameDetail.Height
    
    tvwDB.Left = 60
    tvwDB.Top = 240
    tvwDB.Width = frameLayout.Width - (2 * 60)
    tvwDB.Height = frameLayout.Height - tvwDB.Top - 60
End Sub
Private Sub UserControl_Show()
    LocalSite = "http://" & LCase(XGetComputerName) & "/"
End Sub
Private Sub UserControl_Terminate()
    Trace trcEnter, "UserControl_Terminate()"
    ' Close everything...
    On Error Resume Next
    Set cmdKFC = Nothing
    cnKFC.Close
    Set cnKFC = Nothing
    Trace trcExit, "UserControl_Terminate()"
End Sub


