VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmImages 
   Caption         =   "Image Display"
   ClientHeight    =   4536
   ClientLeft      =   132
   ClientTop       =   360
   ClientWidth     =   8088
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4536
   ScaleWidth      =   8088
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraImages 
      Height          =   2712
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   660
      Width           =   7692
      Begin VB.CheckBox chkThumbnail 
         Alignment       =   1  'Right Justify
         Caption         =   "Is this Image considered a Thumbnail (as opposed to a full image)?"
         DataSource      =   "dtaData"
         Height          =   252
         Left            =   300
         TabIndex        =   18
         Top             =   960
         Width           =   5172
      End
      Begin VB.Frame fraRelated 
         Caption         =   "Related Information"
         Height          =   1332
         Left            =   1068
         TabIndex        =   13
         Top             =   1320
         Width           =   5952
         Begin MSDataListLib.DataCombo dbcTable 
            Height          =   288
            Left            =   960
            TabIndex        =   15
            Top             =   240
            Width           =   4272
            _ExtentX        =   7535
            _ExtentY        =   508
            _Version        =   393216
            Style           =   2
            Text            =   "Table"
         End
         Begin MSDataListLib.DataCombo dbcRecord 
            Height          =   288
            Left            =   948
            TabIndex        =   17
            Top             =   600
            Width           =   4272
            _ExtentX        =   7535
            _ExtentY        =   508
            _Version        =   393216
            Style           =   2
            Text            =   "Record"
         End
         Begin MSDataListLib.DataCombo dbcThumbnail 
            Height          =   288
            Left            =   1476
            TabIndex        =   20
            Top             =   960
            Width           =   4272
            _ExtentX        =   7535
            _ExtentY        =   508
            _Version        =   393216
            Style           =   2
            Text            =   "Thumbnail"
         End
         Begin VB.Label lblThumbnail 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Thumbnail Image:"
            Height          =   192
            Left            =   72
            TabIndex        =   19
            Top             =   1008
            Width           =   1284
         End
         Begin VB.Label lblRecord 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Record:"
            Height          =   192
            Left            =   252
            TabIndex        =   16
            Top             =   648
            Width           =   576
         End
         Begin VB.Label lblTable 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Table:"
            Height          =   192
            Left            =   372
            TabIndex        =   14
            Top             =   288
            Width           =   468
         End
      End
      Begin VB.TextBox txtURL 
         DataSource      =   "dtaData"
         Height          =   288
         Left            =   720
         TabIndex        =   11
         Text            =   "URL"
         Top             =   600
         Width           =   6552
      End
      Begin VB.TextBox txtName 
         DataSource      =   "dtaData"
         Height          =   288
         Left            =   720
         TabIndex        =   1
         Text            =   "Name"
         Top             =   240
         Width           =   6552
      End
      Begin VB.Label lblURL 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "URL:"
         Height          =   192
         Left            =   300
         TabIndex        =   12
         Top             =   648
         Width           =   360
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   192
         Left            =   180
         TabIndex        =   10
         Top             =   288
         Width           =   480
      End
   End
   Begin MSComDlg.CommonDialog dlgImages 
      Left            =   2040
      Top             =   3840
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.Frame fraImages 
      Height          =   2712
      Index           =   2
      Left            =   180
      TabIndex        =   21
      Top             =   660
      Width           =   7692
      Begin VB.CommandButton cmdLoad 
         Caption         =   "&Load Image"
         Height          =   252
         Left            =   4800
         TabIndex        =   24
         Top             =   2400
         Width           =   1392
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "&View Full Picture"
         Height          =   252
         Left            =   6240
         TabIndex        =   23
         Top             =   2400
         Width           =   1392
      End
      Begin VB.PictureBox picImage 
         DataField       =   "Image"
         DataSource      =   "dtaData"
         Height          =   2232
         Left            =   60
         ScaleHeight     =   2184
         ScaleWidth      =   7524
         TabIndex        =   22
         Top             =   180
         Width           =   7572
      End
   End
   Begin VB.Frame fraImages 
      Height          =   2712
      Index           =   1
      Left            =   180
      TabIndex        =   2
      Top             =   660
      Width           =   7692
      Begin RichTextLib.RichTextBox rtxtCaption 
         Height          =   2472
         Left            =   60
         TabIndex        =   3
         Top             =   180
         Width           =   7572
         _ExtentX        =   13356
         _ExtentY        =   4360
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmImages.frx":0000
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   5940
      TabIndex        =   6
      Top             =   3900
      Width           =   972
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   372
      Left            =   6960
      TabIndex        =   5
      Top             =   3900
      Width           =   972
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   4
      Top             =   4284
      Width           =   8088
      _ExtentX        =   14266
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
            Object.Width           =   8784
            Key             =   "Message"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   2
            Object.Width           =   1270
            MinWidth        =   1270
            TextSave        =   "12:09 AM"
            Key             =   "Time"
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adodcMain 
      Height          =   312
      Left            =   468
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
      Caption         =   ""
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
   Begin MSComctlLib.TabStrip tsImages 
      Height          =   3072
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   7812
      _ExtentX        =   13780
      _ExtentY        =   5419
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            Key             =   "General"
            Object.Tag             =   "General"
            Object.ToolTipText     =   "General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Caption"
            Key             =   "Caption"
            Object.Tag             =   "Caption"
            Object.ToolTipText     =   "Caption"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Image"
            Key             =   "Image"
            Object.Tag             =   "Image"
            Object.ToolTipText     =   "Image"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbMain 
      Align           =   1  'Align Top
      Height          =   288
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   8088
      _ExtentX        =   14266
      _ExtentY        =   508
      ButtonWidth     =   487
      ButtonHeight    =   466
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlSmall"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Report"
            Object.ToolTipText     =   "Report"
            ImageKey        =   "Report"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SQL"
            Object.ToolTipText     =   "SQL Window"
            ImageKey        =   "SQL"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New record"
            ImageKey        =   "NewRecord"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Modify"
            Object.ToolTipText     =   "Modify record"
            ImageKey        =   "Modify"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete record"
            ImageKey        =   "Delete"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh data"
            ImageKey        =   "Refresh"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Filter"
            Object.ToolTipText     =   "Filter"
            ImageKey        =   "Filter"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Search"
            Object.ToolTipText     =   "Search"
            ImageKey        =   "Search"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "List"
            Object.ToolTipText     =   "List all records"
            ImageKey        =   "List"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList imlSmall 
      Left            =   240
      Top             =   60
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":00DC
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":03F8
            Key             =   "Warning"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":0720
            Key             =   "List"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":0A48
            Key             =   "xNew"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":31FC
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":3650
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":3AA4
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":4570
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":4898
            Key             =   "Sort"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":4CEC
            Key             =   "SQL"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":5140
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":5594
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":59EC
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":5B48
            Key             =   "Blank"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":5CA4
            Key             =   "NewRecord"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblIDlabel 
      AutoSize        =   -1  'True
      Caption         =   "ID:"
      Height          =   192
      Left            =   120
      TabIndex        =   9
      Top             =   4020
      Width           =   192
   End
   Begin VB.Label lblID 
      AutoSize        =   -1  'True
      Caption         =   "lblID"
      DataSource      =   "dtaData"
      Height          =   192
      Left            =   408
      TabIndex        =   8
      Top             =   4020
      Width           =   324
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileReport 
         Caption         =   "&Report"
      End
      Begin VB.Menu mnuFileSQL 
         Caption         =   "&SQL"
      End
      Begin VB.Menu mnuFileSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuRecords 
      Caption         =   "&Records"
      Begin VB.Menu mnuRecordsNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuRecordsModify 
         Caption         =   "&Modify"
      End
      Begin VB.Menu mnuRecordsDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuRecordsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecordsRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuRecordsSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecordsFilter 
         Caption         =   "&Filter"
      End
      Begin VB.Menu mnuRecordsSearch 
         Caption         =   "&Search"
      End
      Begin VB.Menu mnuRecordsSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecordsList 
         Caption         =   "&List"
      End
   End
End
Attribute VB_Name = "frmImages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const ChunkSize As Long = 4096
Public WithEvents rsMain As ADODB.Recordset
Attribute rsMain.VB_VarHelpID = -1
Private Sub cmdCancel_Click()
    CancelCommand Me, rsMain
End Sub
Private Sub cmdOK_Click()
    OKCommand Me, rsMain
End Sub
Private Sub Form_Load()
    Set adoConn = New ADODB.Connection
    adoConn.Open "FileDSN=" & gstrFileDSN
    
    Set rsMain = New ADODB.Recordset
    rsMain.CursorLocation = adUseClient
    SQLmain = "select * from [Images]"
    SQLfilter = vbNullString
    SQLkey = "ID"
    rsMain.Open SQLmain, adoConn, adOpenKeyset, adLockBatchOptimistic
    DBcollection.Add "rsMain", rsMain
    
    Set adodcMain.Recordset = rsMain
    BindField lblID, "ID", rsMain
    BindField txtName, "Name", rsMain
    BindField txtURL, "URL", rsMain
    'BindField picImage, "Image", rsMain
    BindField chkThumbnail, "Thumbnail", rsMain
    BindField rtxtCaption, "Caption", rsMain

    dbcTable.Enabled = False
    dbcRecord.Enabled = False
    dbcThumbnail.Enabled = False
    
    Set tsImages.SelectedItem = tsImages.Tabs(1)
    ProtectFields Me
    mode = modeDisplay
    fTransaction = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Cancel = CloseConnection(Me)
End Sub
Private Sub mnuRecordsFilter_Click()
    FilterCommand Me, rsMain, SQLkey
End Sub
Private Sub mnuRecordsDelete_Click()
    DeleteCommand Me, rsMain
End Sub
Private Sub mnuRecordsList_Click()
    ListCommand Me, rsMain
End Sub
Private Sub mnuRecordsModify_Click()
    ModifyCommand Me
    
    Set tsImages.SelectedItem = tsImages.Tabs(1)
    txtName.SetFocus
End Sub
Private Sub mnuRecordsNew_Click()
    NewCommand Me, rsMain
    
    Set tsImages.SelectedItem = tsImages.Tabs(1)
    txtName.SetFocus
End Sub
Private Sub mnuRecordsRefresh_Click()
    RefreshCommand rsMain, SQLkey
End Sub
Private Sub mnuRecordsSearch_Click()
    SearchCommand Me, rsMain, SQLkey
End Sub
Private Sub mnuFileExit_Click()
    Unload Me
End Sub
Private Sub mnuFileReport_Click()
    ReportCommand Me, rsMain, App.Path & "\Reports\Images.rpt"
End Sub
Private Sub mnuFileSQL_Click()
    SQLCommand "Images"
End Sub
Private Sub rsMain_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Dim Caption As String
    Dim strTempFile As String
    
    If Not pRecordset.BOF And Not pRecordset.EOF Then
        Caption = "Reference #" & pRecordset.Bookmark & ": " & pRecordset("ID") & ": " & pRecordset("Name")
        
        strTempFile = App.Path & "\Images\temp.dat"
        If DecodeImage(strTempFile) Then
            picImage.Picture = LoadPicture(strTempFile)
            Kill strTempFile
        End If
        If Not rsMain Is Nothing Then
            picImage.PaintPicture picImage.Picture, 0, 0, picImage.Width, picImage.Height
        End If
    End If
    UpdatePosition Me, Caption, pRecordset
End Sub
Private Sub tbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Report"
            mnuFileReport_Click
        Case "SQL"
            mnuFileSQL_Click
        Case "New"
            mnuRecordsNew_Click
        Case "Modify"
            mnuRecordsModify_Click
        Case "Delete"
            mnuRecordsDelete_Click
        Case "Refresh"
            mnuRecordsRefresh_Click
        Case "Filter"
            mnuRecordsFilter_Click
        Case "Search"
            mnuRecordsSearch_Click
        Case "List"
            mnuRecordsList_Click
    End Select
End Sub
'=================================================================================
Private Sub cmdLoad_Click()
    Dim strImagePath As String
    
    If Not IsNull(rsMain("Image")) Then
        If MsgBox("Are you sure you want to replace this image with another?", vbInformation + vbYesNo, Me.Caption) = vbNo Then
            Exit Sub
        End If
    End If
    
    With dlgImages
        .DialogTitle = "Select New Image"
        .FileName = vbNullString
        .Filter = "All Picture Files|*.jpg;*.gif;*.bmp;*.dib;*.ico;*.cur;*.wmf;*.emf|JPEG Images (*.jpg)|*.jpg|CompuServe GIF Images (*.gif)|*.gif|Windows Bitmaps (*.bmp;*.dib)|*.bmp;*.dib|Icons (*.ico;*.cur)|*.ico;*.cur|Metafiles (*.wmf;*.emf)|*.wmf;*.emf|All Files (*.*)|*.*"
        .FilterIndex = 1
        .ShowOpen
        strImagePath = .FileName
    End With
    picImage.Picture = LoadPicture(strImagePath)
    If Not EncodeImage(strImagePath) Then MsgBox "Unable to encode image", vbExclamation, Me.Caption
End Sub
Private Sub cmdView_Click()
'    Dim strTemp As String
'    If rsMain.BOF Or rsMain.EOF Then Exit Sub
'    strTemp = ParsePath(gstrFileDSN, DrvDir) & "temp.jpg"
'    If DecodeImage(strTemp) Then
'        Load frmPicture
'        frmPicture.strPictureFile = strTemp
'        frmPicture.Show vbModal
'    End If
'    Load frmPicture
    frmPicture.Show vbModal
End Sub
Private Function DecodeImage(ByVal strTempFile As String) As Boolean
    Dim FileUnit As Integer
    Dim Bytes As Long
    Dim BytesLeft As Long
    Dim strData As Variant
    
    BytesLeft = rsMain("Image").ActualSize
    If BytesLeft = 0 Then
        DecodeImage = False
        Exit Function
    End If
    
    picImage.Picture = Nothing
    FileUnit = FreeFile
    Open strTempFile For Binary Access Write As #FileUnit
    While BytesLeft > 0
        If BytesLeft > ChunkSize Then
            Bytes = ChunkSize
        Else
            Bytes = BytesLeft
        End If
        strData = rsMain("Image").GetChunk(Bytes)
        BytesLeft = BytesLeft - Bytes
        Put #FileUnit, , strData
    Wend
    Close #FileUnit
    DecodeImage = True
End Function
Private Function EncodeImage(ByVal strImageFile As String) As Boolean
    Dim FileUnit As Integer
    Dim Bytes As Long
    Dim BytesLeft As Long
    Dim bData() As Byte
    
    FileUnit = FreeFile
    Open strImageFile For Binary Access Read As #FileUnit
    BytesLeft = FileLen(strImageFile)
    If BytesLeft = 0 Then
        EncodeImage = False
        GoTo ExitSub
    End If
    
    While BytesLeft > 0
        If BytesLeft > ChunkSize Then
            Bytes = ChunkSize
        Else
            Bytes = BytesLeft
        End If
        ReDim bData(1 To Bytes)
        Get #FileUnit, , bData()
        rsMain("Image").AppendChunk bData()
        BytesLeft = BytesLeft - Bytes
    Wend
    EncodeImage = True
    
ExitSub:
    Close #FileUnit
End Function
Private Sub tsImages_Click()
    Dim i As Integer
    Dim strTempFile As String
    
    With tsImages
        For i = 0 To .Tabs.Count - 1
            If i = .SelectedItem.Index - 1 Then
                fraImages(i).Enabled = True
                fraImages(i).ZOrder
                If i = .Tabs.Count - 1 And Not rsMain.EOF Then
                    If IsNull(rsMain("Image")) Then Exit Sub
                    strTempFile = ParsePath(gstrFileDSN, DrvDir) & "temp.dat"
                    If DecodeImage(strTempFile) Then
                        picImage.Picture = LoadPicture(strTempFile)
                        Kill strTempFile
                    End If
                    If Not rsMain Is Nothing Then
                        picImage.PaintPicture picImage.Picture, 0, 0, picImage.Width, picImage.Height
                    End If
                End If
            Else
                fraImages(i).Enabled = False
            End If
        Next
    End With
End Sub
Private Sub txtURL_GotFocus()
    TextSelected
End Sub
Private Sub txtName_GotFocus()
    TextSelected
End Sub
Private Sub txtName_Validate(Cancel As Boolean)
    If Not txtName.Enabled Then Exit Sub
    If txtName.Text = vbNullString Then
        MsgBox "Name must be specified!", vbExclamation, Me.Caption
        Cancel = True
    End If
End Sub
