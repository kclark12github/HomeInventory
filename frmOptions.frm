VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOptions 
   Caption         =   "Options"
   ClientHeight    =   1404
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   ScaleHeight     =   1404
   ScaleWidth      =   5820
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   372
      Left            =   2964
      TabIndex        =   7
      Top             =   900
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   1884
      TabIndex        =   6
      Top             =   900
      Width           =   972
   End
   Begin VB.CommandButton cmdBrowseDSN 
      Caption         =   "&Browse"
      Height          =   288
      Left            =   4020
      TabIndex        =   5
      Top             =   480
      Width           =   1572
   End
   Begin VB.TextBox txtDSN 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   288
      Left            =   1752
      TabIndex        =   3
      Top             =   480
      Width           =   2172
   End
   Begin VB.CommandButton cmdBrowseImages 
      Caption         =   "&Select New Image"
      Height          =   288
      Left            =   4020
      TabIndex        =   2
      Top             =   120
      Width           =   1572
   End
   Begin VB.TextBox txtBackground 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   288
      Left            =   1740
      TabIndex        =   0
      Top             =   120
      Width           =   2172
   End
   Begin MSComDlg.CommonDialog dlgOptions 
      Left            =   180
      Top             =   900
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.Label lblDSN 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Data Source Name:"
      Height          =   192
      Left            =   228
      TabIndex        =   4
      Top             =   528
      Width           =   1416
   End
   Begin VB.Label lblBackground 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Background Image:"
      Height          =   192
      Left            =   228
      TabIndex        =   1
      Top             =   168
      Width           =   1404
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strFileDSN As String
Dim strImagePath As String
Private Sub cmdBrowseDSN_Click()
    With dlgOptions
        .DialogTitle = "Select Database"
        .InitDir = gstrODBCFileDSNDir
        .FileName = strFileDSN
        .Filter = "File DSNs (*.dsn)|*.dsn|All Files (*.*)|*.*"
        .FilterIndex = 1
        .ShowOpen
        strFileDSN = .FileName
    End With
    txtDSN.Text = ParsePath(strFileDSN, FileNameBaseExt)
End Sub
Private Sub cmdBrowseImages_Click()
    Dim CurrentPath As String
    Dim CurrentDrive As String
    Dim CurrentImage As String
    
    CurrentPath = ParsePath(strImagePath, DrvDirNoSlash)
    CurrentDrive = ParsePath(strImagePath, DrvOnly)
    CurrentImage = ParsePath(strImagePath, FileNameBaseExt)
    ChDrive CurrentDrive
    ChDir CurrentPath
    With dlgOptions
        .DialogTitle = "Select New Background Image"
        .FileName = CurrentImage
        .Filter = "All Picture Files|*.jpg;*.gif;*.bmp;*.dib;*.ico;*.cur;*.wmf;*.emf|JPEG Images (*.jpg)|*.jpg|CompuServe GIF Images (*.gif)|*.gif|Windows Bitmaps (*.bmp;*.dib)|*.bmp;*.dib|Icons (*.ico;*.cur)|*.ico;*.cur|Metafiles (*.wmf;*.emf)|*.wmf;*.emf|All Files (*.*)|*.*"
        .FilterIndex = 1
        .ShowOpen    ' Call the open file procedure.
        strImagePath = .FileName
    End With
    txtBackground.Text = ParsePath(strImagePath, FileNameBaseExt)
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
    gstrFileDSN = strFileDSN
    SaveSetting App.FileDescription, "Environment", "FileDSN", strFileDSN
    gstrImagePath = strImagePath
    SaveSetting App.FileDescription, "Environment", "ImagePath", strImagePath
    Unload Me
End Sub
Private Sub Form_Activate()
    If strFileDSN = vbNullString Then cmdBrowseDSN_Click
End Sub
Private Sub Form_Load()
    strFileDSN = GetSetting(App.FileDescription, "Environment", "FileDSN", gstrFileDSN)
    txtDSN.Text = ParsePath(strFileDSN, FileNameBaseExt)
    strImagePath = GetSetting(App.FileDescription, "Environment", "ImagePath", gstrImagePath)
    txtBackground.Text = ParsePath(strImagePath, FileNameBaseExt)
End Sub
