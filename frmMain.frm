VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Home Inventory"
   ClientHeight    =   912
   ClientLeft      =   36
   ClientTop       =   492
   ClientWidth     =   2112
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912
   ScaleWidth      =   2112
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgMain 
      Left            =   120
      Top             =   60
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.PictureBox picBackground 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   8364
      Left            =   60
      Picture         =   "frmMain.frx":030A
      ScaleHeight     =   8316
      ScaleWidth      =   10800
      TabIndex        =   0
      Top             =   60
      Width           =   10848
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileBackground 
         Caption         =   "&Select Background"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuDataBase 
      Caption         =   "&DataBase"
      Begin VB.Menu mnuDataBaseBooks 
         Caption         =   "&Books..."
      End
      Begin VB.Menu mnuDataBaseHobby 
         Caption         =   "&Hobby..."
      End
      Begin VB.Menu mnuDataBaseMusic 
         Caption         =   "&Music..."
      End
      Begin VB.Menu mnuDataBaseSoftware 
         Caption         =   "&Software"
      End
      Begin VB.Menu mnuDataBaseUSNavyShips 
         Caption         =   "&US Navy Ships..."
      End
      Begin VB.Menu mnuDataBaseVideoTapes 
         Caption         =   "&Video Tapes..."
      End
      Begin VB.Menu mnuDataBaseKFC 
         Caption         =   "&WebLinks (KFC)..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About Home Inventory..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cnKFC As New ADODB.Connection
Private cmdKFC As New ADODB.Command
Const gstrProvider = "Microsoft.Jet.OLEDB.3.51"
'Const gstrConnectionString = "E:\WebShare\wwwroot\Access\KFC.mdb"
Const gstrRunTimeUserName = "admin"
Const gstrRunTimePassword = ""
Const gstrDBlocation = "E:\WebShare\wwwroot\Access\"
Const gstrDefaultImagePath = "E:\WebShare\wwwroot\Aircraft\Fighter Aircraft\F-14 Tomcat\F14_102.jpg"
Const iMinWidth = 2184
Const iMinHeight = 1440

Public gstrImagePath As String
Public DBcollection As New DataBaseCollection
Public Enum ActionMode
    modeDisplay = 0
    modeAdd = 1
    modeModify = 2
    modeDelete = 3
End Enum
Public Sub OpenFields(pForm As Form)
    Dim ctl As Control
    For Each ctl In pForm.Controls
        Select Case TypeName(ctl)
            Case "TextBox", "DataCombo", "ComboBox"
                'ctl.Locked = False
                ctl.Enabled = True
                ctl.BackColor = vbWindowBackground
            Case "CheckBox"
                ctl.Enabled = True
        End Select
    Next ctl
End Sub
Public Sub ProtectFields(pForm As Form)
    Dim ctl As Control
    For Each ctl In pForm.Controls
        Select Case TypeName(ctl)
            Case "TextBox", "DataCombo", "ComboBox"
                'ctl.Locked = True
                ctl.Enabled = False
                ctl.BackColor = vbButtonFace
            Case "CheckBox"
                ctl.Enabled = False
        End Select
    Next ctl
End Sub
Private Sub LoadBackground()
    Dim iWidth As Integer
    Dim iHeight As Integer
    Dim scWidth As Integer
    Dim scHeight As Integer
    Dim borderWidth As Integer
    Dim borderHeight As Integer
    
    scWidth = Screen.Width / Screen.TwipsPerPixelX
    scHeight = Screen.Height / Screen.TwipsPerPixelY
    
    borderWidth = Me.Width - Me.ScaleWidth
    borderHeight = Me.Height - Me.ScaleHeight
    
    picBackground.Picture = LoadPicture(gstrImagePath)
    picBackground.Move Me.ScaleLeft, Me.ScaleTop
    iWidth = picBackground.Width + (2 * picBackground.Left) + borderWidth
    iHeight = picBackground.Height + (2 * picBackground.Top) + borderHeight
    If iWidth < iMinWidth Then iWidth = iMinWidth
    If iHeight < iMinHeight Then iHeight = iMinHeight
    Me.Width = iWidth
    Me.Height = iHeight
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub
Private Sub LoadDBcoll(DBname As String)
    DBcollection.Add DBname, DBname, gstrDBlocation & DBname & ".mdb", gstrProvider, gstrRunTimeUserName, gstrRunTimePassword, DBname
End Sub
Private Sub Form_Load()
    LoadDBcoll "Books"
    LoadDBcoll "Hobby"
    LoadDBcoll "KFC"
    LoadDBcoll "Music"
    LoadDBcoll "Software"
    LoadDBcoll "US Navy Ships"
    LoadDBcoll "UserAccessInfo"
    LoadDBcoll "VideoTapes"
    
    gstrImagePath = GetSetting(App.FileDescription, "Environment", "ImagePath", gstrDefaultImagePath)
    LoadBackground
End Sub
Private Sub mnuDataBaseBooks_Click()
    frmBooks.Show vbModal
End Sub
Private Sub mnuFileBackground_Click()
    Dim CurrentPath As String
    Dim CurrentDrive As String
    Dim CurrentImage As String
    
    CurrentPath = ParsePath(gstrImagePath, DrvDirNoSlash)
    CurrentDrive = ParsePath(gstrImagePath, DrvOnly)
    CurrentImage = ParsePath(gstrImagePath, FileNameBaseExt)
    ChDrive CurrentDrive
    ChDir CurrentPath
    With dlgMain
        .DialogTitle = "Select New Background Image"
        .FileName = CurrentImage
        .Filter = "All Picture Files|*.jpg;*.gif;*.bmp;*.dib;*.ico;*.cur;*.wmf;*.emf|JPEG Images (*.jpg)|*.jpg|CompuServe GIF Images (*.gif)|*.gif|Windows Bitmaps (*.bmp;*.dib)|*.bmp;*.dib|Icons (*.ico;*.cur)|*.ico;*.cur|Metafiles (*.wmf;*.emf)|*.wmf;*.emf|All Files (*.*)|*.*"
        .FilterIndex = 1
        .ShowOpen    ' Call the open file procedure.
        gstrImagePath = .FileName
        SaveSetting App.FileDescription, "Environment", "ImagePath", gstrImagePath
    End With
    LoadBackground
End Sub
Private Sub mnuFileExit_Click()
    Unload Me
End Sub
Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal
End Sub
