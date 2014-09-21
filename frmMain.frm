VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Home Inventory"
   ClientHeight    =   1920
   ClientLeft      =   36
   ClientTop       =   492
   ClientWidth     =   3612
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   3612
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
      Left            =   360
      Picture         =   "frmMain.frx":030A
      ScaleHeight     =   8316
      ScaleWidth      =   10800
      TabIndex        =   0
      Top             =   300
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
         Caption         =   "&Books"
      End
      Begin VB.Menu mnuDataBaseHobby 
         Caption         =   "&Hobby"
         Begin VB.Menu mnuDataBaseHobbyAircraftDesignations 
            Caption         =   "Aircraft Designations"
         End
         Begin VB.Menu mnuDataBaseHobbyAircraftModels 
            Caption         =   "&Aircraft Models"
         End
         Begin VB.Menu mnuDataBaseHobbyArmorCarModels 
            Caption         =   "Armor && &Car Models"
         End
         Begin VB.Menu mnuDataBaseHobbyBlueAngelsHistory 
            Caption         =   "&Blue Angels History"
         End
         Begin VB.Menu mnuDataBaseHobbyCompanies 
            Caption         =   "&Companies"
         End
         Begin VB.Menu mnuDataBaseHobbyDecals 
            Caption         =   "&Decals"
         End
         Begin VB.Menu mnuDataBaseHobbyDetailSets 
            Caption         =   "Detai&l Sets"
         End
         Begin VB.Menu mnuDataBaseHobbyNavalModels 
            Caption         =   "&Naval Models"
         End
         Begin VB.Menu mnuDataBaseHobbyPaintsBrushes 
            Caption         =   "&Paints && Brushes"
         End
         Begin VB.Menu mnuDataBaseHobbyRockets 
            Caption         =   "&Rockets"
         End
         Begin VB.Menu mnuDataBaseHobbySciFiSpaceModels 
            Caption         =   "&SciFi && Space Models"
         End
         Begin VB.Menu mnuDataBaseHobbyTools 
            Caption         =   "&Tools"
         End
         Begin VB.Menu mnuDataBaseHobbyTrains 
            Caption         =   "T&rains"
         End
         Begin VB.Menu mnuDataBaseHobbyVideoResearch 
            Caption         =   "&Video Research"
         End
      End
      Begin VB.Menu mnuDataBaseMusic 
         Caption         =   "&Music"
      End
      Begin VB.Menu mnuDataBaseSoftware 
         Caption         =   "&Software"
      End
      Begin VB.Menu mnuDataBaseUSNavyShips 
         Caption         =   "&US Navy Ships"
         Begin VB.Menu mnuUSNavyShipsClasses 
            Caption         =   "&Classes"
         End
         Begin VB.Menu mnuUSNavyShipsClassifications 
            Caption         =   "Classificatio&ns"
         End
         Begin VB.Menu mnuUSNavyShipsCommands 
            Caption         =   "C&ommands"
         End
         Begin VB.Menu mnuUSNavyShipsHomePorts 
            Caption         =   "&Home Ports"
         End
         Begin VB.Menu mnuUSNavyShipsShips 
            Caption         =   "&Ships"
         End
      End
      Begin VB.Menu mnuDataBaseVideoTapes 
         Caption         =   "&Video Tapes"
      End
      Begin VB.Menu mnuDataBaseKFC 
         Caption         =   "&WebLinks (KFC)"
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

Public frmReport As Form
Public rdcReport As CRAXDRT.Report
Public MinWidth As Integer
Public MinHeight As Integer

Public gstrImagePath As String
Public DBcollection As New DataBaseCollection
Public Enum ActionMode
    modeDisplay = 0
    modeAdd = 1
    modeModify = 2
    modeDelete = 3
End Enum
Public Sub BindField(ctl As Control, DataField As String, DataSource As ADODB.Recordset, Optional RowSource As ADODB.Recordset, Optional BoundColumn As String, Optional ListField As String)
    Select Case TypeName(ctl)
        Case "CheckBox", "Label", "TextBox"
            Set ctl.DataSource = DataSource
            ctl.DataField = DataField
        Case "DataCombo"
            Set ctl.DataSource = DataSource
            ctl.DataField = DataField
            Set ctl.RowSource = RowSource
            ctl.BoundColumn = BoundColumn
            ctl.ListField = ListField
    End Select
End Sub
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
    MinWidth = iMinWidth
    MinHeight = iMinHeight
    
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
Private Sub mnuDataBaseHobbyAircraftDesignations_Click()
    frmAircraftDesignations.Show vbModal
End Sub
Private Sub mnuDataBaseHobbyAircraftModels_Click()
    frmAircraftModels.Show vbModal
End Sub
Private Sub mnuDataBaseHobbyArmorCarModels_Click()
    frmArmorCarModels.Show vbModal
End Sub
Private Sub mnuDataBaseHobbyBlueAngelsHistory_Click()
    frmBlueAngelsHistory.Show vbModal
End Sub
Private Sub mnuDataBaseHobbyCompanies_Click()
    frmCompanies.Show vbModal
End Sub
Private Sub mnuDataBaseHobbyDecals_Click()
    frmDecals.Show vbModal
End Sub
Private Sub mnuDataBaseHobbyDetailSets_Click()
    frmDetailSets.Show vbModal
End Sub
Private Sub mnuDataBaseHobbyNavalModels_Click()
    frmNavalModels.Show vbModal
End Sub
Private Sub mnuDataBaseHobbyPaintsBrushes_Click()
    frmPaintsBrushes.Show vbModal
End Sub
Private Sub mnuDataBaseHobbyRockets_Click()
    frmRockets.Show vbModal
End Sub
Private Sub mnuDataBaseHobbySciFiSpaceModels_Click()
    frmSciFiNASAModels.Show vbModal
End Sub
Private Sub mnuDataBaseHobbyTools_Click()
    frmTools.Show vbModal
End Sub
Private Sub mnuDataBaseHobbyTrains_Click()
    frmTrains.Show vbModal
End Sub
Private Sub mnuDataBaseHobbyVideoResearch_Click()
    frmVideoResearch.Show vbModal
End Sub
Private Sub mnuDataBaseMusic_Click()
    frmMusic.Show vbModal
End Sub
Private Sub mnuDataBaseSoftware_Click()
    frmSoftware.Show vbModal
End Sub
Private Sub mnuDataBaseVideoTapes_Click()
    'frmVideoTapes.Show vbModal
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
Private Sub picBackground_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbKeyRButton Then PopupMenu mnuFile
End Sub
