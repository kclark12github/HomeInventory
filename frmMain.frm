VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ken's Stuff..."
   ClientHeight    =   4230
   ClientLeft      =   30
   ClientTop       =   495
   ClientWidth     =   5925
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   5925
   Begin VB.PictureBox picWindow 
      Height          =   3792
      Left            =   0
      ScaleHeight     =   3735
      ScaleWidth      =   5655
      TabIndex        =   3
      Top             =   0
      Width           =   5712
      Begin VB.PictureBox picImage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   10425
         Left            =   0
         Picture         =   "frmMain.frx":2CFA
         ScaleHeight     =   10395
         ScaleWidth      =   13500
         TabIndex        =   4
         Top             =   0
         Width           =   13530
      End
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   3975
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9102
            Key             =   "DatabasePath"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   2
            Object.Width           =   1270
            MinWidth        =   1270
            TextSave        =   "3:55 PM"
            Key             =   "Time"
         EndProperty
      EndProperty
   End
   Begin VB.HScrollBar scrollH 
      Height          =   192
      LargeChange     =   1000
      Left            =   0
      SmallChange     =   100
      TabIndex        =   1
      Top             =   3780
      Width           =   5712
   End
   Begin VB.VScrollBar scrollV 
      Height          =   3792
      LargeChange     =   1000
      Left            =   5700
      SmallChange     =   100
      TabIndex        =   0
      Top             =   0
      Width           =   192
   End
   Begin MSComDlg.CommonDialog dlgMain 
      Left            =   120
      Top             =   60
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.Timer timMain 
      Left            =   2340
      Top             =   1860
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOptions 
         Caption         =   "&Options"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuDatabase 
      Caption         =   "&DataBase"
      Begin VB.Menu mnuDataBaseBooks 
         Caption         =   "&Books"
      End
      Begin VB.Menu mnuDataBaseCollectables 
         Caption         =   "&Collectables"
         Begin VB.Menu mnuDataBaseCollectablesCars 
            Caption         =   "&Cars (Hot Wheels, Matchbox, etc.)"
         End
         Begin VB.Menu mnuDataBaseCollectablesStarWars 
            Caption         =   "&Star Wars"
         End
         Begin VB.Menu mnuDataBaseCollectablesTYCollectables 
            Caption         =   "&TY Collectables (Beanie Babies, etc.)"
         End
      End
      Begin VB.Menu mnuDataBaseHobby 
         Caption         =   "&Hobby"
         Begin VB.Menu mnuDataBaseHobbyKits 
            Caption         =   "&Kits"
         End
         Begin VB.Menu mnuDataBaseHobbyDecals 
            Caption         =   "&Decals"
         End
         Begin VB.Menu mnuDataBaseHobbyDetailSets 
            Caption         =   "Detai&l Sets"
         End
         Begin VB.Menu mnuDataBaseHobbyFinishingProducts 
            Caption         =   "&Finishing Products"
         End
         Begin VB.Menu mnuDataBaseHobbySep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDataBaseHobbyTools 
            Caption         =   "&Tools"
         End
         Begin VB.Menu mnuDataBaseHobbyVideoResearch 
            Caption         =   "&Video Research"
         End
         Begin VB.Menu mnuDataBaseHobbySep2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDataBaseHobbyRockets 
            Caption         =   "&Rockets"
         End
         Begin VB.Menu mnuDataBaseHobbyTrains 
            Caption         =   "T&rains"
         End
         Begin VB.Menu mnuDataBaseHobbySep3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDataBaseHobbyCompanies 
            Caption         =   "&Companies"
         End
         Begin VB.Menu mnuDataBaseHobbySep4 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDataBaseHobbyAircraftDesignations 
            Caption         =   "Aircraft Designations"
         End
         Begin VB.Menu mnuDataBaseHobbyBlueAngelsHistory 
            Caption         =   "&Blue Angels History"
         End
      End
      Begin VB.Menu mnuDataBaseImages 
         Caption         =   "&Images"
      End
      Begin VB.Menu mnuDataBaseMusic 
         Caption         =   "&Music"
      End
      Begin VB.Menu mnuDataBaseSoftware 
         Caption         =   "&Software"
      End
      Begin VB.Menu mnuDataBaseUSNavyShips 
         Caption         =   "&US Navy Ships"
         Begin VB.Menu mnuDataBaseUSNavyShipsShips 
            Caption         =   "&Ships"
         End
         Begin VB.Menu mnuDataBaseUSNavyShipsSep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDataBaseUSNavyShipsClasses 
            Caption         =   "&Classes"
         End
         Begin VB.Menu mnuDataBaseUSNavyShipsClassifications 
            Caption         =   "Classificatio&ns"
         End
      End
      Begin VB.Menu mnuDataBaseVideoLibrary 
         Caption         =   "&Video Library"
         Begin VB.Menu mnuDataBaseVideoLibraryMovies 
            Caption         =   "&Movies"
         End
         Begin VB.Menu mnuDataBaseVideoLibrarySpecials 
            Caption         =   "&Specials"
         End
         Begin VB.Menu mnuDataBaseVideoLibraryTVEpisodes 
            Caption         =   "&TV Episodes"
         End
      End
      Begin VB.Menu mnuDataBaseKFC 
         Caption         =   "&WebLinks"
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
Public saveTop As Single
Public saveLeft As Single
Public saveCaption As String
Public saveIcon As Variant
Private fActivated As Boolean
Private fSplash As Boolean
Private Sub DoMenu(frm As Form, Caption, Optional Modal = vbModal)
    Me.MousePointer = vbHourglass
    saveTop = Me.Top
    saveLeft = Me.Left
    saveCaption = Me.Caption
    Set saveIcon = Me.Icon
'    frmMain.Hide
    
    fSplash = True
    Load frmSplash
    frmSplash.lblActivity.Caption = Caption
    frmSplash.Show vbModeless, frmMain
    DoEvents
    
    'Effectively hide frmMain, but keep it's TaskBar icon intact...
    Me.Top = -Me.Height
    Me.Left = -Me.Width
    DoEvents
    'Me.ShowInTaskbar = False
    
    'Load the target form...
    Load frm
    'Me.ShowInTaskbar = True
    Me.Caption = saveCaption & " - " & frm.Caption
    Set Me.Icon = frm.Icon
    Me.MousePointer = vbDefault
    frmSplash.Hide
    DoEvents
    fSplash = False
'    frm.ShowInTaskbar = True
    frm.Show Modal
'    frmMain.Show vbModal
End Sub
Private Sub ShowMain()
    If Me.Top < 0 Or Me.Top > Screen.Height Then Me.Top = saveTop Else saveTop = Me.Top
    If Me.Left < 0 Or Me.Left > Screen.Width Then Me.Left = saveLeft Else saveLeft = Me.Left
    Me.Caption = saveCaption
    Set Me.Icon = saveIcon
End Sub
Private Sub LoadBackground()
    Dim iWidth As Integer
    Dim iHeight As Integer
    Dim scWidth As Integer
    Dim scHeight As Integer
    Dim borderWidth As Integer
    Dim borderHeight As Integer
    Dim NeedHBar As Boolean
    Dim NeedVBar As Boolean
    
    On Error Resume Next
    scWidth = Screen.Width / Screen.TwipsPerPixelX
    scHeight = Screen.Height / Screen.TwipsPerPixelY
    
    borderWidth = Me.Width - Me.ScaleWidth
    borderHeight = Me.Height - Me.ScaleHeight
    
    picImage.Picture = LoadPicture(gstrImagePath)
    If Err.Number <> 0 Then
        MsgBox Err.Description & " (" & Err.Number & "). Using default image...", vbExclamation, Me.Caption
        gstrImagePath = gstrDefaultImagePath & "\" & gstrDefaultImage
        picImage.Picture = LoadPicture(gstrImagePath)
        If Err.Number <> 0 Then
            MsgBox Err.Description & " (" & Err.Number & "). Bagging this image crap... We didn't need no stinking images anyway...", vbExclamation, Me.Caption
            Exit Sub
        End If
        SaveSetting App.FileDescription, "Environment", "ImagePath", gstrImagePath
    End If
    picImage.Move 0, 0
    
    'Everything is governed by the size of the picture...
    iWidth = picImage.Width + borderWidth
    iHeight = borderHeight + picImage.Height
    
    scrollH.Visible = False
    If iWidth < iMinWidth Then
        iWidth = iMinWidth
    ElseIf iWidth >= Screen.Width Then
        iWidth = Screen.Width
        scrollH.Visible = True
        scrollH.Value = 0
    End If
    
    ScrollV.Visible = False
    If iHeight < iMinHeight Then
        iHeight = iMinHeight
    ElseIf iHeight > Screen.Height Then
        iHeight = Screen.Height
        ScrollV.Visible = True
        ScrollV.Value = 0
    End If
    Me.Width = iWidth
    Me.Height = iHeight
    
    iWidth = iWidth - borderWidth
    If ScrollV.Visible Then iWidth = ScrollV.Left
    iHeight = iHeight - borderHeight
    If scrollH.Visible Then iHeight = scrollH.Top
    picWindow.Move 0, 0, iWidth, iHeight
    
    'Center form...
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub
Private Sub Form_Activate()
    If Not fSplash Then ShowMain
    If fActivated Then Exit Sub
    fActivated = True
    
    Me.MousePointer = vbHourglass
    gstrFileDSN = GetSetting(App.FileDescription, "Environment", "FileDSN", "")
    If gstrFileDSN = vbNullString Then
        mnuFileOptions_Click
    End If
    
    If ParsePath(gstrFileDSN, DrvDirNoSlash) = gstrODBCFileDSNDir Then
        sbStatus.Panels("DatabasePath").Text = ParsePath(gstrFileDSN, FileNameBase)
    Else
        sbStatus.Panels("DatabasePath").Text = ParsePath(gstrFileDSN, DrvDirFileNameBase)
    End If
    Me.Caption = ParsePath(gstrFileDSN, FileNameBase) & "..."
    Me.MousePointer = vbDefault
End Sub
Private Sub Form_Load()
    Dim iPos As Integer
    Dim rTop As Single
    Dim rLeft As Single
    Dim rHeight As Single
    Dim rWidth As Single
    
    fActivated = False
    MinWidth = iMinWidth
    MinHeight = iMinHeight
    
    gstrODBCFileDSNDir = VbRegQueryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\ODBC File DSN", "DefaultDSNDir")
    gstrDefaultImagePath = App.Path & "\Images"
    gstrImagePath = GetSetting(App.FileDescription, "Environment", "ImagePath", gstrDefaultImagePath & "\" & gstrDefaultImage)
    LoadBackground
    gfUseFilterMethod = GetSetting(App.FileDescription, "Environment", "UseFilterMethod", False)
    
    If GetSetting(App.FileDescription, "Environment", "DimensionsSaved", False) Then
        rTop = GetSetting(App.FileDescription, "Environment", "Top")
        rLeft = GetSetting(App.FileDescription, "Environment", "Left")
        rHeight = GetSetting(App.FileDescription, "Environment", "Height")
        rWidth = GetSetting(App.FileDescription, "Environment", "Width")
    End If
    If rTop > 0 And rTop <= Screen.Height Then Me.Top = rTop
    If rLeft > 0 And rLeft <= Screen.Height Then Me.Left = rLeft
    saveTop = Me.Top
    saveLeft = Me.Left
    saveCaption = Me.Caption
    Set saveIcon = Me.Icon
    
    gfTraceMode = GetSetting(App.FileDescription, "Environment", "TraceMode", False)
    gstrTraceFile = GetSetting(App.FileDescription, "Environment", "TraceFile", ParsePath(App.Path, DrvDir) & "Trace.log")
    If gfTraceMode Then
        Call Trace(trcBody, String(132, "="))
        Call Trace(trcBody, App.FileDescription & " Start - " & gstrTraceFile)
    End If
    
    DBcollection.Clear
    
    SetDateFormats
    fmtDate = fmtFullDateTime
End Sub
Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        If scrollH.Visible Then
            scrollH.Top = Me.ScaleHeight - scrollH.Height - sbStatus.Height
            scrollH.Left = 0
            scrollH.Width = Me.ScaleWidth
            If ScrollV.Visible Then scrollH.Width = scrollH.Width - ScrollV.Width
            scrollH.Max = picImage.Width - Me.ScaleWidth
            scrollH.SmallChange = picImage.Width / 100
            scrollH.LargeChange = picImage.Width / 20
            picWindow.Height = scrollH.Top
        End If
        
        If ScrollV.Visible Then
            ScrollV.Top = 0
            ScrollV.Left = Me.ScaleWidth - ScrollV.Width
            ScrollV.Height = Me.ScaleHeight - sbStatus.Height
            If scrollH.Visible Then ScrollV.Height = ScrollV.Height - scrollH.Height
            ScrollV.Max = picImage.Height - Me.ScaleHeight
            ScrollV.SmallChange = picImage.Height / 100
            ScrollV.LargeChange = picImage.Height / 20
            picWindow.Width = ScrollV.Left
        End If
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set DBcollection = Nothing
    Call Trace(trcBody, App.FileDescription & " Exit.")
    Call Trace(trcBody, String(132, "="))
    Call ShowMain
    Call SaveSetting(App.FileDescription, "Environment", "DimensionsSaved", True)
    Call SaveSetting(App.FileDescription, "Environment", "Top", Me.Top)
    Call SaveSetting(App.FileDescription, "Environment", "Left", Me.Left)
    Call SaveSetting(App.FileDescription, "Environment", "Height", Me.Height)
    Call SaveSetting(App.FileDescription, "Environment", "Width", Me.Width)
End Sub
Private Sub mnuDataBaseBooks_Click()
    Call DoMenu(frmBooks, "Books")
End Sub
Private Sub mnuDataBaseHobbyAircraftDesignations_Click()
    Call DoMenu(frmAircraftDesignations, "Hobby \ Aircraft Designations")
End Sub
Private Sub mnuDataBaseHobbyBlueAngelsHistory_Click()
    Call DoMenu(frmBlueAngelsHistory, "Hobby \ Blue Angels History")
End Sub
Private Sub mnuDataBaseHobbyCompanies_Click()
    Call DoMenu(frmCompanies, "Hobby \ Companies")
End Sub
Private Sub mnuDataBaseHobbyDecals_Click()
    Call DoMenu(frmDecals, "Hobby \ Decals")
End Sub
Private Sub mnuDataBaseHobbyDetailSets_Click()
    Call DoMenu(frmDetailSets, "Bobby \ Detail Sets")
End Sub
Private Sub mnuDataBaseHobbyKits_Click()
    Call DoMenu(frmKits, "Hobby \ Kits")
End Sub
Private Sub mnuDataBaseHobbyFinishingProducts_Click()
    Call DoMenu(frmFinishingProducts, "Hobby \ Finishing Products")
End Sub
Private Sub mnuDataBaseHobbyRockets_Click()
    Call DoMenu(frmRockets, "Hobby \ Rockets")
End Sub
Private Sub mnuDataBaseHobbyTools_Click()
    Call DoMenu(frmTools, "Hobby \ Tools")
End Sub
Private Sub mnuDataBaseHobbyTrains_Click()
    Call DoMenu(frmTrains, "Hobby \ Trains")
End Sub
Private Sub mnuDataBaseHobbyVideoResearch_Click()
    Call DoMenu(frmVideoResearch, "Hobby \ Video Research")
End Sub
Private Sub mnuDataBaseImages_Click()
    Call DoMenu(frmImages, "Images")
End Sub
Private Sub mnuDataBaseKFC_Click()
    Call DoMenu(frmWebLinks, "KFC") ', vbModeless)
End Sub
Private Sub mnuDataBaseMusic_Click()
    Call DoMenu(frmMusic, "Music")
End Sub
Private Sub mnuDataBaseSoftware_Click()
    Call DoMenu(frmSoftware, "Software")
End Sub
Private Sub mnuDataBaseUSNavyShipsClasses_Click()
    Call DoMenu(frmUSNClasses, "US Navy Ships \ Classes")
End Sub
Private Sub mnuDataBaseUSNavyShipsClassifications_Click()
    Call DoMenu(frmUSNClassifications, "US Navy Ships \ Classifications")
End Sub
Private Sub mnuDataBaseUSNavyShipsShips_Click()
    Call DoMenu(frmUSNShips, "US Navy Ships \ Ships")
End Sub
Private Sub mnuDataBaseVideoLibraryMovies_Click()
    Call DoMenu(frmMovies, "Video Library \ Movies")
End Sub
Private Sub mnuDataBaseVideoLibrarySpecials_Click()
    Call DoMenu(frmSpecials, "Video Library \ Specials")
End Sub
Private Sub mnuDataBaseVideoLibraryTVEpisodes_Click()
    Call DoMenu(frmTVEpisodes, "Video Library \ TV Episodes")
End Sub
Private Sub mnuFileOptions_Click()
'    Me.MousePointer = vbHourglass
'    Load frmOptions
'    Me.MousePointer = vbDefault
    frmOptions.Show vbModal
    LoadBackground
End Sub
Private Sub mnuFileExit_Click()
    Unload Me
    End
End Sub
Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal
End Sub
Private Sub picImage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbKeyRButton Then PopupMenu mnuFile
End Sub
Private Sub scrollH_Change()
    picImage.Left = -scrollH.Value
End Sub
Private Sub scrollV_Change()
    picImage.Top = -ScrollV.Value
End Sub
