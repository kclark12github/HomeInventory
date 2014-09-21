VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ken's Stuff..."
   ClientHeight    =   4236
   ClientLeft      =   36
   ClientTop       =   492
   ClientWidth     =   5916
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4236
   ScaleWidth      =   5916
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picWindow 
      Height          =   3792
      Left            =   0
      ScaleHeight     =   3744
      ScaleWidth      =   5664
      TabIndex        =   3
      Top             =   0
      Width           =   5712
      Begin VB.PictureBox picImage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   8340
         Left            =   0
         Picture         =   "frmMain.frx":2CFA
         ScaleHeight     =   8316
         ScaleWidth      =   10800
         TabIndex        =   4
         Top             =   0
         Width           =   10824
      End
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   2
      Top             =   3984
      Width           =   5916
      _ExtentX        =   10435
      _ExtentY        =   445
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
            TextSave        =   "12:08 AM"
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
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
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
   Begin VB.Menu mnuDataBase 
      Caption         =   "&DataBase"
      Begin VB.Menu mnuDataBaseBooks 
         Caption         =   "&Books"
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
Private fActivated As Boolean
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
    
    scrollV.Visible = False
    If iHeight < iMinHeight Then
        iHeight = iMinHeight
    ElseIf iHeight > Screen.Height Then
        iHeight = Screen.Height
        scrollV.Visible = True
        scrollV.Value = 0
    End If
    Me.Width = iWidth
    Me.Height = iHeight
    
    iWidth = iWidth - borderWidth
    If scrollV.Visible Then iWidth = scrollV.Left
    iHeight = iHeight - borderHeight
    If scrollH.Visible Then iHeight = scrollH.Top
    picWindow.Move 0, 0, iWidth, iHeight
    
    'Center form...
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub
Private Sub Form_Activate()
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
    
    fActivated = False
    MinWidth = iMinWidth
    MinHeight = iMinHeight
    
    gstrODBCFileDSNDir = VbRegQueryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\ODBC File DSN", "DefaultDSNDir")
    gstrDefaultImagePath = App.Path & "\Images"
    gstrImagePath = GetSetting(App.FileDescription, "Environment", "ImagePath", gstrDefaultImagePath & "\" & gstrDefaultImage)
    LoadBackground
    gfUseFilterMethod = GetSetting(App.FileDescription, "Environment", "UseFilterMethod", False)
    
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
            If scrollV.Visible Then scrollH.Width = scrollH.Width - scrollV.Width
            scrollH.Max = picImage.Width - Me.ScaleWidth
            scrollH.SmallChange = picImage.Width / 100
            scrollH.LargeChange = picImage.Width / 20
            picWindow.Height = scrollH.Top
        End If
        
        If scrollV.Visible Then
            scrollV.Top = 0
            scrollV.Left = Me.ScaleWidth - scrollV.Width
            scrollV.Height = Me.ScaleHeight - sbStatus.Height
            If scrollH.Visible Then scrollV.Height = scrollV.Height - scrollH.Height
            scrollV.Max = picImage.Height - Me.ScaleHeight
            scrollV.SmallChange = picImage.Height / 100
            scrollV.LargeChange = picImage.Height / 20
            picWindow.Width = scrollV.Left
        End If
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call Trace(trcBody, App.FileDescription & " Exit.")
    Call Trace(trcBody, String(132, "="))
End Sub
Private Sub mnuDataBaseBooks_Click()
    Me.MousePointer = vbHourglass
    Load frmBooks
    Me.MousePointer = vbDefault
    frmBooks.Show vbModal
End Sub
Private Sub mnuDataBaseHobbyAircraftDesignations_Click()
    Me.MousePointer = vbHourglass
    Load frmAircraftDesignations
    Me.MousePointer = vbDefault
    frmAircraftDesignations.Show vbModal
End Sub
Private Sub mnuDataBaseHobbyBlueAngelsHistory_Click()
    Me.MousePointer = vbHourglass
    Load frmBlueAngelsHistory
    Me.MousePointer = vbDefault
    frmBlueAngelsHistory.Show vbModal
End Sub
Private Sub mnuDataBaseHobbyCompanies_Click()
    Me.MousePointer = vbHourglass
    Load frmCompanies
    Me.MousePointer = vbDefault
    frmCompanies.Show vbModal
End Sub
Private Sub mnuDataBaseHobbyDecals_Click()
    Me.MousePointer = vbHourglass
    Load frmDecals
    Me.MousePointer = vbDefault
    frmDecals.Show vbModal
End Sub
Private Sub mnuDataBaseHobbyDetailSets_Click()
    Me.MousePointer = vbHourglass
    Load frmDetailSets
    Me.MousePointer = vbDefault
    frmDetailSets.Show vbModal
End Sub
Private Sub mnuDataBaseHobbyKits_Click()
    Me.MousePointer = vbHourglass
    Load frmKits
    Me.MousePointer = vbDefault
    frmKits.Show vbModal
End Sub
Private Sub mnuDataBaseHobbyFinishingProducts_Click()
    Me.MousePointer = vbHourglass
    Load frmFinishingProducts
    Me.MousePointer = vbDefault
    frmFinishingProducts.Show vbModal
End Sub
Private Sub mnuDataBaseHobbyRockets_Click()
    Me.MousePointer = vbHourglass
    Load frmRockets
    Me.MousePointer = vbDefault
    frmRockets.Show vbModal
End Sub
Private Sub mnuDataBaseHobbyTools_Click()
    Me.MousePointer = vbHourglass
    Load frmTools
    Me.MousePointer = vbDefault
    frmTools.Show vbModal
End Sub
Private Sub mnuDataBaseHobbyTrains_Click()
    Me.MousePointer = vbHourglass
    Load frmTrains
    Me.MousePointer = vbDefault
    frmTrains.Show vbModal
End Sub
Private Sub mnuDataBaseHobbyVideoResearch_Click()
    Me.MousePointer = vbHourglass
    Load frmVideoResearch
    Me.MousePointer = vbDefault
    frmVideoResearch.Show vbModal
End Sub
Private Sub mnuDataBaseImages_Click()
    Me.MousePointer = vbHourglass
    Load frmImages
    Me.MousePointer = vbDefault
    frmImages.Show vbModal
End Sub
Private Sub mnuDataBaseKFC_Click()
    frmWebLinks.Show vbModeless
End Sub
Private Sub mnuDataBaseMusic_Click()
    Me.MousePointer = vbHourglass
    Load frmMusic
    Me.MousePointer = vbDefault
    frmMusic.Show vbModal
End Sub
Private Sub mnuDataBaseSoftware_Click()
    Me.MousePointer = vbHourglass
    Load frmSoftware
    Me.MousePointer = vbDefault
    frmSoftware.Show vbModal
End Sub
Private Sub mnuDataBaseUSNavyShipsClasses_Click()
    Me.MousePointer = vbHourglass
    Load frmUSNClasses
    Me.MousePointer = vbDefault
    frmUSNClasses.Show vbModal
End Sub
Private Sub mnuDataBaseUSNavyShipsClassifications_Click()
    Me.MousePointer = vbHourglass
    Load frmUSNClassifications
    Me.MousePointer = vbDefault
    frmUSNClassifications.Show vbModal
End Sub
Private Sub mnuDataBaseUSNavyShipsShips_Click()
    Me.MousePointer = vbHourglass
    Load frmUSNShips
    Me.MousePointer = vbDefault
    frmUSNShips.Show vbModal
End Sub
Private Sub mnuDataBaseVideoLibraryMovies_Click()
    Me.MousePointer = vbHourglass
    Load frmMovies
    Me.MousePointer = vbDefault
    frmMovies.Show vbModal
End Sub
Private Sub mnuDataBaseVideoLibrarySpecials_Click()
    Me.MousePointer = vbHourglass
    Load frmSpecials
    Me.MousePointer = vbDefault
    frmSpecials.Show vbModal
End Sub
Private Sub mnuDataBaseVideoLibraryTVEpisodes_Click()
    Me.MousePointer = vbHourglass
    Load frmTVEpisodes
    Me.MousePointer = vbDefault
    frmTVEpisodes.Show vbModal
End Sub
Private Sub mnuFileOptions_Click()
    Me.MousePointer = vbHourglass
    Load frmOptions
    Me.MousePointer = vbDefault
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
    picImage.Top = -scrollV.Value
End Sub
