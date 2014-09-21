VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOptions 
   Caption         =   "Options"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   5820
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   5820
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraTrace 
      Caption         =   "Trace Information"
      Height          =   912
      Left            =   60
      TabIndex        =   9
      Top             =   1560
      Width           =   5532
      Begin VB.CommandButton cmdBrowseTraceFile 
         Caption         =   "Select &Trace File"
         Height          =   288
         Left            =   3888
         TabIndex        =   12
         Top             =   480
         Width           =   1572
      End
      Begin VB.TextBox txtTraceFile 
         BackColor       =   &H8000000F&
         Height          =   288
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   3672
      End
      Begin VB.CheckBox chkTraceMode 
         Caption         =   "Trace Mode"
         Height          =   252
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1392
      End
   End
   Begin VB.Frame fraBackground 
      Caption         =   "Background Image"
      Height          =   612
      Left            =   60
      TabIndex        =   6
      Top             =   120
      Width           =   5532
      Begin VB.TextBox txtBackground 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   288
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   3672
      End
      Begin VB.CommandButton cmdBrowseImages 
         Caption         =   "&Select New Image"
         Height          =   288
         Left            =   3900
         TabIndex        =   7
         Top             =   240
         Width           =   1572
      End
   End
   Begin VB.Frame fraDSN 
      Caption         =   "Data Source Name"
      Height          =   612
      Left            =   60
      TabIndex        =   3
      Top             =   840
      Width           =   5532
      Begin VB.TextBox txtDSN 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   288
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3672
      End
      Begin VB.CommandButton cmdBrowseDSN 
         Caption         =   "&Browse"
         Height          =   288
         Left            =   3888
         TabIndex        =   4
         Top             =   240
         Width           =   1572
      End
   End
   Begin VB.CheckBox chkUseFilterMethod 
      Caption         =   "Use Filter Method"
      Height          =   252
      Left            =   1824
      TabIndex        =   2
      Top             =   2520
      Width           =   2172
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   372
      Left            =   2904
      TabIndex        =   1
      Top             =   2880
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   1824
      TabIndex        =   0
      Top             =   2880
      Width           =   972
   End
   Begin MSComDlg.CommonDialog dlgOptions 
      Left            =   5400
      Top             =   2880
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'frmOptions - frmOptions.frm
'   Options/Properties Form...
'   Copyright © 1999-2002, Ken Clark
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Description:
'   08/20/02    Started History;
'=================================================================================================================================
Option Explicit
Dim strFileDSN As String
Dim strImagePath As String
Private Sub chkTraceMode_Click()
    cmdBrowseTraceFile.Enabled = (chkTraceMode.Value = vbChecked)
End Sub
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
    
    On Error Resume Next
    CurrentPath = ParsePath(strImagePath, DrvDirNoSlash)
    CurrentDrive = ParsePath(strImagePath, DrvOnly)
    CurrentImage = ParsePath(strImagePath, FileNameBaseExt)
    ChDrive CurrentDrive
    If Dir(CurrentPath, vbDirectory) <> vbNullString Then ChDir CurrentPath
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
Private Sub cmdBrowseTraceFile_Click()
    Dim CurrentPath As String
    Dim CurrentDrive As String
    Dim CurrentImage As String
    
    On Error Resume Next
    CurrentPath = ParsePath(txtTraceFile.Text, DrvDirNoSlash)
    CurrentDrive = ParsePath(txtTraceFile.Text, DrvOnly)
    CurrentImage = ParsePath(txtTraceFile.Text, FileNameBaseExt)
    ChDrive CurrentDrive
    If Dir(CurrentPath, vbDirectory) <> vbNullString Then ChDir CurrentPath
    With dlgOptions
        .DialogTitle = "Select New Trace File"
        .FileName = ParsePath(txtTraceFile.Text, FileNameBaseExt)
        .Filter = "Log Files|*.log|All Files (*.*)|*.*"
        .FilterIndex = 1
        .ShowOpen    ' Call the open file procedure.
        txtTraceFile.Text = .FileName
    End With
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
    gstrFileDSN = strFileDSN
    Call SaveRegistrySetting(HKEY_CURRENT_USER, "Software\KClark\" & App.FileDescription & "\Environment", "FileDSN", strFileDSN)
    gstrImagePath = strImagePath
    Call SaveRegistrySetting(HKEY_CURRENT_USER, "Software\KClark\" & App.FileDescription & "\Environment", "ImagePath", strImagePath)
    gfUseFilterMethod = (chkUseFilterMethod.Value = vbChecked)
    Call SaveRegistrySetting(HKEY_CURRENT_USER, "Software\KClark\" & App.FileDescription & "\Environment", "UseFilterMethod", gfUseFilterMethod)
    
    Call UpdateTraceFile
    Call SaveRegistrySetting(HKEY_CURRENT_USER, "Software\KClark\" & App.FileDescription & "\Environment", "TraceMode", gfTraceMode)
    Call SaveRegistrySetting(HKEY_CURRENT_USER, "Software\KClark\" & App.FileDescription & "\Environment", "TraceFile", gstrTraceFile)
    Unload Me
End Sub
Private Sub Form_Activate()
    If strFileDSN = vbNullString Then cmdBrowseDSN_Click
End Sub
Private Sub Form_Load()
    strFileDSN = GetRegistrySetting(HKEY_CURRENT_USER, "Software\KClark\" & App.FileDescription & "\Environment", "FileDSN", gstrFileDSN)
    txtDSN.Text = ParsePath(strFileDSN, FileNameBaseExt)
    strImagePath = GetRegistrySetting(HKEY_CURRENT_USER, "Software\KClark\" & App.FileDescription & "\Environment", "ImagePath", gstrImagePath)
    txtBackground.Text = ParsePath(strImagePath, FileNameBaseExt)
    If gfUseFilterMethod Then
        chkUseFilterMethod.Value = vbChecked
    Else
        chkUseFilterMethod.Value = vbUnchecked
    End If
    If gfTraceMode Then
        chkTraceMode.Value = vbChecked
    Else
        chkTraceMode.Value = vbUnchecked
    End If
    txtTraceFile.Text = gstrTraceFile
    cmdBrowseTraceFile.Enabled = (chkTraceMode.Value = vbChecked)
End Sub
Private Sub txtBackground_Validate(Cancel As Boolean)
    If Trim(txtBackground.Text) = vbNullString Then Cancel = True
End Sub
Private Sub txtDSN_Validate(Cancel As Boolean)
    If Trim(txtDSN.Text) = vbNullString Then Cancel = True
End Sub
Private Sub txtTraceFile_Validate(Cancel As Boolean)
    If Trim(txtTraceFile.Text) = vbNullString Then Cancel = True
End Sub
Private Sub UpdateTraceFile()
    If chkTraceMode.Value = vbUnchecked Then
        Call Trace(trcBody, "Trace File Closed.")
        Call Trace(trcBody, String(132, "="))
        gfTraceMode = False
    Else
        gfTraceMode = True
        If gstrTraceFile <> txtTraceFile.Text Then
            Call Trace(trcBody, "Trace File Closed.")
            Call Trace(trcBody, String(132, "="))
        End If
        gstrTraceFile = txtTraceFile.Text
        Call Trace(trcBody, String(132, "="))
        Call Trace(trcBody, "Trace File Opened - " & gstrTraceFile)
    End If
End Sub
