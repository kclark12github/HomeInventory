VERSION 5.00
Begin VB.Form frmWebLinks 
   Caption         =   "Web Shortcuts"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   7905
   Icon            =   "frmWebLinks.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin HomeInventory.kfcWebLinks wlWebLinks 
      Height          =   4152
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7692
      _ExtentX        =   13573
      _ExtentY        =   7329
   End
End
Attribute VB_Name = "frmWebLinks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fActivated As Boolean
Private Sub Form_Activate()
    Me.Top = frmMain.saveTop + ((frmMain.Height - Me.Height) / 2)
    Me.Left = frmMain.saveLeft + ((frmMain.Width - Me.Width) / 2)
    DoEvents
    If Not fActivated Then
        fActivated = True
        Me.MousePointer = vbHourglass
        wlWebLinks.PopulateMenu
        Me.MousePointer = vbNormal
    End If
    DoEvents
End Sub
Private Sub Form_Load()
    fActivated = False
    Me.Move 0, 0, frmMain.Width, frmMain.Height
End Sub
Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        wlWebLinks.Width = Me.ScaleWidth
        wlWebLinks.Height = Me.ScaleHeight
        wlWebLinks.Move 0, 0
    End If
End Sub

