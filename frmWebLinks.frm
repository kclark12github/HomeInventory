VERSION 5.00
Begin VB.Form frmWebLinks 
   Caption         =   "Web Shortcuts"
   ClientHeight    =   4152
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   7896
   Icon            =   "frmWebLinks.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4152
   ScaleWidth      =   7896
   StartUpPosition =   1  'CenterOwner
   Begin HomeInventory.kfcWebLinks wlWebLinks 
      Height          =   4152
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7692
      _ExtentX        =   13568
      _ExtentY        =   7324
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
    If Not fActivated Then
        fActivated = True
        Me.MousePointer = vbHourglass
        wlWebLinks.PopulateMenu
        Me.MousePointer = vbNormal
    End If
End Sub
Private Sub Form_Load()
    fActivated = False
End Sub
Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        wlWebLinks.Width = Me.ScaleWidth
        wlWebLinks.Height = Me.ScaleHeight
        wlWebLinks.Move 0, 0
    End If
End Sub

