VERSION 5.00
Object = "{AFB5A97B-6C4F-11D2-BDFF-98840BC10000}#7.0#0"; "WEBLINKS.OCX"
Begin VB.Form frmWebLinks 
   Caption         =   "Web Shortcuts"
   ClientHeight    =   4716
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   7896
   Icon            =   "frmWebLinks.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4716
   ScaleWidth      =   7896
   StartUpPosition =   1  'CenterOwner
   Begin WebLinks.kfcWebLinks wlWebLinks 
      Height          =   4032
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   7752
      _ExtentX        =   13674
      _ExtentY        =   7112
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

