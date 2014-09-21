VERSION 5.00
Begin VB.Form frmPicture 
   Caption         =   "Picture"
   ClientHeight    =   5004
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   6636
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5004
   ScaleWidth      =   6636
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picImage 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3324
      Left            =   240
      ScaleHeight     =   3276
      ScaleWidth      =   4560
      TabIndex        =   2
      Top             =   300
      Width           =   4608
   End
   Begin VB.VScrollBar scrollV 
      Height          =   3792
      LargeChange     =   1000
      Left            =   5640
      SmallChange     =   100
      TabIndex        =   1
      Top             =   0
      Width           =   192
   End
   Begin VB.HScrollBar scrollH 
      Height          =   192
      LargeChange     =   1000
      Left            =   0
      SmallChange     =   100
      TabIndex        =   0
      Top             =   3780
      Width           =   5652
   End
End
Attribute VB_Name = "frmPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strPictureFile As String
Private Sub Form_Activate()
    picImage.Picture = LoadPicture(strPictureFile)
End Sub
Private Sub Form_Resize()
    picImage.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    If Me.WindowState <> vbMinimized Then
        If scrollH.Visible Then
            scrollH.Top = Me.ScaleHeight - scrollH.Height
            scrollH.Left = 0
            scrollH.Width = Me.ScaleWidth - scrollV.Width
            scrollH.Max = picImage.Width - Me.ScaleWidth
            scrollH.SmallChange = picImage.Width / 1000
            scrollH.LargeChange = picImage.Width / 50
        End If
        
        If scrollV.Visible Then
            scrollV.Top = 0
            scrollV.Left = Me.ScaleWidth - scrollV.Width
            scrollV.Height = Me.ScaleHeight - scrollH.Height
            scrollV.Max = picImage.Height - Me.ScaleHeight
            scrollV.SmallChange = picImage.Height / 1000
            scrollV.LargeChange = picImage.Height / 50
        End If
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Kill strPictureFile
End Sub
Private Sub scrollH_Change()
    picImage.Left = -scrollH.Value
End Sub
Private Sub scrollV_Change()
    picImage.Top = -scrollV.Value
End Sub

