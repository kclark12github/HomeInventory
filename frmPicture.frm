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
Const iMinWidth = 2184
Const iMinHeight = 1440
Dim fActivated As Boolean
Private Sub LoadImage()
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
    
    picImage.Picture = frmImages.picImage.Picture
    picImage.Move 0, 0 ', picImage.Picture.Width, picImage.Picture.Height
    
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
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub
Private Sub Form_Activate()
    If fActivated Then Exit Sub
    fActivated = True
End Sub
Private Sub Form_Load()
    fActivated = False
    Me.Caption = frmImages.rsMain("Name")
    LoadImage
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
Private Sub scrollH_Change()
    picImage.Left = -scrollH.Value
End Sub
Private Sub scrollV_Change()
    picImage.Top = -scrollV.Value
End Sub

