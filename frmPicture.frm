VERSION 5.00
Begin VB.Form frmPicture 
   Caption         =   "Picture"
   ClientHeight    =   4068
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   5964
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4068
   ScaleWidth      =   5964
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar scrollH 
      Height          =   192
      LargeChange     =   1000
      Left            =   0
      SmallChange     =   100
      TabIndex        =   3
      Top             =   3780
      Width           =   5712
   End
   Begin VB.VScrollBar scrollV 
      Height          =   3792
      LargeChange     =   1000
      Left            =   5700
      SmallChange     =   100
      TabIndex        =   2
      Top             =   0
      Width           =   192
   End
   Begin VB.PictureBox picWindow 
      Height          =   4032
      Left            =   0
      ScaleHeight     =   3984
      ScaleWidth      =   5904
      TabIndex        =   0
      Top             =   0
      Width           =   5952
      Begin VB.PictureBox picImage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         ForeColor       =   &H80000008&
         Height          =   3804
         Left            =   0
         ScaleHeight     =   3780
         ScaleWidth      =   5664
         TabIndex        =   1
         Top             =   0
         Width           =   5688
      End
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
End Sub
Private Sub Form_Load()
    fActivated = False
    Me.Caption = frmImages.rsMain("Name")
    LoadImage
End Sub
Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        If scrollH.Visible Then
            scrollH.Top = Me.ScaleHeight - scrollH.Height
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
            scrollV.Height = Me.ScaleHeight
            If scrollH.Visible Then scrollV.Height = scrollV.Height - scrollH.Height
            scrollV.Max = picImage.Height - Me.ScaleHeight
            scrollV.SmallChange = picImage.Height / 100
            scrollV.LargeChange = picImage.Height / 20
            picWindow.Width = scrollV.Left
        End If
    End If
End Sub
Private Sub scrollH_Change()
    picImage.Left = -scrollH.Value
End Sub
Private Sub scrollV_Change()
    picImage.Top = -scrollV.Value
End Sub

