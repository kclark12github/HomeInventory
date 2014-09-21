VERSION 5.00
Begin VB.Form frmImage 
   Caption         =   "Image Display"
   ClientHeight    =   6216
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   8256
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6216
   ScaleWidth      =   8256
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
Attribute VB_Name = "frmImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoConn As ADODB.Connection
Dim WithEvents rsMain As ADODB.Recordset
Attribute rsMain.VB_VarHelpID = -1
Dim mode As ActionMode
Dim fTransaction As Boolean
Dim DBinfo As DataBaseInfo
Private Sub Form_Load()
    Set adoConn = New ADODB.Connection
    Set rsMain = New ADODB.Recordset
    Set DBinfo = frmMain.DBcollection("US Navy Ships")
    With DBinfo
        adoConn.Provider = .Provider
        adoConn.CommandTimeout = 60
        adoConn.ConnectionTimeout = 60
        adoConn.Open .PathName, .UserName, .Password
    End With
    rsMain.CursorLocation = adUseClient
    rsMain.Open "select * from [Images]", adoConn, adOpenKeyset, adLockBatchOptimistic
    frmMain.BindField picImage, "Image", rsMain
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
Private Sub rsMain_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Dim Caption As String
    Dim i As Long
    
    'ReDim pic(rsMain("Image").ActualSize)
    
    'On Error GoTo ErrorHandler
    'For i = 1 To rsMain("Image").ActualSize
    '    pic(i) = rsMain("Image").GetChunk(1)
    'Next
    'picImage.PaintPicture pic, 0, 0
    Exit Sub

ErrorHandler:
    MsgBox Err.Description & " (Error " & Err.Number & ")", vbExclamation, Me.Caption
    Resume Next
End Sub
Private Sub scrollH_Change()
    picImage.Left = -scrollH.Value
End Sub
Private Sub scrollV_Change()
    picImage.Top = -scrollV.Value
End Sub

