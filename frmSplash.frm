VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1872
   ClientLeft      =   48
   ClientTop       =   48
   ClientWidth     =   4848
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1872
   ScaleWidth      =   4848
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraSplash 
      Height          =   1695
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4755
      Begin VB.Label lblActivity 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   396
         Left            =   108
         TabIndex        =   5
         Top             =   660
         Width           =   4560
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblSplash 
         BackStyle       =   0  'Transparent
         Caption         =   "Loading..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   432
         Left            =   60
         TabIndex        =   4
         Top             =   180
         Width           =   3912
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      CausesValidation=   0   'False
      Height          =   2928
      Left            =   420
      Picture         =   "frmActivitySplash.frx":0000
      ScaleHeight     =   2880
      ScaleWidth      =   3840
      TabIndex        =   1
      Top             =   1380
      Visible         =   0   'False
      Width           =   3888
      Begin VB.Label lblPicSplash 
         BackStyle       =   0  'Transparent
         Caption         =   "Loading..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   432
         Left            =   840
         TabIndex        =   3
         Top             =   1260
         Width           =   3912
      End
      Begin VB.Label lblPicActivity 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   945
         TabIndex        =   2
         Top             =   1920
         Width           =   2880
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmActivitySplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'frmSplash - frmSplash.frm
'   Splash Screen...
'   Copyright © 2001, SunGard Investor Accounting Systems
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Problem:    Programmer:     Description:
'   03/16/00    None        Ken Clark       "Deactivated" the test pattern image;
'   02/06/00    None        Ken Clark       Adjusted label positions for different screen resolutions;
'   01/13/00    None        Ken Clark       Centered labels;
'   01/06/00    None        JAD             Made the descriptive labels visible again
'   01/06/00    None        Ken Clark       Resized form to cover entire Picture1 regardless of screen resolution;
'   11/09/99    None        JAD             Added MousePointer in Form_Load
'   11/04/99    None        Ken Clark       Created;
'=================================================================================================================================
Option Explicit
Dim savOperatorSounds As Boolean
Private Sub Form_Activate()
    'savOperatorSounds = operatorSounds
    'operatorSounds = True
    'Call DoSounds("tp2.wav", False, False)
End Sub
Private Sub Form_Deactivate()
    'frmMain.mmcSound.Wait = True
    'frmMain.mmcSound.Command = "Stop"
    'operatorSounds = savOperatorSounds
End Sub
Private Sub Form_Load()
    Dim wPad As Single
    Dim hPad As Single
    
    'wPad = Me.Width - Me.ScaleWidth
    'hPad = Me.Height - Me.ScaleHeight
    
    'Me.Width = Picture1.Width + wPad
    'Me.Height = Picture1.Height + hPad
    
    ''Sizing based solely on the "Please Stand By..." image...
    'lblPicSplash.Move ((Me.ScaleWidth - lblPicSplash.Width) / 2) + (Me.ScaleWidth / 19), (Me.ScaleHeight / 2) - lblPicSplash.Height - (Me.ScaleHeight / 30)
    'lblPicActivity.Width = Me.ScaleWidth * (12 / 19)
    'lblPicActivity.Move (Me.ScaleWidth - lblPicActivity.Width) / 2, (Me.ScaleHeight / 2) + (Me.ScaleHeight / 30)
    
    MousePointer = vbHourglass
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmMain.timMain.Enabled = False
End Sub
