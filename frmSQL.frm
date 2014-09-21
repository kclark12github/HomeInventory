VERSION 5.00
Begin VB.Form frmSQL 
   Caption         =   "SQL"
   ClientHeight    =   5256
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   7596
   LinkTopic       =   "Form1"
   ScaleHeight     =   5256
   ScaleWidth      =   7596
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   372
      Left            =   6600
      TabIndex        =   5
      Top             =   480
      Width           =   972
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "&Execute"
      Default         =   -1  'True
      Height          =   372
      Left            =   6600
      TabIndex        =   4
      Top             =   60
      Width           =   972
   End
   Begin VB.Frame frameResults 
      Caption         =   "Results"
      Height          =   1092
      Left            =   120
      TabIndex        =   2
      Top             =   1380
      Width           =   3492
      Begin VB.TextBox txtResults 
         Height          =   852
         Left            =   60
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   180
         Width           =   3372
      End
   End
   Begin VB.Frame frameSQL 
      Caption         =   "SQL Statement"
      Height          =   1092
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3492
      Begin VB.TextBox txtSQL 
         Height          =   852
         Left            =   60
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   3372
      End
   End
End
Attribute VB_Name = "frmSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const MarginTwips As Integer = 60
Private Sub Form_Load()
    
End Sub
Private Sub Form_Resize()
    cmdExecute.Move Me.ScaleWidth - cmdExecute.Width - MarginTwips, MarginTwips
    cmdClose.Move Me.ScaleWidth - cmdClose.Width - MarginTwips, MarginTwips
    frameSQL.Move MarginTwips, MarginTwips, Me.ScaleWidth - cmdExecute.Width - (3 * MarginTwips), Me.ScaleHeight / 3
    frameResults.Move MarginTwips, frameSQL.Top + MarginTwips, frameSQL.Width, Me.ScaleHeight - frameSQL.Height - (3 * MarginTwips)
End Sub
