VERSION 5.00
Begin VB.Form frmOptions 
   Caption         =   "Options"
   ClientHeight    =   1404
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   ScaleHeight     =   1404
   ScaleWidth      =   5820
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   372
      Left            =   2964
      TabIndex        =   7
      Top             =   900
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   1884
      TabIndex        =   6
      Top             =   900
      Width           =   972
   End
   Begin VB.CommandButton cmdBrowseDSN 
      Caption         =   "&Browse"
      Height          =   288
      Left            =   4020
      TabIndex        =   5
      Top             =   480
      Width           =   1572
   End
   Begin VB.TextBox txtDSN 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   288
      Left            =   1752
      TabIndex        =   3
      Top             =   480
      Width           =   2172
   End
   Begin VB.CommandButton cmdBrowseImages 
      Caption         =   "&Select New Image"
      Height          =   288
      Left            =   4020
      TabIndex        =   2
      Top             =   120
      Width           =   1572
   End
   Begin VB.TextBox txtBackground 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   288
      Left            =   1740
      TabIndex        =   0
      Top             =   120
      Width           =   2172
   End
   Begin VB.Label lblDSN 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Data Source Name:"
      Height          =   192
      Left            =   228
      TabIndex        =   4
      Top             =   528
      Width           =   1416
   End
   Begin VB.Label lblBackground 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Background Image:"
      Height          =   192
      Left            =   228
      TabIndex        =   1
      Top             =   168
      Width           =   1404
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

