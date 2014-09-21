VERSION 5.00
Object = "{AFB5A97B-6C4F-11D2-BDFF-98840BC10000}#1.0#0"; "WEBLINKS.OCX"
Begin VB.Form frmWebLinks 
   Caption         =   "Web Menu Control"
   ClientHeight    =   4716
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   7896
   Icon            =   "frmWebLinks.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4716
   ScaleWidth      =   7896
   StartUpPosition =   3  'Windows Default
   Begin WebLinks.kfcWebLinks kfcWebLinks1 
      Height          =   4032
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   7752
      _ExtentX        =   13674
      _ExtentY        =   7112
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   372
      Left            =   3300
      TabIndex        =   0
      Top             =   4200
      Width           =   1272
   End
End
Attribute VB_Name = "frmWebLinks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    End
End Sub

Private Sub Form_Load()
    Me.Show
    kfcWebLinks1.PopulateMenu
End Sub
