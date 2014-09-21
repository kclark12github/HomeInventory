VERSION 5.00
Object = "{AFB5A97B-6C4F-11D2-BDFF-98840BC10000}#5.0#0"; "WEBLINKS.OCX"
Begin VB.Form frmWebLinks 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Web Menu Control"
   ClientHeight    =   4716
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   7896
   Icon            =   "frmWebLinks.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4716
   ScaleWidth      =   7896
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
    Unload Me
End Sub
Private Sub Form_Load()
    kfcWebLinks1.PopulateMenu
End Sub
