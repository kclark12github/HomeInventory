VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search"
   ClientHeight    =   2550
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Default         =   -1  'True
      Height          =   372
      Left            =   4560
      TabIndex        =   33
      Top             =   0
      Width           =   972
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   372
      Left            =   4560
      TabIndex        =   32
      Top             =   420
      Width           =   972
   End
   Begin MSDataListLib.DataCombo dbcFields 
      Height          =   288
      Index           =   0
      Left            =   360
      TabIndex        =   66
      Top             =   120
      Width           =   4212
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dbcFields 
      Height          =   288
      Index           =   1
      Left            =   360
      TabIndex        =   67
      Top             =   120
      Width           =   4212
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dbcFields 
      Height          =   288
      Index           =   2
      Left            =   360
      TabIndex        =   68
      Top             =   120
      Width           =   4212
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dbcFields 
      Height          =   288
      Index           =   3
      Left            =   360
      TabIndex        =   69
      Top             =   120
      Width           =   4212
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dbcFields 
      Height          =   288
      Index           =   4
      Left            =   360
      TabIndex        =   70
      Top             =   120
      Width           =   4212
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dbcFields 
      Height          =   288
      Index           =   5
      Left            =   360
      TabIndex        =   71
      Top             =   120
      Width           =   4212
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dbcFields 
      Height          =   288
      Index           =   6
      Left            =   360
      TabIndex        =   72
      Top             =   120
      Width           =   4212
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dbcFields 
      Height          =   288
      Index           =   7
      Left            =   360
      TabIndex        =   73
      Top             =   120
      Width           =   4212
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dbcFields 
      Height          =   288
      Index           =   8
      Left            =   360
      TabIndex        =   74
      Top             =   120
      Width           =   4212
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dbcFields 
      Height          =   288
      Index           =   9
      Left            =   360
      TabIndex        =   75
      Top             =   120
      Width           =   4212
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dbcFields 
      Height          =   288
      Index           =   10
      Left            =   360
      TabIndex        =   76
      Top             =   120
      Width           =   4212
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dbcFields 
      Height          =   288
      Index           =   11
      Left            =   360
      TabIndex        =   77
      Top             =   120
      Width           =   4212
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dbcFields 
      Height          =   288
      Index           =   12
      Left            =   360
      TabIndex        =   78
      Top             =   120
      Width           =   4212
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dbcFields 
      Height          =   288
      Index           =   13
      Left            =   360
      TabIndex        =   79
      Top             =   120
      Width           =   4212
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dbcFields 
      Height          =   288
      Index           =   14
      Left            =   360
      TabIndex        =   80
      Top             =   120
      Width           =   4212
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dbcFields 
      Height          =   288
      Index           =   15
      Left            =   360
      TabIndex        =   81
      Top             =   120
      Width           =   4212
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dbcFields 
      Height          =   288
      Index           =   16
      Left            =   360
      TabIndex        =   82
      Top             =   120
      Width           =   4212
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dbcFields 
      Height          =   288
      Index           =   17
      Left            =   360
      TabIndex        =   83
      Top             =   120
      Width           =   4212
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dbcFields 
      Height          =   288
      Index           =   18
      Left            =   360
      TabIndex        =   84
      Top             =   120
      Width           =   4212
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dbcFields 
      Height          =   288
      Index           =   19
      Left            =   360
      TabIndex        =   85
      Top             =   120
      Width           =   4212
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dbcFields 
      Height          =   288
      Index           =   20
      Left            =   360
      TabIndex        =   86
      Top             =   120
      Width           =   4212
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dbcFields 
      Height          =   288
      Index           =   21
      Left            =   360
      TabIndex        =   87
      Top             =   120
      Width           =   4212
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dbcFields 
      Height          =   288
      Index           =   22
      Left            =   360
      TabIndex        =   88
      Top             =   120
      Width           =   4212
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dbcFields 
      Height          =   288
      Index           =   23
      Left            =   360
      TabIndex        =   89
      Top             =   120
      Width           =   4212
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dbcFields 
      Height          =   288
      Index           =   24
      Left            =   360
      TabIndex        =   90
      Top             =   120
      Width           =   4212
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dbcFields 
      Height          =   288
      Index           =   25
      Left            =   360
      TabIndex        =   91
      Top             =   120
      Width           =   4212
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dbcFields 
      Height          =   288
      Index           =   26
      Left            =   360
      TabIndex        =   92
      Top             =   120
      Width           =   4212
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dbcFields 
      Height          =   288
      Index           =   27
      Left            =   360
      TabIndex        =   93
      Top             =   120
      Width           =   4212
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dbcFields 
      Height          =   288
      Index           =   28
      Left            =   360
      TabIndex        =   94
      Top             =   120
      Width           =   4212
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dbcFields 
      Height          =   288
      Index           =   29
      Left            =   360
      TabIndex        =   95
      Top             =   120
      Width           =   4212
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dbcFields 
      Height          =   288
      Index           =   30
      Left            =   360
      TabIndex        =   96
      Top             =   120
      Width           =   4212
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dbcFields 
      Height          =   288
      Index           =   31
      Left            =   360
      TabIndex        =   97
      Top             =   120
      Width           =   4212
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   ""
   End
   Begin VB.TextBox txtFields 
      Height          =   312
      Index           =   31
      Left            =   180
      TabIndex        =   31
      Top             =   60
      Width           =   1872
   End
   Begin VB.TextBox txtFields 
      Height          =   312
      Index           =   30
      Left            =   180
      TabIndex        =   30
      Top             =   60
      Width           =   1872
   End
   Begin VB.TextBox txtFields 
      Height          =   312
      Index           =   29
      Left            =   180
      TabIndex        =   29
      Top             =   60
      Width           =   1872
   End
   Begin VB.TextBox txtFields 
      Height          =   312
      Index           =   28
      Left            =   180
      TabIndex        =   28
      Top             =   60
      Width           =   1872
   End
   Begin VB.TextBox txtFields 
      Height          =   312
      Index           =   27
      Left            =   180
      TabIndex        =   27
      Top             =   60
      Width           =   1872
   End
   Begin VB.TextBox txtFields 
      Height          =   312
      Index           =   26
      Left            =   180
      TabIndex        =   26
      Top             =   60
      Width           =   1872
   End
   Begin VB.TextBox txtFields 
      Height          =   312
      Index           =   25
      Left            =   180
      TabIndex        =   25
      Top             =   60
      Width           =   1872
   End
   Begin VB.TextBox txtFields 
      Height          =   312
      Index           =   24
      Left            =   180
      TabIndex        =   24
      Top             =   60
      Width           =   1872
   End
   Begin VB.TextBox txtFields 
      Height          =   312
      Index           =   23
      Left            =   180
      TabIndex        =   23
      Top             =   60
      Width           =   1872
   End
   Begin VB.TextBox txtFields 
      Height          =   312
      Index           =   22
      Left            =   180
      TabIndex        =   22
      Top             =   60
      Width           =   1872
   End
   Begin VB.TextBox txtFields 
      Height          =   312
      Index           =   21
      Left            =   180
      TabIndex        =   21
      Top             =   60
      Width           =   1872
   End
   Begin VB.TextBox txtFields 
      Height          =   312
      Index           =   20
      Left            =   180
      TabIndex        =   20
      Top             =   60
      Width           =   1872
   End
   Begin VB.TextBox txtFields 
      Height          =   312
      Index           =   19
      Left            =   180
      TabIndex        =   19
      Top             =   60
      Width           =   1872
   End
   Begin VB.TextBox txtFields 
      Height          =   312
      Index           =   18
      Left            =   180
      TabIndex        =   18
      Top             =   60
      Width           =   1872
   End
   Begin VB.TextBox txtFields 
      Height          =   312
      Index           =   17
      Left            =   180
      TabIndex        =   17
      Top             =   60
      Width           =   1872
   End
   Begin VB.TextBox txtFields 
      Height          =   312
      Index           =   16
      Left            =   180
      TabIndex        =   16
      Top             =   60
      Width           =   1872
   End
   Begin VB.TextBox txtFields 
      Height          =   312
      Index           =   15
      Left            =   180
      TabIndex        =   15
      Top             =   60
      Width           =   1872
   End
   Begin VB.TextBox txtFields 
      Height          =   312
      Index           =   14
      Left            =   180
      TabIndex        =   14
      Top             =   60
      Width           =   1872
   End
   Begin VB.TextBox txtFields 
      Height          =   312
      Index           =   13
      Left            =   180
      TabIndex        =   13
      Top             =   60
      Width           =   1872
   End
   Begin VB.TextBox txtFields 
      Height          =   312
      Index           =   12
      Left            =   180
      TabIndex        =   12
      Top             =   60
      Width           =   1872
   End
   Begin VB.TextBox txtFields 
      Height          =   312
      Index           =   11
      Left            =   180
      TabIndex        =   11
      Top             =   60
      Width           =   1872
   End
   Begin VB.TextBox txtFields 
      Height          =   312
      Index           =   10
      Left            =   180
      TabIndex        =   10
      Top             =   60
      Width           =   1872
   End
   Begin VB.TextBox txtFields 
      Height          =   312
      Index           =   9
      Left            =   180
      TabIndex        =   9
      Top             =   60
      Width           =   1872
   End
   Begin VB.TextBox txtFields 
      Height          =   312
      Index           =   8
      Left            =   180
      TabIndex        =   8
      Top             =   60
      Width           =   1872
   End
   Begin VB.TextBox txtFields 
      Height          =   312
      Index           =   7
      Left            =   180
      TabIndex        =   7
      Top             =   60
      Width           =   1872
   End
   Begin VB.TextBox txtFields 
      Height          =   312
      Index           =   6
      Left            =   180
      TabIndex        =   6
      Top             =   60
      Width           =   1872
   End
   Begin VB.TextBox txtFields 
      Height          =   312
      Index           =   5
      Left            =   180
      TabIndex        =   5
      Top             =   60
      Width           =   1872
   End
   Begin VB.TextBox txtFields 
      Height          =   312
      Index           =   4
      Left            =   180
      TabIndex        =   4
      Top             =   60
      Width           =   1872
   End
   Begin VB.TextBox txtFields 
      Height          =   312
      Index           =   3
      Left            =   180
      TabIndex        =   3
      Top             =   60
      Width           =   1872
   End
   Begin VB.TextBox txtFields 
      Height          =   312
      Index           =   2
      Left            =   180
      TabIndex        =   2
      Top             =   60
      Width           =   1872
   End
   Begin VB.TextBox txtFields 
      Height          =   312
      Index           =   1
      Left            =   180
      TabIndex        =   1
      Top             =   60
      Width           =   1872
   End
   Begin VB.TextBox txtFields 
      Height          =   312
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   60
      Width           =   4212
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      Height          =   192
      Index           =   31
      Left            =   0
      TabIndex        =   65
      Top             =   120
      Width           =   36
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      Height          =   192
      Index           =   30
      Left            =   0
      TabIndex        =   64
      Top             =   120
      Width           =   36
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      Height          =   192
      Index           =   29
      Left            =   0
      TabIndex        =   63
      Top             =   120
      Width           =   36
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      Height          =   192
      Index           =   28
      Left            =   0
      TabIndex        =   62
      Top             =   120
      Width           =   36
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      Height          =   192
      Index           =   27
      Left            =   0
      TabIndex        =   61
      Top             =   120
      Width           =   36
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      Height          =   192
      Index           =   26
      Left            =   0
      TabIndex        =   60
      Top             =   120
      Width           =   36
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      Height          =   192
      Index           =   25
      Left            =   0
      TabIndex        =   59
      Top             =   120
      Width           =   36
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      Height          =   192
      Index           =   24
      Left            =   0
      TabIndex        =   58
      Top             =   120
      Width           =   36
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      Height          =   192
      Index           =   23
      Left            =   0
      TabIndex        =   57
      Top             =   120
      Width           =   36
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      Height          =   192
      Index           =   22
      Left            =   0
      TabIndex        =   56
      Top             =   120
      Width           =   36
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      Height          =   192
      Index           =   21
      Left            =   0
      TabIndex        =   55
      Top             =   120
      Width           =   36
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      Height          =   192
      Index           =   20
      Left            =   0
      TabIndex        =   54
      Top             =   120
      Width           =   36
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      Height          =   192
      Index           =   19
      Left            =   0
      TabIndex        =   53
      Top             =   120
      Width           =   36
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      Height          =   192
      Index           =   18
      Left            =   0
      TabIndex        =   52
      Top             =   120
      Width           =   36
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      Height          =   192
      Index           =   17
      Left            =   0
      TabIndex        =   51
      Top             =   120
      Width           =   36
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      Height          =   192
      Index           =   16
      Left            =   0
      TabIndex        =   50
      Top             =   120
      Width           =   36
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      Height          =   192
      Index           =   15
      Left            =   0
      TabIndex        =   49
      Top             =   120
      Width           =   36
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      Height          =   192
      Index           =   14
      Left            =   0
      TabIndex        =   48
      Top             =   120
      Width           =   36
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      Height          =   192
      Index           =   13
      Left            =   0
      TabIndex        =   47
      Top             =   120
      Width           =   36
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      Height          =   192
      Index           =   12
      Left            =   0
      TabIndex        =   46
      Top             =   120
      Width           =   36
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      Height          =   192
      Index           =   11
      Left            =   0
      TabIndex        =   45
      Top             =   120
      Width           =   36
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      Height          =   192
      Index           =   10
      Left            =   0
      TabIndex        =   44
      Top             =   120
      Width           =   36
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      Height          =   192
      Index           =   9
      Left            =   0
      TabIndex        =   43
      Top             =   120
      Width           =   36
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      Height          =   192
      Index           =   8
      Left            =   0
      TabIndex        =   42
      Top             =   120
      Width           =   36
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      Height          =   192
      Index           =   7
      Left            =   0
      TabIndex        =   41
      Top             =   120
      Width           =   36
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      Height          =   192
      Index           =   6
      Left            =   0
      TabIndex        =   40
      Top             =   120
      Width           =   36
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      Height          =   192
      Index           =   5
      Left            =   0
      TabIndex        =   39
      Top             =   120
      Width           =   36
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      Height          =   192
      Index           =   4
      Left            =   0
      TabIndex        =   38
      Top             =   120
      Width           =   36
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      Height          =   192
      Index           =   3
      Left            =   0
      TabIndex        =   37
      Top             =   120
      Width           =   36
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      Height          =   192
      Index           =   2
      Left            =   0
      TabIndex        =   36
      Top             =   120
      Width           =   36
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      Height          =   192
      Index           =   1
      Left            =   0
      TabIndex        =   35
      Top             =   120
      Width           =   36
   End
   Begin VB.Label lblFields 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Height          =   192
      Index           =   0
      Left            =   0
      TabIndex        =   34
      Top             =   120
      Width           =   36
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public RS As ADODB.Recordset
Public strSearch As String
Const vSpace As Integer = 60
Const hSpace As Integer = 60
Const StartTop As Integer = 60
Const StartLeft As Integer = 240
Private Sub cmdApply_Click()
    Dim i As Integer
    Dim ctl As Control
    'Parse fields on screen and build a Search for the recordset...
    
    strSearch = vbNullString
    For i = 0 To RS.Fields.Count - 1
        If txtFields(i).Enabled Then
            Set ctl = txtFields(i)
            If Trim(ctl.Text) <> vbNullString Then strSearch = strSearch & RS.Fields(i).Name & " " & ctl.Text & " And "
        Else
            Set ctl = dbcFields(i)
            If Trim(ctl.Text) <> vbNullString Then strSearch = strSearch & RS.Fields(i).Name & "='" & ctl.Text & "' And "
        End If
    Next i
    If Len(strSearch) > 0 Then strSearch = Left(strSearch, Len(strSearch) - 5)  'Get rid of the final " and "...
    On Error Resume Next
    RS.Find strSearch
    If Err.Number > 0 Then
        MsgBox "Invalid search criteria specified:" & vbCr & strSearch, vbExclamation
        Err.Clear
    Else
        Unload Me
    End If
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub dbcFields_GotFocus(Index As Integer)
    TextSelected
End Sub
Private Sub dbcFields_Validate(Index As Integer, Cancel As Boolean)
    dbcFields(Index).Text = Trim(dbcFields(Index).Text)
End Sub
Private Sub Form_Activate()
    Dim Caption As String
    Dim ctl As Control
    Dim frm As Form
    Dim fUseDataCombo As Boolean
    Dim i As Integer
    Dim iField As Integer
    Dim iTop As Integer
    Dim iLeft As Integer
    Dim NewHeight As Integer
    Dim NewTop As Integer
    Dim NewWidth As Integer
    Dim NewLeft As Integer
    
    Call Trace(trcEnter, Me.Name & ".Form_Activate")
    Set frm = Forms(Forms.Count - 2)
    Me.Icon = frm.Icon
    NewWidth = 5832
    NewLeft = Me.Left + (Me.Width - NewWidth) / 2
    Me.Left = NewLeft
    Me.Width = NewWidth
    
    iTop = StartTop
    iLeft = StartLeft
    
    cmdApply.Top = StartTop
    cmdApply.Left = Me.ScaleWidth - cmdApply.Width - hSpace
    cmdCancel.Top = cmdApply.Top + vSpace + cmdApply.Height
    cmdCancel.Left = cmdApply.Left
    
    iField = 0
    For i = 0 To RS.Fields.Count - 1
        fUseDataCombo = False
        Caption = vbNullString
        Call Trace(trcBody, "RS.Fields(i).Name: " & RS.Fields(i).Name)
        Select Case RS.Fields(i).Type
            Case adBinary, adLongVarBinary, adLongVarChar
                'Don't deal with filtering these fields (we could do
                'Memo fields, but eliminating them reduces the size of
                'the filter form for these activites, so let's not)...
                Call Trace(trcBody, vbTab & "Skipping: RS.Fields(i).Name:" & RS.Fields(i).Name)
            Case Else
                If i > 31 Then
                    MsgBox "Warning: Only the first 32 fields can be used to filter your data.", vbInformation
                    Exit For
                Else
                    For Each ctl In frm.Controls
                        If ctl.Tag = vbNullString Then GoTo SkipControl
                        Select Case TypeName(ctl)
                            Case "CheckBox", "DataCombo", "Label", "PictureBox", "RichTextBox", "TextBox"
                                If UCase(ctl.DataField) = UCase(RS.Fields(i).Name) Then
                                    Caption = ctl.Tag
                                    Call Trace(trcBody, vbTab & "Caption: " & Caption)
                                    If TypeName(ctl) = "DataCombo" Then
                                        fUseDataCombo = True
                                        Call Trace(trcBody, vbTab & "BoundColumn: " & ctl.BoundColumn & "; ListField: " & ctl.ListField & "; DataField: " & ctl.DataField)
                                        Set dbcFields(i).DataSource = Nothing
                                        dbcFields(i).DataField = ctl.DataField
                                        Set dbcFields(i).DataSource = ctl.DataSource
                                        Set dbcFields(i).RowSource = Nothing
                                        dbcFields(i).BoundColumn = ctl.BoundColumn
                                        dbcFields(i).ListField = ctl.ListField
                                        Set dbcFields(i).RowSource = ctl.RowSource
                                        
                                        dbcFields(i).BoundText = vbNullString
                                        dbcFields(i).Text = vbNullString
                                    End If
                                    Exit For
                                End If
                        End Select
SkipControl:
                    Next ctl
                    If Caption = vbNullString Then GoTo SkipField
                    If fUseDataCombo Then
                        Set ctl = dbcFields(i)
                    Else
                        Set ctl = txtFields(i)
                    End If
                    
                    ctl.Visible = True
                    ctl.Enabled = True
                    ctl.Top = iTop
                    
                    lblFields(i).Alignment = lblFields(0).Alignment 'Right
                    lblFields(i).Visible = True
                    'lblFields(i).Caption = RS.Fields(i).Name & ":"
                    lblFields(i).Caption = Caption & ":"
                    lblFields(i).Top = ctl.Top + (ctl.Height / 2) - (lblFields(i).Height / 2)
                    lblFields(i).Left = iLeft
                    
                    ctl.Left = lblFields(i).Left + lblFields(i).Width + hSpace
                    ctl.Width = Me.ScaleWidth - lblFields(i).Left - lblFields(i).Width - hSpace - hSpace - cmdApply.Width - hSpace
                    ctl.TabIndex = i
                    
                    iTop = iTop + ctl.Height + vSpace
                    iField = i
                    Set ctl = Nothing
                End If
        End Select
SkipField:
    Next i
    
    If txtFields(iField).Enabled Then Set ctl = txtFields(iField) Else Set ctl = dbcFields(iField)
    NewHeight = (ctl.Top + txtFields(iField).Height + vSpace) + Me.Height - Me.ScaleHeight
    If NewHeight > Me.Height Then
        Me.Top = Me.Top + ((NewHeight - Me.Height) / 2)
    Else
        Me.Top = Me.Top - ((NewHeight - Me.Height) / 2)
    End If
    Me.Height = NewHeight
    If txtFields(0).Enabled Then Set ctl = txtFields(0) Else Set ctl = dbcFields(0)
    ctl.SetFocus

    Set frm = Nothing
    Set ctl = Nothing
    Call Trace(trcExit, Me.Name & ".Form_Activate")
End Sub
Private Sub Form_Load()
    Dim i As Integer
    
    For i = 0 To 31
        dbcFields(i).Visible = False
        dbcFields(i).Enabled = False
        txtFields(i).Visible = False
        txtFields(i).Enabled = False
        lblFields(i).Visible = False
        'lblFields(i).Enabled = False
    Next
End Sub
Private Sub txtFields_GotFocus(Index As Integer)
    TextSelected
End Sub
Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
    If UCase(Right(RS.Fields(Index).Name, 4)) = "SORT" Then KeyPressUcase KeyAscii
End Sub
Private Sub txtFields_Validate(Index As Integer, Cancel As Boolean)
    Dim strField As String
    Dim Operator As String
    Dim Operand As String
    
    strField = Trim(txtFields(Index).Text)
    txtFields(Index).Text = strField
    If strField = vbNullString Then Exit Sub
    
    Operator = vbNullString
    Operand = strField
    If Left(strField, 1) = "=" Then
        Operator = "="
        Operand = Trim(Mid(strField, 2))
    ElseIf Left(strField, 2) = "<=" Then
        Operator = "<="
        Operand = Trim(Mid(strField, 3))
    ElseIf Left(strField, 2) = ">=" Then
        Operator = ">="
        Operand = Trim(Mid(strField, 3))
    ElseIf Left(strField, 1) = "<" Then
        Operator = "<"
        Operand = Trim(Mid(strField, 2))
    ElseIf Left(strField, 1) = ">" Then
        Operator = ">"
        Operand = Trim(Mid(strField, 2))
    ElseIf UCase(Left(strField, 4)) = "NOT " Then
        Operator = "NOT"
        Operand = Trim(Mid(strField, 5))
    ElseIf UCase(Left(strField, 5)) = "LIKE " Then
        Operator = "LIKE"
        Operand = Trim(Mid(strField, 6))
    End If
    
    If Operand <> vbNullString Then
        If Left(Operand, 1) = "'" Or Left(Operand, 1) = "#" Then Operand = Mid(Operand, 2)
        If Right(Operand, 1) = "'" Or Right(Operand, 1) = "#" Then Operand = Left(Operand, Len(Operand) - 1)
        Select Case RS.Fields(Index).Type
            Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar
                If Operator = vbNullString And Right(Operand, 1) <> "%" Then Operand = Operand & "%"
                If Right(Operand, 1) = "%" Then Operator = "LIKE"
                Operand = "'" & SQLQuote(Operand) & "'"
            Case adDate, adDBDate, adDBTime, adDBTimeStamp
                Operand = "#" & Operand & "#"
            Case Else
                Operand = SQLQuote(Operand)
        End Select
    End If
    If Operator = vbNullString Then Operator = "="
    txtFields(Index).Text = Operator & " " & Operand
End Sub


