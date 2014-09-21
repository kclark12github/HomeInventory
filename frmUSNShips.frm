VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmUSNShips 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "US Navy Ships"
   ClientHeight    =   5160
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   7764
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   7764
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraShips 
      Height          =   3132
      Index           =   0
      Left            =   156
      TabIndex        =   17
      Top             =   780
      Width           =   7452
      Begin VB.TextBox txtZipCode 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dddd, MMMM dd, yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   288
         Left            =   5532
         TabIndex        =   8
         Text            =   "Zip Code"
         Top             =   1740
         Width           =   1812
      End
      Begin VB.TextBox txtLocalURL 
         Height          =   288
         Left            =   1440
         TabIndex        =   11
         Text            =   "URL_Local"
         Top             =   2760
         Width           =   5832
      End
      Begin VB.TextBox txtURL 
         Height          =   288
         Left            =   1440
         TabIndex        =   10
         Text            =   "URL_Internet"
         Top             =   2460
         Width           =   5832
      End
      Begin VB.TextBox txtClassDesc 
         Height          =   288
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "Description"
         Top             =   1140
         Width           =   4392
      End
      Begin VB.TextBox txtNumber 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   288
         Left            =   5556
         TabIndex        =   2
         Text            =   "Number"
         Top             =   240
         Width           =   1032
      End
      Begin VB.TextBox txtDesignation 
         Height          =   288
         Left            =   1440
         TabIndex        =   1
         Text            =   "HullNumber"
         Top             =   240
         Width           =   2232
      End
      Begin VB.TextBox txtName 
         Height          =   288
         Left            =   1440
         TabIndex        =   3
         Text            =   "Name"
         Top             =   552
         Width           =   5892
      End
      Begin VB.TextBox txtCommissioned 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dddd, MMMM dd, yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   288
         Left            =   1440
         TabIndex        =   9
         Text            =   "Commissioned"
         Top             =   2052
         Width           =   3252
      End
      Begin MSDataListLib.DataCombo dbcClassification 
         Height          =   288
         Left            =   1440
         TabIndex        =   5
         Top             =   1152
         Width           =   1452
         _ExtentX        =   2561
         _ExtentY        =   508
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   "Classification"
      End
      Begin MSDataListLib.DataCombo dbcCommand 
         Height          =   288
         Left            =   1440
         TabIndex        =   6
         Top             =   1440
         Width           =   3252
         _ExtentX        =   5736
         _ExtentY        =   508
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   "Command"
      End
      Begin MSDataListLib.DataCombo dbcClass 
         Height          =   288
         Left            =   1440
         TabIndex        =   4
         Top             =   840
         Width           =   5892
         _ExtentX        =   10393
         _ExtentY        =   508
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   "Class"
      End
      Begin MSDataListLib.DataCombo dbcHomePort 
         Height          =   288
         Left            =   1440
         TabIndex        =   7
         Top             =   1740
         Width           =   3252
         _ExtentX        =   5736
         _ExtentY        =   508
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   "HomePort"
      End
      Begin VB.Label lblHomePort 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Home Port:"
         Height          =   192
         Left            =   516
         TabIndex        =   29
         Top             =   1788
         Width           =   804
      End
      Begin VB.Label lblZipCode 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Zip Code:"
         Height          =   192
         Left            =   4740
         TabIndex        =   28
         Top             =   1788
         Width           =   696
      End
      Begin VB.Label lblLocalWebSite 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Local WebSite:"
         Height          =   192
         Left            =   240
         TabIndex        =   27
         Top             =   2808
         Width           =   1092
      End
      Begin VB.Label lblWebSite 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "WebSite:"
         Height          =   192
         Left            =   684
         TabIndex        =   26
         Top             =   2508
         Width           =   660
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   7320
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label lblClass 
         AutoSize        =   -1  'True
         Caption         =   "Class:"
         Height          =   192
         Left            =   888
         TabIndex        =   25
         Top             =   900
         Width           =   444
      End
      Begin VB.Label lblCommand 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Command:"
         Height          =   192
         Left            =   552
         TabIndex        =   24
         Top             =   1488
         Width           =   780
      End
      Begin VB.Label lblNumber 
         AutoSize        =   -1  'True
         Caption         =   "Number:"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   192
         Left            =   4824
         TabIndex        =   23
         Top             =   288
         Width           =   612
      End
      Begin VB.Label lblDesignation 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Designation:"
         Height          =   192
         Left            =   432
         TabIndex        =   22
         Top             =   300
         Width           =   900
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   192
         Left            =   852
         TabIndex        =   21
         Top             =   600
         Width           =   480
      End
      Begin VB.Label lblClassification 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Classification:"
         Height          =   192
         Left            =   348
         TabIndex        =   20
         Top             =   1200
         Width           =   984
      End
      Begin VB.Label lblCommissioned 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Commissioned:"
         Height          =   192
         Left            =   240
         TabIndex        =   18
         Top             =   2100
         Width           =   1116
      End
   End
   Begin VB.Frame fraShips 
      Height          =   3132
      Index           =   2
      Left            =   156
      TabIndex        =   32
      Top             =   780
      Width           =   7452
      Begin RichTextLib.RichTextBox rtxtHistory 
         Height          =   2892
         Left            =   60
         TabIndex        =   33
         Top             =   180
         Width           =   7332
         _ExtentX        =   12933
         _ExtentY        =   5101
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmUSNShips.frx":0000
      End
   End
   Begin VB.Frame fraShips 
      Height          =   3132
      Index           =   4
      Left            =   156
      TabIndex        =   36
      Top             =   780
      Width           =   7452
   End
   Begin VB.Frame fraShips 
      Height          =   3132
      Index           =   3
      Left            =   156
      TabIndex        =   34
      Top             =   780
      Width           =   7452
      Begin RichTextLib.RichTextBox rtxtMoreHistory 
         Height          =   2892
         Left            =   60
         TabIndex        =   35
         Top             =   180
         Width           =   7332
         _ExtentX        =   12933
         _ExtentY        =   5101
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmUSNShips.frx":00DF
      End
   End
   Begin VB.Frame fraShips 
      Height          =   3132
      Index           =   1
      Left            =   156
      TabIndex        =   31
      Top             =   780
      Width           =   7452
   End
   Begin MSComctlLib.TabStrip tsShips 
      Height          =   3612
      Left            =   96
      TabIndex        =   0
      Top             =   360
      Width           =   7572
      _ExtentX        =   13356
      _ExtentY        =   6371
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            Key             =   "General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Capabilities"
            Key             =   "Capabilities"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "History..."
            Key             =   "History"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "More History..."
            Key             =   "MoreHistory"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Images"
            Key             =   "Images"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   5700
      TabIndex        =   12
      Top             =   4500
      Width           =   972
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   372
      Left            =   6696
      TabIndex        =   13
      Top             =   4500
      Width           =   972
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   14
      Top             =   4908
      Width           =   7764
      _ExtentX        =   13695
      _ExtentY        =   445
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1270
            MinWidth        =   1270
            Key             =   "Position"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "Status"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8509
            Key             =   "Message"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   2
            Object.Width           =   1270
            MinWidth        =   1270
            TextSave        =   "1:18 AM"
            Key             =   "Time"
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adodcHobby 
      Height          =   312
      Left            =   96
      Top             =   4020
      Width           =   7572
      _ExtentX        =   13356
      _ExtentY        =   550
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "0"
      OLEDBString     =   "0"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlSmall 
      Left            =   3312
      Top             =   4380
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNShips.frx":01C3
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNShips.frx":04DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNShips.frx":0807
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNShips.frx":0B2F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNShips.frx":32E3
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNShips.frx":3737
            Key             =   "List"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNShips.frx":4203
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNShips.frx":4657
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNShips.frx":5123
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNShips.frx":544B
            Key             =   "Sort"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNShips.frx":589F
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNShips.frx":5CF3
            Key             =   "SQL"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNShips.frx":6147
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlLarge 
      Left            =   2892
      Top             =   4380
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNShips.frx":659B
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNShips.frx":69EF
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNShips.frx":74BB
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNShips.frx":77D7
            Key             =   "List"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNShips.frx":82A3
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNShips.frx":86F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNShips.frx":AEAB
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNShips.frx":B2FF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbAction 
      Align           =   1  'Align Top
      Height          =   288
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   7764
      _ExtentX        =   13695
      _ExtentY        =   508
      ButtonWidth     =   487
      ButtonHeight    =   466
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlSmall"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "List"
            Object.ToolTipText     =   "List all records"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh data"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Filter"
            Object.ToolTipText     =   "Filter"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New record"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Modify"
            Object.ToolTipText     =   "Modify record"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete record"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Report"
            Object.ToolTipText     =   "Report"
            ImageIndex      =   7
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Test"
                  Text            =   "Test"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Test2"
                  Text            =   "Test2"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SQL"
            ImageIndex      =   13
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label lblID 
      AutoSize        =   -1  'True
      Caption         =   "lblID"
      Height          =   192
      Left            =   360
      TabIndex        =   16
      Top             =   4596
      Width           =   324
   End
   Begin VB.Label lblIDlabel 
      AutoSize        =   -1  'True
      Caption         =   "ID:"
      Height          =   192
      Left            =   96
      TabIndex        =   15
      Top             =   4596
      Width           =   192
   End
   Begin VB.Menu mnuAction 
      Caption         =   "&Action"
      Begin VB.Menu mnuActionList 
         Caption         =   "&List"
      End
      Begin VB.Menu mnuActionRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuActionFilter 
         Caption         =   "&Filter"
      End
      Begin VB.Menu mnuActionSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActionNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuActionModify 
         Caption         =   "&Modify"
      End
      Begin VB.Menu mnuActionDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuActionSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActionReport 
         Caption         =   "&Report"
      End
      Begin VB.Menu mnuActionSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActionSQL 
         Caption         =   "&SQL"
      End
   End
End
Attribute VB_Name = "frmUSNShips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoConn As ADODB.Connection
Dim WithEvents rsMain As ADODB.Recordset
Attribute rsMain.VB_VarHelpID = -1
Dim rsCommands As New ADODB.Recordset
Dim rsHomePorts As New ADODB.Recordset
Dim rsClasses As New ADODB.Recordset
Dim rsClassifications As New ADODB.Recordset
Dim mode As ActionMode
Dim fTransaction As Boolean
Dim DBinfo As DataBaseInfo
Private Sub cmdCancel_Click()
    Select Case mode
        Case modeDisplay
            Unload Me
        Case modeAdd, modeModify
            rsMain.CancelUpdate
            If mode = modeAdd And Not rsMain.EOF Then rsMain.MoveLast
            adoConn.RollbackTrans
            fTransaction = False
            frmMain.ProtectFields Me
            mode = modeDisplay
            adodcHobby.Enabled = True
    End Select
End Sub
Private Sub cmdOK_Click()
    Select Case mode
        Case modeDisplay
            Unload Me
        Case modeAdd, modeModify
            'Why we need to do this is buggy...
            rsMain("ClassID") = dbcClass.BoundText
            rsMain("ClassificationID") = dbcClassification.BoundText
            rsMain("Command") = dbcCommand.BoundText
            rsMain("HomePort") = dbcHomePort.BoundText
            rsMain.UpdateBatch
            adoConn.CommitTrans
            fTransaction = False
            frmMain.ProtectFields Me
            mode = modeDisplay
            adodcHobby.Enabled = True
            
            mnuActionRefresh_Click
    End Select
End Sub
Private Sub dbcClass_Validate(Cancel As Boolean)
    If Not dbcClass.Enabled Then Exit Sub
    If dbcClass.Text = vbNullString Then
        MsgBox "Class must be specified!", vbExclamation, Me.Caption
        dbcClass.SetFocus
        Cancel = True
    End If
    If rsClasses.Bookmark <> dbcClass.SelectedItem Then rsClasses.Bookmark = dbcClass.SelectedItem
End Sub
Private Sub dbcClassification_Validate(Cancel As Boolean)
    If Not dbcClassification.Enabled Then Exit Sub
    If dbcClassification.Text = vbNullString Then
        MsgBox "Classification must be specified!", vbExclamation, Me.Caption
        dbcClassification.SetFocus
        Cancel = True
    End If
    If rsClassifications.Bookmark <> dbcClassification.SelectedItem Then rsClassifications.Bookmark = dbcClassification.SelectedItem
End Sub
Private Sub dbcCommand_GotFocus()
    TextSelected
End Sub
Private Sub dbcCommand_Validate(Cancel As Boolean)
    If dbcCommand.Text = vbNullString Then dbcCommand.Text = "Unknown"
    If rsCommands.Bookmark <> dbcCommand.SelectedItem Then rsCommands.Bookmark = dbcCommand.SelectedItem
End Sub
Private Sub dbcHomePort_GotFocus()
    TextSelected
End Sub
Private Sub dbcHomePort_Validate(Cancel As Boolean)
    If dbcHomePort.Text = vbNullString Then dbcHomePort.Text = "Unknown"
    If rsHomePorts.Bookmark <> dbcHomePort.SelectedItem Then rsHomePorts.Bookmark = dbcHomePort.SelectedItem
End Sub
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
    rsMain.Open "select * from [Ships] order by HullNumber", adoConn, adOpenKeyset, adLockBatchOptimistic
    rsMain.MoveFirst
    
    rsCommands.CursorLocation = adUseClient
    rsCommands.Open "select distinct Command from [Ships] order by Command", adoConn, adOpenStatic, adLockReadOnly
    
    rsHomePorts.CursorLocation = adUseClient
    rsHomePorts.Open "select distinct HomePort from [Ships] order by HomePort", adoConn, adOpenStatic, adLockReadOnly
    
    rsClasses.CursorLocation = adUseClient
    rsClasses.Open "select ID, Name from [Class] order by Name", adoConn, adOpenStatic, adLockReadOnly
    
    rsClassifications.CursorLocation = adUseClient
    rsClassifications.Open "select * from [Classification] order by Type", adoConn, adOpenStatic, adLockReadOnly
    
    Set adodcHobby.Recordset = rsMain
    frmMain.BindField lblID, "ID", rsMain
    
    'General
    frmMain.BindField txtDesignation, "HullNumber", rsMain
    frmMain.BindField txtNumber, "Number", rsMain
    frmMain.BindField txtName, "Name", rsMain
    frmMain.BindField dbcClass, "ClassID", rsMain, rsClasses, "ID", "Name"
    frmMain.BindField dbcClassification, "ClassificationID", rsMain, rsClassifications, "ID", "Type"
    frmMain.BindField txtClassDesc, "Description", rsClassifications
    frmMain.BindField dbcCommand, "Command", rsMain, rsCommands, "Command", "Command"
    frmMain.BindField dbcHomePort, "HomePort", rsMain, rsHomePorts, "HomePort", "HomePort"
    frmMain.BindField txtZipCode, "Zip Code", rsMain
    frmMain.BindField txtCommissioned, "Commissioned", rsMain
    frmMain.BindField txtURL, "URL_Internet", rsMain
    frmMain.BindField txtLocalURL, "URL_Local", rsMain
    'Capabilities...
    
    'History...
    frmMain.BindField rtxtHistory, "History", rsMain
    
    'More History...
    frmMain.BindField rtxtMoreHistory, "More History", rsMain
    
    'Images...
    
    Set tsShips.SelectedItem = tsShips.Tabs(1)
    frmMain.ProtectFields Me
    mode = modeDisplay
    fTransaction = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If fTransaction Then
        MsgBox "Please complete the current operation before closing the window.", vbExclamation, Me.Caption
        Cancel = 1
        Exit Sub
    End If
    CloseRecordset rsMain, True
    CloseRecordset rsCommands, True
    CloseRecordset rsHomePorts, True
    CloseRecordset rsClasses, True
    CloseRecordset rsClassifications, True
    
    On Error Resume Next
    adoConn.Close
    If Err.Number = 3246 Then
        adoConn.RollbackTrans
        fTransaction = False
        adoConn.Close
    End If
    Set adoConn = Nothing
End Sub
Private Sub mnuActionList_Click()
    Dim frm As Form
    
    Load frmList
    frmList.Caption = Me.Caption & " List"
    If frmMain.Width > Me.Width And frmMain.Height > Me.Height Then
        Set frm = frmMain
    Else
        Set frm = Me
    End If
    frmList.Top = frm.Top
    frmList.Left = frm.Left
    frmList.Width = frm.Width
    frmList.Height = frm.Height
    
    Set frmList.rsList = rsMain
    
    adoConn.BeginTrans
    fTransaction = True
    frmList.Show vbModal
    If rsMain.Filter <> vbNullString And rsMain.Filter <> 0 Then
        sbStatus.Panels("Message").Text = "Filter: " & rsMain.Filter
    End If
    adoConn.CommitTrans
    fTransaction = False
End Sub
Private Sub mnuActionRefresh_Click()
    Dim SaveBookmark As String
    
    On Error Resume Next
    SaveBookmark = rsMain("Reference")
    rsMain.Requery
    rsMain.Find "Reference='" & SQLQuote(SaveBookmark) & "'"
    rsClasses.Requery
    rsClassifications.Requery
    rsCommands.Requery
    rsHomePorts.Requery
End Sub
Private Sub mnuActionFilter_Click()
    Dim frm As Form
    
    Load frmFilter
    frmFilter.Caption = Me.Caption & " Filter"
    If frmMain.Width > Me.Width And frmMain.Height > Me.Height Then
        Set frm = frmMain
    Else
        Set frm = Me
    End If
    frmFilter.Top = frm.Top
    frmFilter.Left = frm.Left
    frmFilter.Width = frm.Width
    frmFilter.Height = frm.Height
    
    Set frmFilter.RS = rsMain
    frmFilter.Show vbModal
    If rsMain.Filter <> vbNullString And rsMain.Filter <> 0 Then
        sbStatus.Panels("Message").Text = "Filter: " & rsMain.Filter
    End If
End Sub
Private Sub mnuActionNew_Click()
    mode = modeAdd
    frmMain.OpenFields Me
    txtClassDesc.Locked = True
    txtClassDesc.BackColor = vbButtonFace
    adodcHobby.Enabled = False
    rsMain.AddNew
    adoConn.BeginTrans
    fTransaction = True
    
    txtCommissioned.Text = Format(Now(), "dddd, MMMM dd, yyyy")
    txtDesignation.SetFocus
End Sub
Private Sub mnuActionDelete_Click()
    mode = modeDelete
    If MsgBox("Are you sure you want to permanently delete this record...?", vbYesNo, Me.Caption) = vbYes Then
        rsMain.Delete
        rsMain.MoveNext
        If rsMain.EOF Then rsMain.MoveLast
    End If
    mode = modeDisplay
End Sub
Private Sub mnuActionModify_Click()
    mode = modeModify
    frmMain.OpenFields Me
    txtClassDesc.Locked = True
    txtClassDesc.BackColor = vbButtonFace
    adodcHobby.Enabled = False
    adoConn.BeginTrans
    fTransaction = True
    
    txtDesignation.SetFocus
End Sub
Private Sub mnuActionReport_Click()
    Dim frm As Form
    Dim scrApplication As New CRAXDRT.Application
    Dim Report As New CRAXDRT.Report
    Dim vRS As ADODB.Recordset
    
    MakeVirtualRecordset adoConn, rsMain, vRS
    
    Load frmViewReport
    frmViewReport.Caption = Me.Caption & " Report"
    If frmMain.Width > Me.Width And frmMain.Height > Me.Height Then
        Set frm = frmMain
    Else
        Set frm = Me
    End If
    frmViewReport.Top = frm.Top
    frmViewReport.Left = frm.Left
    frmViewReport.Width = frm.Width
    frmViewReport.Height = frm.Height
    frmViewReport.WindowState = vbMaximized
    
    Set Report = scrApplication.OpenReport(App.Path & "\Reports\USN Ships.rpt", crOpenReportByTempCopy)
    Report.Database.SetDataSource vRS, 3, 1
    Report.ReadRecords
    
    frmViewReport.scrViewer.ReportSource = Report
    frmViewReport.Show vbModal
    
    Set scrApplication = Nothing
    Set Report = Nothing
    vRS.Close
    Set vRS = Nothing
End Sub
Private Sub mnuActionSQL_Click()
    Load frmSQL
    Set frmSQL.cnSQL = adoConn
    frmSQL.txtDatabase.Text = DBinfo.PathName
    frmSQL.dbcTables.BoundText = "Kits"
    frmSQL.Show vbModal
End Sub
Private Sub rsMain_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Dim Caption As String
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    If rsMain.BOF And rsMain.EOF Then
        Caption = "No Records"
    ElseIf rsMain.EOF Then
        Caption = "EOF"
    ElseIf rsMain.BOF Then
        Caption = "BOF"
    Else
        Caption = "Reference #" & rsMain.Bookmark & ": " & rsMain("HullNumber") & " " & rsMain("Name")
        
        i = InStr(Caption, "&")
        If i > 0 Then Caption = Left(Caption, i) & "&" & Mid(Caption, i + 1)
        If rsMain.Filter <> vbNullString And rsMain.Filter <> 0 Then
            sbStatus.Panels("Message").Text = "Filter: " & rsMain.Filter
        End If
        sbStatus.Panels("Position").Text = "Record " & rsMain.Bookmark & " of " & rsMain.RecordCount
    End If
    
    If Not rsClassifications Is Nothing Then
        If rsClassifications.State <> adStateClosed Then
            rsClassifications.MoveFirst
            rsClassifications.Find "ID=" & rsMain("ClassificationID").Value
        End If
    End If
    adodcHobby.Caption = Caption
    Exit Sub

ErrorHandler:
    MsgBox Err.Description & " (Error " & Err.Number & ")", vbExclamation, Me.Caption
    Resume Next
End Sub
Private Sub tbAction_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "List"
            mnuActionList_Click
        Case "Refresh"
            mnuActionRefresh_Click
        Case "Filter"
            mnuActionFilter_Click
        Case "New"
            mnuActionNew_Click
        Case "Modify"
            mnuActionModify_Click
        Case "Delete"
            mnuActionDelete_Click
        Case "Report"
            mnuActionReport_Click
        Case "SQL"
            mnuActionSQL_Click
    End Select
End Sub
Private Sub tsShips_Click()
    Dim i As Integer
    
    With tsShips
        For i = 0 To .Tabs.Count - 1
            If i = .SelectedItem.Index - 1 Then
                fraShips(i).Enabled = True
                fraShips(i).ZOrder
            Else
                fraShips(i).Enabled = False
            End If
        Next
    End With
End Sub
Private Sub txtDesignation_GotFocus()
    TextSelected
End Sub
Private Sub txtDesignation_KeyPress(KeyAscii As Integer)
    KeyPressUcase KeyAscii
End Sub
Private Sub txtCommissioned_GotFocus()
    TextSelected
End Sub
Private Sub txtCommissioned_Validate(Cancel As Boolean)
    On Error Resume Next
    txtCommissioned.Text = Format(txtCommissioned.Text, "dddd, MMMM dd, yyyy")
    If Not IsDate(txtCommissioned.Text) Then
        MsgBox "Invalid date format", vbExclamation
        Cancel = True
        Exit Sub
    End If
End Sub
Private Sub txtLocalURL_GotFocus()
    TextSelected
End Sub
Private Sub txtName_GotFocus()
    TextSelected
End Sub
Private Sub txtName_Validate(Cancel As Boolean)
    If Not txtName.Enabled Then Exit Sub
    If txtName.Text = vbNullString Then
        MsgBox "Name must be specified!", vbExclamation, Me.Caption
        Cancel = True
    End If
End Sub
Private Sub txtNumber_GotFocus()
    TextSelected
End Sub
Private Sub txtURL_GotFocus()
    TextSelected
End Sub
Private Sub txtZipCode_GotFocus()
    TextSelected
End Sub
