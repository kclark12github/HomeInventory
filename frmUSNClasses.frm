VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmUSNClasses 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "US Navy Ship Classes"
   ClientHeight    =   5172
   ClientLeft      =   120
   ClientTop       =   348
   ClientWidth     =   7788
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5172
   ScaleWidth      =   7788
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraShips 
      Height          =   3132
      Index           =   0
      Left            =   156
      TabIndex        =   34
      Top             =   780
      Width           =   7452
      Begin VB.TextBox txtYear 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   288
         Left            =   1260
         TabIndex        =   2
         Text            =   "Year"
         Top             =   552
         Width           =   1452
      End
      Begin VB.TextBox txtName 
         Height          =   288
         Left            =   1260
         TabIndex        =   1
         Text            =   "Name"
         Top             =   252
         Width           =   5892
      End
      Begin VB.TextBox txtClassDesc 
         Height          =   288
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "Description"
         Top             =   840
         Width           =   4392
      End
      Begin MSDataListLib.DataCombo dbcClassification 
         Height          =   288
         Left            =   1260
         TabIndex        =   3
         Top             =   852
         Width           =   1452
         _ExtentX        =   2561
         _ExtentY        =   508
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   "Classification"
      End
      Begin VB.Label lblYear 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Year:"
         Height          =   192
         Left            =   792
         TabIndex        =   37
         Top             =   600
         Width           =   384
      End
      Begin VB.Label lblClassification 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Classification:"
         Height          =   192
         Left            =   168
         TabIndex        =   36
         Top             =   900
         Width           =   984
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   192
         Left            =   672
         TabIndex        =   35
         Top             =   300
         Width           =   480
      End
   End
   Begin VB.Frame fraShips 
      Height          =   3132
      Index           =   1
      Left            =   156
      TabIndex        =   25
      Top             =   780
      Width           =   7452
      Begin VB.TextBox txtSpeed 
         Height          =   312
         Left            =   1260
         TabIndex        =   12
         Text            =   "Speed"
         Top             =   2760
         Width           =   6072
      End
      Begin RichTextLib.RichTextBox rtxtDisplacement 
         Height          =   312
         Left            =   1260
         TabIndex        =   5
         Top             =   240
         Width           =   6072
         _ExtentX        =   10710
         _ExtentY        =   550
         _Version        =   393217
         ScrollBars      =   2
         TextRTF         =   $"frmUSNClasses.frx":0000
      End
      Begin RichTextLib.RichTextBox rtxtLength 
         Height          =   312
         Left            =   1260
         TabIndex        =   6
         Top             =   600
         Width           =   6072
         _ExtentX        =   10710
         _ExtentY        =   550
         _Version        =   393217
         MultiLine       =   0   'False
         TextRTF         =   $"frmUSNClasses.frx":00E1
      End
      Begin RichTextLib.RichTextBox rtxtBeam 
         Height          =   312
         Left            =   1260
         TabIndex        =   7
         Top             =   960
         Width           =   6072
         _ExtentX        =   10710
         _ExtentY        =   550
         _Version        =   393217
         MultiLine       =   0   'False
         TextRTF         =   $"frmUSNClasses.frx":01BC
      End
      Begin RichTextLib.RichTextBox rtxtDraft 
         Height          =   312
         Left            =   1260
         TabIndex        =   8
         Top             =   1320
         Width           =   6072
         _ExtentX        =   10710
         _ExtentY        =   550
         _Version        =   393217
         MultiLine       =   0   'False
         TextRTF         =   $"frmUSNClasses.frx":0295
      End
      Begin RichTextLib.RichTextBox rtxtPropulsion 
         Height          =   312
         Left            =   1260
         TabIndex        =   9
         Top             =   1680
         Width           =   6072
         _ExtentX        =   10710
         _ExtentY        =   550
         _Version        =   393217
         MultiLine       =   0   'False
         TextRTF         =   $"frmUSNClasses.frx":036F
      End
      Begin RichTextLib.RichTextBox rtxtBoilers 
         Height          =   312
         Left            =   1260
         TabIndex        =   10
         Top             =   2040
         Width           =   6072
         _ExtentX        =   10710
         _ExtentY        =   550
         _Version        =   393217
         MultiLine       =   0   'False
         TextRTF         =   $"frmUSNClasses.frx":044E
      End
      Begin RichTextLib.RichTextBox rtxtManning 
         Height          =   312
         Left            =   1260
         TabIndex        =   11
         Top             =   2400
         Width           =   6072
         _ExtentX        =   10710
         _ExtentY        =   550
         _Version        =   393217
         MultiLine       =   0   'False
         TextRTF         =   $"frmUSNClasses.frx":052A
      End
      Begin VB.Label lblDisplacement 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Displacement:"
         Height          =   192
         Left            =   180
         TabIndex        =   33
         Top             =   300
         Width           =   1032
      End
      Begin VB.Label lblLength 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Length:"
         Height          =   192
         Left            =   696
         TabIndex        =   32
         Top             =   660
         Width           =   516
      End
      Begin VB.Label lblBeam 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Beam:"
         Height          =   192
         Left            =   744
         TabIndex        =   31
         Top             =   1020
         Width           =   468
      End
      Begin VB.Label lblDraft 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Draft:"
         Height          =   192
         Left            =   840
         TabIndex        =   30
         Top             =   1380
         Width           =   372
      End
      Begin VB.Label lblPropulsion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Propulsion:"
         Height          =   192
         Left            =   408
         TabIndex        =   29
         Top             =   1740
         Width           =   804
      End
      Begin VB.Label lblBoilers 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Boilers:"
         Height          =   192
         Left            =   672
         TabIndex        =   28
         Top             =   2100
         Width           =   540
      End
      Begin VB.Label lblSpeed 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Speed:"
         Height          =   192
         Left            =   684
         TabIndex        =   27
         Top             =   2820
         Width           =   528
      End
      Begin VB.Label lblManning 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Manning:"
         Height          =   192
         Left            =   564
         TabIndex        =   26
         Top             =   2460
         Width           =   648
      End
   End
   Begin VB.Frame fraShips 
      Height          =   3132
      Index           =   2
      Left            =   156
      TabIndex        =   38
      Top             =   780
      Width           =   7452
      Begin RichTextLib.RichTextBox rtxtAircraft 
         Height          =   312
         Left            =   1260
         TabIndex        =   13
         Top             =   240
         Width           =   6072
         _ExtentX        =   10710
         _ExtentY        =   550
         _Version        =   393217
         ScrollBars      =   2
         TextRTF         =   $"frmUSNClasses.frx":0606
      End
      Begin RichTextLib.RichTextBox rtxtMissiles 
         Height          =   312
         Left            =   1260
         TabIndex        =   14
         Top             =   600
         Width           =   6072
         _ExtentX        =   10710
         _ExtentY        =   550
         _Version        =   393217
         MultiLine       =   0   'False
         TextRTF         =   $"frmUSNClasses.frx":06E3
      End
      Begin RichTextLib.RichTextBox rtxtGuns 
         Height          =   312
         Left            =   1260
         TabIndex        =   15
         Top             =   960
         Width           =   6072
         _ExtentX        =   10710
         _ExtentY        =   550
         _Version        =   393217
         MultiLine       =   0   'False
         TextRTF         =   $"frmUSNClasses.frx":07C0
      End
      Begin RichTextLib.RichTextBox rtxtASW 
         Height          =   312
         Left            =   1260
         TabIndex        =   16
         Top             =   1320
         Width           =   6072
         _ExtentX        =   10710
         _ExtentY        =   550
         _Version        =   393217
         MultiLine       =   0   'False
         TextRTF         =   $"frmUSNClasses.frx":0899
      End
      Begin RichTextLib.RichTextBox rtxtRadars 
         Height          =   312
         Left            =   1260
         TabIndex        =   17
         Top             =   1680
         Width           =   6072
         _ExtentX        =   10710
         _ExtentY        =   550
         _Version        =   393217
         MultiLine       =   0   'False
         TextRTF         =   $"frmUSNClasses.frx":0979
      End
      Begin RichTextLib.RichTextBox rtxtSonars 
         Height          =   312
         Left            =   1260
         TabIndex        =   18
         Top             =   2040
         Width           =   6072
         _ExtentX        =   10710
         _ExtentY        =   550
         _Version        =   393217
         MultiLine       =   0   'False
         TextRTF         =   $"frmUSNClasses.frx":0A54
      End
      Begin RichTextLib.RichTextBox rtxtFireControl 
         Height          =   312
         Left            =   1260
         TabIndex        =   19
         Top             =   2400
         Width           =   6072
         _ExtentX        =   10710
         _ExtentY        =   550
         _Version        =   393217
         MultiLine       =   0   'False
         TextRTF         =   $"frmUSNClasses.frx":0B2F
      End
      Begin RichTextLib.RichTextBox rtxtEW 
         Height          =   312
         Left            =   1740
         TabIndex        =   20
         Top             =   2760
         Width           =   5592
         _ExtentX        =   9864
         _ExtentY        =   550
         _Version        =   393217
         MultiLine       =   0   'False
         TextRTF         =   $"frmUSNClasses.frx":0C10
      End
      Begin VB.Label lblAircraft 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Aircraft:"
         Height          =   192
         Left            =   684
         TabIndex        =   46
         Top             =   300
         Width           =   528
      End
      Begin VB.Label lblMissiles 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Missiles:"
         Height          =   192
         Left            =   588
         TabIndex        =   45
         Top             =   660
         Width           =   624
      End
      Begin VB.Label lblGuns 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Guns:"
         Height          =   192
         Left            =   804
         TabIndex        =   44
         Top             =   1020
         Width           =   408
      End
      Begin VB.Label lblASW 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ASW Weapons:"
         Height          =   192
         Left            =   60
         TabIndex        =   43
         Top             =   1380
         Width           =   1152
      End
      Begin VB.Label lblRadars 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Radars:"
         Height          =   192
         Left            =   636
         TabIndex        =   42
         Top             =   1740
         Width           =   576
      End
      Begin VB.Label lblSonars 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sonars:"
         Height          =   192
         Left            =   660
         TabIndex        =   41
         Top             =   2100
         Width           =   552
      End
      Begin VB.Label lblFireControl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fire Control:"
         Height          =   192
         Left            =   360
         TabIndex        =   40
         Top             =   2460
         Width           =   852
      End
      Begin VB.Label lblEW 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Electronic Warfare:"
         Height          =   192
         Left            =   300
         TabIndex        =   39
         Top             =   2820
         Width           =   1356
      End
   End
   Begin VB.Frame fraShips 
      Height          =   3132
      Index           =   3
      Left            =   156
      TabIndex        =   24
      Top             =   780
      Width           =   7452
      Begin RichTextLib.RichTextBox rtxtDescription 
         Height          =   2892
         Left            =   60
         TabIndex        =   21
         Top             =   180
         Width           =   7332
         _ExtentX        =   12933
         _ExtentY        =   5101
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmUSNClasses.frx":0CE7
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   372
      Left            =   6696
      TabIndex        =   23
      Top             =   4500
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   5700
      TabIndex        =   22
      Top             =   4500
      Width           =   972
   End
   Begin MSComctlLib.TabStrip tsClasses 
      Height          =   3612
      Left            =   96
      TabIndex        =   0
      Top             =   360
      Width           =   7572
      _ExtentX        =   13356
      _ExtentY        =   6371
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            Key             =   "General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Characteristics"
            Key             =   "Characteristics"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Weapons"
            Key             =   "Weapons"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Description..."
            Key             =   "Description"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   47
      Top             =   4920
      Width           =   7788
      _ExtentX        =   13737
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
            Object.Width           =   8551
            Key             =   "Message"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   2
            Object.Width           =   1270
            MinWidth        =   1270
            TextSave        =   "11:01 PM"
            Key             =   "Time"
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adodcMain 
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
      Caption         =   ""
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
   Begin MSComctlLib.Toolbar tbMain 
      Align           =   1  'Align Top
      Height          =   288
      Left            =   0
      TabIndex        =   51
      Top             =   0
      Width           =   7788
      _ExtentX        =   13737
      _ExtentY        =   508
      ButtonWidth     =   487
      ButtonHeight    =   466
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlSmall"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Report"
            Object.ToolTipText     =   "Report"
            ImageKey        =   "Report"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SQL"
            Object.ToolTipText     =   "SQL Window"
            ImageKey        =   "SQL"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New record"
            ImageKey        =   "NewRecord"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Modify"
            Object.ToolTipText     =   "Modify record"
            ImageKey        =   "Modify"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete record"
            ImageKey        =   "Delete"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh data"
            ImageKey        =   "Refresh"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Filter"
            Object.ToolTipText     =   "Filter"
            ImageKey        =   "Filter"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Search"
            Object.ToolTipText     =   "Search"
            ImageKey        =   "Search"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "List"
            Object.ToolTipText     =   "List all records"
            ImageKey        =   "List"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList imlSmall 
      Left            =   240
      Top             =   60
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNClasses.frx":0DCA
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNClasses.frx":10E6
            Key             =   "Warning"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNClasses.frx":140E
            Key             =   "List"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNClasses.frx":1736
            Key             =   "xNew"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNClasses.frx":3EEA
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNClasses.frx":433E
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNClasses.frx":4792
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNClasses.frx":525E
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNClasses.frx":5586
            Key             =   "Sort"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNClasses.frx":59DA
            Key             =   "SQL"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNClasses.frx":5E2E
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNClasses.frx":6282
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNClasses.frx":66DA
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNClasses.frx":6836
            Key             =   "Blank"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNClasses.frx":6992
            Key             =   "NewRecord"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblIDlabel 
      AutoSize        =   -1  'True
      Caption         =   "ID:"
      Height          =   192
      Left            =   96
      TabIndex        =   50
      Top             =   4596
      Width           =   192
   End
   Begin VB.Label lblID 
      AutoSize        =   -1  'True
      Caption         =   "lblID"
      Height          =   192
      Left            =   360
      TabIndex        =   49
      Top             =   4596
      Width           =   324
   End
   Begin VB.Label lblA 
      AutoSize        =   -1  'True
      Caption         =   "A"
      Height          =   192
      Left            =   1620
      TabIndex        =   48
      Top             =   4560
      Visible         =   0   'False
      Width           =   108
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileReport 
         Caption         =   "&Report"
      End
      Begin VB.Menu mnuFileSQL 
         Caption         =   "&SQL"
      End
      Begin VB.Menu mnuFileSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuRecords 
      Caption         =   "&Records"
      Begin VB.Menu mnuRecordsNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuRecordsModify 
         Caption         =   "&Modify"
      End
      Begin VB.Menu mnuRecordsDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuRecordsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecordsRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuRecordsSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecordsFilter 
         Caption         =   "&Filter"
      End
      Begin VB.Menu mnuRecordsSearch 
         Caption         =   "&Search"
      End
      Begin VB.Menu mnuRecordsSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecordsList 
         Caption         =   "&List"
      End
   End
End
Attribute VB_Name = "frmUSNClasses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents rsMain As ADODB.Recordset
Attribute rsMain.VB_VarHelpID = -1
Dim rsClassifications As New ADODB.Recordset
Private Sub cmdCancel_Click()
    CancelCommand Me, rsMain
End Sub
Private Sub cmdOK_Click()
    OKCommand Me, rsMain
End Sub
Private Sub Form_Load()
    EstablishConnection adoConn
    
    Set rsMain = New ADODB.Recordset
    rsMain.CursorLocation = adUseClient
    SQLmain = "select * from [Class] order by Classification, Year"
    SQLfilter = vbNullString
    SQLkey = "Reference"
    rsMain.Open SQLmain, adoConn, adOpenKeyset, adLockBatchOptimistic
    DBcollection.Add "rsMain", rsMain
    rsMain.MoveFirst
    
    rsClassifications.CursorLocation = adUseClient
    rsClassifications.Open "select * from [Classification] order by Type", adoConn, adOpenStatic, adLockReadOnly
    DBcollection.Add "rsClassifications", rsClassifications
    
    Set adodcMain.Recordset = rsMain
    BindField lblID, "ID", rsMain
    BindFields
    
    Set tsClasses.SelectedItem = tsClasses.Tabs(1)
    ProtectFields Me
    mode = modeDisplay
    fTransaction = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Cancel = CloseConnection(Me)
End Sub
Private Sub mnuRecordsFilter_Click()
    FilterCommand Me, rsMain, SQLkey
End Sub
Private Sub mnuRecordsDelete_Click()
    DeleteCommand Me, rsMain
End Sub
Private Sub mnuRecordsList_Click()
    ListCommand Me, rsMain
End Sub
Private Sub mnuRecordsModify_Click()
    ModifyCommand Me
    
    Set tsClasses.SelectedItem = tsClasses.Tabs(1)
    txtName.SetFocus
End Sub
Private Sub mnuRecordsNew_Click()
    NewCommand Me, rsMain
    
    txtYear.Text = Format(Now(), "0000")
    Set tsClasses.SelectedItem = tsClasses.Tabs(1)
    txtName.SetFocus
End Sub
Private Sub mnuRecordsRefresh_Click()
    RefreshCommand rsMain, "ID"
End Sub
Private Sub mnuRecordsSearch_Click()
    SearchCommand Me, rsMain, SQLkey
End Sub
Private Sub mnuFileExit_Click()
    Unload Me
End Sub
Private Sub mnuFileReport_Click()
    ReportCommand Me, rsMain, App.Path & "\Reports\USN Ship Classes.rpt"
End Sub
Private Sub mnuFileSQL_Click()
    SQLCommand "Class"
End Sub
Private Sub rsMain_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Dim Caption As String
    
    If Not pRecordset.BOF And Not pRecordset.EOF Then
        Caption = "Reference #" & pRecordset.Bookmark & ": " & pRecordset("Name")
        If pRecordset("Year") > 0 Then Caption = Caption & " of " & pRecordset("Year")
    End If
    UpdatePosition Me, Caption, pRecordset
    DefaultClassificationDesc
End Sub
Private Sub tbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Report"
            mnuFileReport_Click
        Case "SQL"
            mnuFileSQL_Click
        Case "New"
            mnuRecordsNew_Click
        Case "Modify"
            mnuRecordsModify_Click
        Case "Delete"
            mnuRecordsDelete_Click
        Case "Refresh"
            mnuRecordsRefresh_Click
        Case "Filter"
            mnuRecordsFilter_Click
        Case "Search"
            mnuRecordsSearch_Click
        Case "List"
            mnuRecordsList_Click
    End Select
End Sub
Private Sub tsClasses_Click()
    Dim i As Integer
    
    With tsClasses
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
'=================================================================================
Private Sub BindFields()
    'General
    BindField txtName, "Name", rsMain
    BindField dbcClassification, "Classification", rsMain, rsClassifications, "Type", "Type"
    BindField txtClassDesc, "Description", rsClassifications
    BindField txtYear, "Year", rsMain
    
    'Characteristics...
    BindField rtxtDisplacement, "Displacement", rsMain
    BindField rtxtLength, "Length", rsMain
    BindField rtxtBeam, "Beam", rsMain
    BindField rtxtDraft, "Draft", rsMain
    BindField rtxtPropulsion, "Propulsion", rsMain
    BindField rtxtBoilers, "Boilers", rsMain
    BindField txtSpeed, "Speed", rsMain
    BindField rtxtManning, "Manning", rsMain
    
    'Weapons...
    BindField rtxtAircraft, "Aircraft", rsMain
    BindField rtxtMissiles, "Missiles", rsMain
    BindField rtxtGuns, "Guns", rsMain
    BindField rtxtASW, "ASW Weapons", rsMain
    BindField rtxtRadars, "Radars", rsMain
    BindField rtxtSonars, "Sonars", rsMain
    BindField rtxtFireControl, "Fire Control", rsMain
    BindField rtxtEW, "EW", rsMain
    
    'Description...
    BindField rtxtDescription, "Description", rsMain
End Sub
Private Sub dbcClassification_Validate(Cancel As Boolean)
    If Not dbcClassification.Enabled Then Exit Sub
    If dbcClassification.Text = vbNullString And txtClassDesc.Text <> "Unused" And txtClassDesc.Text <> "Unassigned" Then
        MsgBox "Classification must be specified!", vbExclamation, Me.Caption
        dbcClassification.SetFocus
        Cancel = True
    End If
    If dbcValidate(rsMain("Classification"), dbcClassification) = 0 Then Cancel = True
    If rsClassifications.Bookmark <> dbcClassification.SelectedItem Then rsClassifications.Bookmark = dbcClassification.SelectedItem
End Sub
Private Sub DefaultClassificationDesc()
    If rsClassifications Is Nothing Then Exit Sub
    If rsClassifications.State = adStateClosed Then Exit Sub
    If rsMain.BOF Or rsMain.EOF Then Exit Sub
    
    rsClassifications.MoveFirst
    rsClassifications.Find "ID=" & rsMain("ClassificationID").Value
End Sub
Private Sub txtYear_GotFocus()
    TextSelected
End Sub
Private Sub txtYear_Validate(Cancel As Boolean)
    On Error Resume Next
    txtYear.Text = Format(txtYear.Text, "0000")
    If txtYear.Text = 0 Or (txtYear.Text >= 1776 And txtYear.Text <= 2999) Then
    Else
        MsgBox "Invalid year format", vbExclamation
        Cancel = True
        Exit Sub
    End If
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

