VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
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
      Index           =   1
      Left            =   156
      TabIndex        =   32
      Top             =   780
      Width           =   7452
      Begin VB.TextBox txtSpeed 
         Height          =   312
         Left            =   1260
         TabIndex        =   47
         Text            =   "Speed"
         Top             =   2760
         Width           =   6072
      End
      Begin RichTextLib.RichTextBox rtxtDisplacement 
         Height          =   312
         Left            =   1260
         TabIndex        =   35
         Top             =   240
         Width           =   6072
         _ExtentX        =   10710
         _ExtentY        =   550
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmUSNShips.frx":0000
      End
      Begin RichTextLib.RichTextBox rtxtLength 
         Height          =   312
         Left            =   1260
         TabIndex        =   37
         Top             =   600
         Width           =   6072
         _ExtentX        =   10710
         _ExtentY        =   550
         _Version        =   393217
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         TextRTF         =   $"frmUSNShips.frx":00E1
      End
      Begin RichTextLib.RichTextBox rtxtBeam 
         Height          =   312
         Left            =   1260
         TabIndex        =   39
         Top             =   960
         Width           =   6072
         _ExtentX        =   10710
         _ExtentY        =   550
         _Version        =   393217
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         TextRTF         =   $"frmUSNShips.frx":01BC
      End
      Begin RichTextLib.RichTextBox rtxtDraft 
         Height          =   312
         Left            =   1260
         TabIndex        =   41
         Top             =   1320
         Width           =   6072
         _ExtentX        =   10710
         _ExtentY        =   550
         _Version        =   393217
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         TextRTF         =   $"frmUSNShips.frx":0295
      End
      Begin RichTextLib.RichTextBox rtxtPropulsion 
         Height          =   312
         Left            =   1260
         TabIndex        =   43
         Top             =   1680
         Width           =   6072
         _ExtentX        =   10710
         _ExtentY        =   550
         _Version        =   393217
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         TextRTF         =   $"frmUSNShips.frx":036F
      End
      Begin RichTextLib.RichTextBox rtxtBoilers 
         Height          =   312
         Left            =   1260
         TabIndex        =   45
         Top             =   2040
         Width           =   6072
         _ExtentX        =   10710
         _ExtentY        =   550
         _Version        =   393217
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         TextRTF         =   $"frmUSNShips.frx":044E
      End
      Begin RichTextLib.RichTextBox rtxtManning 
         Height          =   312
         Left            =   1260
         TabIndex        =   46
         Top             =   2400
         Width           =   6072
         _ExtentX        =   10710
         _ExtentY        =   550
         _Version        =   393217
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         TextRTF         =   $"frmUSNShips.frx":052A
      End
      Begin VB.Label lblManning 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Manning:"
         Height          =   192
         Left            =   564
         TabIndex        =   49
         Top             =   2460
         Width           =   648
      End
      Begin VB.Label lblSpeed 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Speed:"
         Height          =   192
         Left            =   684
         TabIndex        =   48
         Top             =   2820
         Width           =   528
      End
      Begin VB.Label lblBoilers 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Boilers:"
         Height          =   192
         Left            =   672
         TabIndex        =   44
         Top             =   2100
         Width           =   540
      End
      Begin VB.Label lblPropulsion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Propulsion:"
         Height          =   192
         Left            =   408
         TabIndex        =   42
         Top             =   1740
         Width           =   804
      End
      Begin VB.Label lblDraft 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Draft:"
         Height          =   192
         Left            =   840
         TabIndex        =   40
         Top             =   1380
         Width           =   372
      End
      Begin VB.Label lblBeam 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Beam:"
         Height          =   192
         Left            =   744
         TabIndex        =   38
         Top             =   1020
         Width           =   468
      End
      Begin VB.Label lblLength 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Length:"
         Height          =   192
         Left            =   696
         TabIndex        =   36
         Top             =   660
         Width           =   516
      End
      Begin VB.Label lblDisplacement 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Displacement:"
         Height          =   192
         Left            =   180
         TabIndex        =   34
         Top             =   300
         Width           =   1032
      End
   End
   Begin VB.Frame fraShips 
      Height          =   3132
      Index           =   0
      Left            =   156
      TabIndex        =   14
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
         TabIndex        =   7
         Text            =   "Zip Code"
         Top             =   1740
         Width           =   1812
      End
      Begin VB.TextBox txtLocalURL 
         Height          =   288
         Left            =   1440
         TabIndex        =   10
         Text            =   "URL_Local"
         Top             =   2760
         Width           =   5832
      End
      Begin VB.TextBox txtURL 
         Height          =   288
         Left            =   1440
         TabIndex        =   9
         Text            =   "URL_Internet"
         Top             =   2460
         Width           =   5832
      End
      Begin VB.TextBox txtClassDesc 
         Height          =   288
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "Description"
         Top             =   1140
         Width           =   4392
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
         TabIndex        =   2
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
         TabIndex        =   8
         Text            =   "Commissioned"
         Top             =   2052
         Width           =   3252
      End
      Begin MSDataListLib.DataCombo dbcClassification 
         Height          =   288
         Left            =   1440
         TabIndex        =   4
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
         TabIndex        =   5
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
         TabIndex        =   3
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
         TabIndex        =   6
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
         TabIndex        =   25
         Top             =   1788
         Width           =   804
      End
      Begin VB.Label lblZipCode 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Zip Code:"
         Height          =   192
         Left            =   4740
         TabIndex        =   24
         Top             =   1788
         Width           =   696
      End
      Begin VB.Label lblLocalWebSite 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Local WebSite:"
         Height          =   192
         Left            =   240
         TabIndex        =   23
         Top             =   2808
         Width           =   1092
      End
      Begin VB.Label lblWebSite 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "WebSite:"
         Height          =   192
         Left            =   684
         TabIndex        =   22
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
         TabIndex        =   21
         Top             =   900
         Width           =   444
      End
      Begin VB.Label lblCommand 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Command:"
         Height          =   192
         Left            =   552
         TabIndex        =   20
         Top             =   1488
         Width           =   780
      End
      Begin VB.Label lblDesignation 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Designation:"
         Height          =   192
         Left            =   432
         TabIndex        =   19
         Top             =   300
         Width           =   900
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   192
         Left            =   852
         TabIndex        =   18
         Top             =   600
         Width           =   480
      End
      Begin VB.Label lblClassification 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Classification:"
         Height          =   192
         Left            =   348
         TabIndex        =   17
         Top             =   1200
         Width           =   984
      End
      Begin VB.Label lblCommissioned 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Commissioned:"
         Height          =   192
         Left            =   240
         TabIndex        =   15
         Top             =   2100
         Width           =   1116
      End
   End
   Begin VB.Frame fraShips 
      Height          =   3132
      Index           =   5
      Left            =   156
      TabIndex        =   33
      Top             =   780
      Width           =   7452
      Begin MSDataGridLib.DataGrid dgdImages 
         Height          =   2832
         Left            =   120
         TabIndex        =   66
         Top             =   180
         Width           =   7212
         _ExtentX        =   12721
         _ExtentY        =   4995
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraShips 
      Height          =   3132
      Index           =   4
      Left            =   156
      TabIndex        =   30
      Top             =   780
      Width           =   7452
      Begin RichTextLib.RichTextBox rtxtMoreHistory 
         Height          =   2892
         Left            =   60
         TabIndex        =   31
         Top             =   180
         Width           =   7332
         _ExtentX        =   12933
         _ExtentY        =   5101
         _Version        =   393217
         TextRTF         =   $"frmUSNShips.frx":0606
      End
   End
   Begin VB.Frame fraShips 
      Height          =   3132
      Index           =   3
      Left            =   156
      TabIndex        =   28
      Top             =   780
      Width           =   7452
      Begin RichTextLib.RichTextBox rtxtHistory 
         Height          =   2892
         Left            =   60
         TabIndex        =   29
         Top             =   180
         Width           =   7332
         _ExtentX        =   12933
         _ExtentY        =   5101
         _Version        =   393217
         TextRTF         =   $"frmUSNShips.frx":06EA
      End
   End
   Begin VB.Frame fraShips 
      Height          =   3132
      Index           =   2
      Left            =   156
      TabIndex        =   27
      Top             =   780
      Width           =   7452
      Begin RichTextLib.RichTextBox rtxtAircraft 
         Height          =   312
         Left            =   1260
         TabIndex        =   51
         Top             =   240
         Width           =   6072
         _ExtentX        =   10710
         _ExtentY        =   550
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmUSNShips.frx":07C9
      End
      Begin RichTextLib.RichTextBox rtxtMissiles 
         Height          =   312
         Left            =   1260
         TabIndex        =   53
         Top             =   600
         Width           =   6072
         _ExtentX        =   10710
         _ExtentY        =   550
         _Version        =   393217
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         TextRTF         =   $"frmUSNShips.frx":08A6
      End
      Begin RichTextLib.RichTextBox rtxtGuns 
         Height          =   312
         Left            =   1260
         TabIndex        =   55
         Top             =   960
         Width           =   6072
         _ExtentX        =   10710
         _ExtentY        =   550
         _Version        =   393217
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         TextRTF         =   $"frmUSNShips.frx":0983
      End
      Begin RichTextLib.RichTextBox rtxtASW 
         Height          =   312
         Left            =   1260
         TabIndex        =   57
         Top             =   1320
         Width           =   6072
         _ExtentX        =   10710
         _ExtentY        =   550
         _Version        =   393217
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         TextRTF         =   $"frmUSNShips.frx":0A5C
      End
      Begin RichTextLib.RichTextBox rtxtRadars 
         Height          =   312
         Left            =   1260
         TabIndex        =   59
         Top             =   1680
         Width           =   6072
         _ExtentX        =   10710
         _ExtentY        =   550
         _Version        =   393217
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         TextRTF         =   $"frmUSNShips.frx":0B3C
      End
      Begin RichTextLib.RichTextBox rtxtSonars 
         Height          =   312
         Left            =   1260
         TabIndex        =   61
         Top             =   2040
         Width           =   6072
         _ExtentX        =   10710
         _ExtentY        =   550
         _Version        =   393217
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         TextRTF         =   $"frmUSNShips.frx":0C17
      End
      Begin RichTextLib.RichTextBox rtxtFireControl 
         Height          =   312
         Left            =   1260
         TabIndex        =   63
         Top             =   2400
         Width           =   6072
         _ExtentX        =   10710
         _ExtentY        =   550
         _Version        =   393217
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         TextRTF         =   $"frmUSNShips.frx":0CF2
      End
      Begin RichTextLib.RichTextBox rtxtEW 
         Height          =   312
         Left            =   1740
         TabIndex        =   65
         Top             =   2760
         Width           =   5592
         _ExtentX        =   9864
         _ExtentY        =   550
         _Version        =   393217
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         TextRTF         =   $"frmUSNShips.frx":0DD3
      End
      Begin VB.Label lblEW 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Electronic Warfare:"
         Height          =   192
         Left            =   300
         TabIndex        =   64
         Top             =   2820
         Width           =   1356
      End
      Begin VB.Label lblFireControl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fire Control:"
         Height          =   192
         Left            =   360
         TabIndex        =   62
         Top             =   2460
         Width           =   852
      End
      Begin VB.Label lblSonars 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sonars:"
         Height          =   192
         Left            =   660
         TabIndex        =   60
         Top             =   2100
         Width           =   552
      End
      Begin VB.Label lblRadars 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Radars:"
         Height          =   192
         Left            =   636
         TabIndex        =   58
         Top             =   1740
         Width           =   576
      End
      Begin VB.Label lblASW 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ASW Weapons:"
         Height          =   192
         Left            =   60
         TabIndex        =   56
         Top             =   1380
         Width           =   1152
      End
      Begin VB.Label lblGuns 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Guns:"
         Height          =   192
         Left            =   804
         TabIndex        =   54
         Top             =   1020
         Width           =   408
      End
      Begin VB.Label lblMissiles 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Missiles:"
         Height          =   192
         Left            =   588
         TabIndex        =   52
         Top             =   660
         Width           =   624
      End
      Begin VB.Label lblAircraft 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Aircraft:"
         Height          =   192
         Left            =   684
         TabIndex        =   50
         Top             =   300
         Width           =   528
      End
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
         NumTabs         =   6
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
            Caption         =   "History..."
            Key             =   "History"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "More History..."
            Key             =   "MoreHistory"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
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
      TabIndex        =   68
      Top             =   4500
      Width           =   972
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   372
      Left            =   6696
      TabIndex        =   69
      Top             =   4500
      Width           =   972
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   11
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
            TextSave        =   "10:20 PM"
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
            Picture         =   "frmUSNShips.frx":0EAA
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNShips.frx":11C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNShips.frx":14EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNShips.frx":1816
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNShips.frx":3FCA
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNShips.frx":441E
            Key             =   "List"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNShips.frx":4EEA
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNShips.frx":533E
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNShips.frx":5E0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNShips.frx":6132
            Key             =   "Sort"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNShips.frx":6586
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNShips.frx":69DA
            Key             =   "SQL"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNShips.frx":6E2E
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
            Picture         =   "frmUSNShips.frx":7282
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNShips.frx":76D6
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNShips.frx":81A2
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNShips.frx":84BE
            Key             =   "List"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNShips.frx":8F8A
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNShips.frx":93DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNShips.frx":BB92
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSNShips.frx":BFE6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbAction 
      Align           =   1  'Align Top
      Height          =   288
      Left            =   0
      TabIndex        =   26
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
   Begin VB.Label lblA 
      AutoSize        =   -1  'True
      Caption         =   "A"
      Height          =   192
      Left            =   0
      TabIndex        =   67
      Top             =   0
      Visible         =   0   'False
      Width           =   108
   End
   Begin VB.Label lblID 
      AutoSize        =   -1  'True
      Caption         =   "lblID"
      Height          =   192
      Left            =   360
      TabIndex        =   13
      Top             =   4596
      Width           =   324
   End
   Begin VB.Label lblIDlabel 
      AutoSize        =   -1  'True
      Caption         =   "ID:"
      Height          =   192
      Left            =   96
      TabIndex        =   12
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
Dim WithEvents rsMain As ADODB.Recordset
Attribute rsMain.VB_VarHelpID = -1
Dim rsCommands As New ADODB.Recordset
Dim rsHomePorts As New ADODB.Recordset
Dim rsClasses As New ADODB.Recordset
Dim rsClassifications As New ADODB.Recordset
Dim rsImages As New ADODB.Recordset
Dim DBinfo As DataBaseInfo
Dim MouseY As Single
Dim MouseX As Single
Private SortDESC() As Boolean
Private Sub cmdCancel_Click()
    CancelCommand Me, rsMain
End Sub
Private Sub cmdOK_Click()
    If mode = modeAdd Or mode = modeModify Then
        If InStr(txtDesignation.Text, "-") Then
            rsMain("Number") = Mid(txtDesignation.Text, InStr(txtDesignation.Text, "-") + 1)
        End If
    End If
    OKCommand Me, rsMain
End Sub
Private Sub Form_Load()
    Set adoConn = New ADODB.Connection
    adoConn.Open "FileDSN=" & gstrFileDSN
    
    Set rsMain = New ADODB.Recordset
    rsMain.CursorLocation = adUseClient
    SQLmain = "select * from [Ships] order by Classification, Number"
    SQLfilter = vbNullString
    SQLkey = "Reference"
    rsMain.Open SQLmain, adoConn, adOpenKeyset, adLockBatchOptimistic
    DBcollection.Add "rsMain", rsMain
    rsMain.MoveFirst
    
    rsClassifications.CursorLocation = adUseClient
    rsClassifications.Open "select * from [Classification] order by Type", adoConn, adOpenStatic, adLockReadOnly
    DBcollection.Add "rsClassifications", rsClassifications
    
    rsCommands.CursorLocation = adUseClient
    rsCommands.Open "select distinct Command from [Ships] order by Command", adoConn, adOpenStatic, adLockReadOnly
    DBcollection.Add "rsCommands", rsCommands
    
    rsHomePorts.CursorLocation = adUseClient
    rsHomePorts.Open "select distinct HomePort from [Ships] order by HomePort", adoConn, adOpenStatic, adLockReadOnly
    DBcollection.Add "rsHomePorts", rsHomePorts
    
    rsClasses.CursorLocation = adUseClient
    rsClasses.Open "select * from [Class] order by Name", adoConn, adOpenStatic, adLockReadOnly
    DBcollection.Add "rsClasses", rsClasses
    
    Set adodcMain.Recordset = rsMain
    BindField lblID, "ID", rsMain

    BindFields
    
    Set tsShips.SelectedItem = tsShips.Tabs(1)
    ProtectFields Me
    mode = modeDisplay
    fTransaction = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Cancel = CloseConnection(Me)
End Sub
Private Sub mnuActionFilter_Click()
    FilterCommand Me, rsMain, SQLkey
End Sub
Private Sub mnuActionDelete_Click()
    DeleteCommand Me, rsMain
End Sub
Private Sub mnuActionList_Click()
    ListCommand Me, rsMain
End Sub
Private Sub mnuActionModify_Click()
    ModifyCommand Me
    
    Set tsShips.SelectedItem = tsShips.Tabs(1)
End Sub
Private Sub mnuActionNew_Click()
    NewCommand Me, rsMain

    txtCommissioned.Text = Format(Now(), "dddd, MMMM dd, yyyy")
    Set tsShips.SelectedItem = tsShips.Tabs(1)
End Sub
Private Sub mnuActionRefresh_Click()
    RefreshCommand rsMain, SQLkey
End Sub
Private Sub mnuActionReport_Click()
    ReportCommand Me, rsMain, App.Path & "\Reports\USN Ships.rpt"
End Sub
Private Sub mnuActionSQL_Click()
    SQLCommand "Ships"
End Sub
Private Sub rsMain_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Dim Caption As String
    
    If Not pRecordset.BOF And Not pRecordset.EOF Then Caption = "Reference #" & pRecordset.Bookmark & ": " & pRecordset("HullNumber") & " " & pRecordset("Name")
    UpdatePosition Me, Caption, pRecordset
    DefaultClassificationDesc
    DefaultClassDetails
    ResetImageGrid
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
'=================================================================================
Private Sub BindFields()
    'General
    BindField txtDesignation, "HullNumber", rsMain
    'BindField txtNumber, "Number", rsMain
    BindField txtName, "Name", rsMain
    BindField dbcClass, "ClassID", rsMain, rsClasses, "ID", "Name"
    BindField dbcClassification, "Classification", rsMain, rsClassifications, "Type", "Type"
    BindField txtClassDesc, "Description", rsClassifications
    BindField dbcCommand, "Command", rsMain, rsCommands, "Command", "Command"
    BindField dbcHomePort, "HomePort", rsMain, rsHomePorts, "HomePort", "HomePort"
    BindField txtZipCode, "Zip Code", rsMain
    BindField txtCommissioned, "Commissioned", rsMain
    BindField txtURL, "URL_Internet", rsMain
    BindField txtLocalURL, "URL_Local", rsMain
    
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
    
    'History...
    BindField rtxtHistory, "History", rsMain
    
    'More History...
    BindField rtxtMoreHistory, "More History", rsMain
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
Private Sub DefaultClassDetails()
    If rsClasses Is Nothing Then Exit Sub
    If rsClasses.State = adStateClosed Then Exit Sub
    If rsMain.BOF Or rsMain.EOF Then Exit Sub
    
    rsClasses.MoveFirst
    rsClasses.Find "ID=" & rsMain("ClassID").Value

    BindFields
    
    'Characteristics...
    If IsNull(rsMain("Displacement")) And Not IsNull(rsClasses("Displacement")) Then BindField rtxtDisplacement, "Displacement", rsClasses
    If IsNull(rsMain("Length")) And Not IsNull(rsClasses("Length")) Then BindField rtxtLength, "Length", rsClasses
    If IsNull(rsMain("Beam")) And Not IsNull(rsClasses("Beam")) Then BindField rtxtBeam, "Beam", rsClasses
    If IsNull(rsMain("Draft")) And Not IsNull(rsClasses("Draft")) Then BindField rtxtDraft, "Draft", rsClasses
    If IsNull(rsMain("Propulsion")) And Not IsNull(rsClasses("Propulsion")) Then BindField rtxtPropulsion, "Propulsion", rsClasses
    If IsNull(rsMain("Boilers")) And Not IsNull(rsClasses("Boilers")) Then BindField rtxtBoilers, "Boilers", rsClasses
    If IsNull(rsMain("Speed")) And Not IsNull(rsClasses("Speed")) Then BindField txtSpeed, "Speed", rsClasses
    If IsNull(rsMain("Manning")) And Not IsNull(rsClasses("Manning")) Then BindField rtxtManning, "Manning", rsClasses
    
    'Weapons...
    If IsNull(rsMain("Aircraft")) And Not IsNull(rsClasses("Aircraft")) Then BindField rtxtAircraft, "Aircraft", rsClasses
    If IsNull(rsMain("Missiles")) And Not IsNull(rsClasses("Missiles")) Then BindField rtxtMissiles, "Missiles", rsClasses
    If IsNull(rsMain("Guns")) And Not IsNull(rsClasses("Guns")) Then BindField rtxtGuns, "Guns", rsClasses
    If IsNull(rsMain("ASW Weapons")) And Not IsNull(rsClasses("ASW Weapons")) Then BindField rtxtASW, "ASW Weapons", rsClasses
    If IsNull(rsMain("Radars")) And Not IsNull(rsClasses("Radars")) Then BindField rtxtRadars, "Radars", rsClasses
    If IsNull(rsMain("Sonars")) And Not IsNull(rsClasses("Sonars")) Then BindField rtxtSonars, "Sonars", rsClasses
    If IsNull(rsMain("Fire Control")) And Not IsNull(rsClasses("Fire Control")) Then BindField rtxtFireControl, "Fire Control", rsClasses
    If IsNull(rsMain("EW")) And Not IsNull(rsClasses("EW")) Then BindField rtxtEW, "EW", rsClasses
    
    'History...
    If IsNull(rsMain("History")) And Not IsNull(rsClasses("Description")) Then BindField rtxtHistory, "Description", rsClasses
End Sub
Private Sub DefaultClassificationDesc()
    If rsClassifications Is Nothing Then Exit Sub
    If rsClassifications.State = adStateClosed Then Exit Sub
    If rsMain.BOF Or rsMain.EOF Then Exit Sub
    
    rsClassifications.MoveFirst
    rsClassifications.Find "ID=" & rsMain("ClassificationID").Value
End Sub
Private Sub dgdImages_DblClick()
    Dim col As Column
    Dim ColRight As Single
    Dim ColumnFormat As New StdDataFormat
    Dim DataWidth As Long
    Dim iCol As Integer
    Dim ResizeWindow As Single
    Dim rsTemp As ADODB.Recordset
    Dim WidestData As Long
    
    Me.MousePointer = vbHourglass
    
    ResizeWindow = 36
    For iCol = dgdImages.LeftCol To dgdImages.Columns.Count - 1
        Set col = dgdImages.Columns(iCol)
        If col.Visible And col.Width > 0 Then
            ColRight = col.Left + col.Width
            If MouseY <= col.Top And MouseX >= (ColRight - ResizeWindow) And MouseX <= (ColRight + ResizeWindow) Then
                dgdImages.ClearSelCols
                lblA.Caption = col.Caption
                WidestData = lblA.Width
                Set ColumnFormat = col.DataFormat
                If Not rsImages.BOF And Not rsImages.EOF Then
                    Set rsTemp = rsImages.Clone(adLockReadOnly)
                    rsTemp.MoveFirst
                    While Not rsTemp.EOF
                        If Not IsNull(rsTemp(col.Caption).Value) Then
                            If Not ColumnFormat Is Nothing Then
                                lblA.Caption = Format(rsTemp(col.Caption).Value, col.DataFormat.Format)
                            Else
                                lblA.Caption = CStr(rsTemp(col.Caption).Value)
                            End If
                            DataWidth = lblA.Width
                            If DataWidth > WidestData Then WidestData = DataWidth
                        End If
                        rsTemp.MoveNext
                    Wend
                    CloseRecordset rsTemp, True
                End If
                Set ColumnFormat = Nothing
                col.Width = WidestData + (4 * ResizeWindow)
                If col.Width > dgdImages.Width Then col.Width = col.Width - ResizeWindow
                GoTo ExitSub
            End If
        End If
    Next iCol
    
ExitSub:
    Me.MousePointer = vbDefault
End Sub
Private Sub dgdImages_HeadClick(ByVal ColIndex As Integer)
    If rsImages.BOF And rsImages.EOF Then Exit Sub
    If SortDESC(ColIndex) Then
        rsImages.Sort = dgdImages.Columns(ColIndex).Caption & " DESC"
    Else
        rsImages.Sort = dgdImages.Columns(ColIndex).Caption & " ASC"
    End If
    SortDESC(ColIndex) = Not SortDESC(ColIndex)
End Sub
Private Sub dgdImages_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X
    MouseY = Y
End Sub
Private Sub ResetImageGrid()
    If rsMain.BOF Or rsMain.EOF Then Exit Sub
    
    Set dgdImages.DataSource = Nothing
    CloseRecordset rsImages, False
    rsImages.CursorLocation = adUseClient
    rsImages.Open "select * from [Images] where TableName='Ships' And TableID=" & rsMain("ID"), adoConn, adOpenKeyset, adLockReadOnly
    Set dgdImages.DataSource = rsImages
    ReDim SortDESC(0 To dgdImages.Columns.Count - 1)
    dgdImages.Columns("TableID").Visible = False
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
    Dim tempDT As Date
    
    If Trim(txtCommissioned.Text) = vbNullString Then Exit Sub
    
    On Error Resume Next
    txtCommissioned.Text = Format(txtCommissioned.Text, "dddd, MMMM dd, yyyy")
    tempDT = CDate(Format(CDate(Mid(txtCommissioned.Text, InStr(txtCommissioned.Text, ", ") + 2)), "dd-MMM-yyyy"))
    If Err.Number <> 0 Then
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
Private Sub txtURL_GotFocus()
    TextSelected
End Sub
Private Sub txtZipCode_GotFocus()
    TextSelected
End Sub
