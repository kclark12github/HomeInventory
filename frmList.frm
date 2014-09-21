VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmList 
   Caption         =   "List"
   ClientHeight    =   2556
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   2556
   ScaleWidth      =   3744
   StartUpPosition =   1  'CenterOwner
   Begin MSDataGridLib.DataGrid dgdList 
      Height          =   612
      Left            =   1260
      TabIndex        =   0
      Top             =   960
      Width           =   1212
      _ExtentX        =   2138
      _ExtentY        =   1080
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
Attribute VB_Name = "frmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mnuList As Menu
Public rsList As ADODB.Recordset
Private Key As String
Private SortDESC() As Boolean
Private Sub dgdList_HeadClick(ByVal ColIndex As Integer)
    If SortDESC(ColIndex) Then
        rsList.Sort = dgdList.Columns(ColIndex).Caption & " DESC"
    Else
        rsList.Sort = dgdList.Columns(ColIndex).Caption & " ASC"
    End If
    SortDESC(ColIndex) = Not SortDESC(ColIndex)
End Sub
Private Sub dgdList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbKeyRButton Then PopupMenu mnuList
End Sub
Private Sub Form_Activate()
    Dim i As Integer
    
    ReDim SortDESC(0 To dgdList.Columns.Count - 1)
    
    'Get the column settings for the display...
    For i = 0 To dgdList.Columns.Count - 1
        dgdList.Columns(i).Width = GetSetting(App.ProductName, Me.Caption & " Settings", dgdList.Columns(i).Caption & " Width", dgdList.Columns(i).Width)
    Next
End Sub
Private Sub Form_Load()
    dgdList.Top = 0
    dgdList.Left = 0
    dgdList.Width = Me.ScaleWidth
    dgdList.Height = Me.ScaleHeight
End Sub
Private Sub Form_Resize()
    If Me.Width < frmMain.MinWidth Then Me.Width = frmMain.MinWidth
    If Me.Height < frmMain.MinHeight Then Me.Height = frmMain.MinHeight
    dgdList.Width = Me.ScaleWidth
    dgdList.Height = Me.ScaleHeight
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    
    rsList.UpdateBatch
    
    'Save the column settings for the next display...
    For i = 0 To dgdList.Columns.Count - 1
        SaveSetting App.ProductName, Me.Caption & " Settings", dgdList.Columns(i).Caption & " Width", dgdList.Columns(i).Width
    Next
End Sub
