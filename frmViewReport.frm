VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.2#0"; "CRVIEWER.DLL"
Begin VB.Form frmViewReport 
   Caption         =   "View Report"
   ClientHeight    =   2556
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   2556
   ScaleWidth      =   3744
   StartUpPosition =   2  'CenterScreen
   Begin CRVIEWERLibCtl.CRViewer scrViewer 
      Height          =   2592
      Left            =   420
      TabIndex        =   0
      Top             =   0
      Width           =   2892
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
   End
End
Attribute VB_Name = "frmViewReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Public rdcReport As CRAXDRT.Report
Private fActivated As Boolean
Private Sub DisplayViewer()
    Screen.MousePointer = vbHourglass
    'scrViewer.ReportSource = rdcReport
    'frmMain.rdcReport.ReadRecords
    scrViewer.ViewReport
    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Activate()
    If Not fActivated Then
        fActivated = True
        DisplayViewer
    End If
End Sub
Private Sub Form_Load()
    fActivated = False
    Me.Icon = Forms(Forms.Count - 2).Icon
End Sub
Private Sub Form_Paint()
    If Not fActivated Then
        fActivated = True
        DisplayViewer
    End If
End Sub
Private Sub Form_Resize()
    scrViewer.Top = 0
    scrViewer.Left = 0
    scrViewer.Height = ScaleHeight
    scrViewer.Width = ScaleWidth
End Sub
