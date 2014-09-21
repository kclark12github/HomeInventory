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
   StartUpPosition =   1  'CenterOwner
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
         BeginProperty Font 
            Name            =   "r_ansi"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
Dim InitialWidth As Double
Dim InitialHeight As Double
Dim RecordsAffected As Long
Public cnSQL As ADODB.Connection
Private Sub cmdClose_Click()
    Set cnSQL = Nothing
    Unload Me
End Sub
Private Sub cmdExecute_Click()
    Dim adoRS As ADODB.Recordset
    Dim adoError As ADODB.Error
    Dim fld As ADODB.Field
    Dim ErrorCount As Long
    Dim strOutput As String
    
    On Error Resume Next
    txtResults.Text = ""
    Set adoRS = cnSQL.Execute(txtSQL.Text, RecordsAffected)
    For Each adoError In cnSQL.Errors
        If Trim(adoError.Description) <> vbNullString Then
            txtResults.Text = txtResults.Text & adoError.Description & "(" & Hex(adoError.Number) & ")" & vbCrLf
            txtResults.Text = txtResults.Text & vbTab & "Source: " & adoError.Source & vbCrLf & _
                vbTab & "SQL State: " & adoError.SQLState & vbCrLf & _
                vbTab & "Native Error: " & adoError.NativeError & vbCrLf
            ErrorCount = ErrorCount + 1
        End If
    Next
    If ErrorCount > 0 Then Exit Sub
    
    If UCase(Mid(txtSQL.Text, 1, 6)) = "SELECT" Then
        'Print Column Headers...
        strOutput = vbNullString
        For Each fld In adoRS.Fields
            strOutput = strOutput & fld.Name
            Select Case fld.Type
                Case adVarChar, adChar
                    strOutput = strOutput & String(fld.DefinedSize - Len(fld.Name) + 1, " ")
                Case adInteger, adCurrency
                    strOutput = strOutput & String(10 - Len(fld.Name) + 1, " ")
            End Select
        Next
        'txtResults.Text = strOutput
        Debug.Print strOutput
        strOutput = String(Len(strOutput), "-")
        'txtResults.Text = txtResults.Text & vbCrLf & strOutput
        Debug.Print strOutput
        
        'Now print a row for each record...
        adoRS.MoveFirst
        While Not adoRS.EOF
            strOutput = vbNullString
            For Each fld In adoRS.Fields
                Select Case fld.Type
                    Case adVarChar, adChar
                        strOutput = strOutput & fld.Value & String(fld.DefinedSize - Len(fld.Value) + 1, " ")
                    Case adCurrency
                        strOutput = strOutput & Format(fld.Value, "Currency") & String(10 - Len(Format(fld.Value, "Currency")) + 1, " ")
                    Case adInteger
                        strOutput = strOutput & fld.Value & String(10 - Len(fld.Value) + 1, " ")
                End Select
            Next
            'txtResults.Text = txtResults.Text & strOutput & vbCrLf
            Debug.Print strOutput
            adoRS.MoveNext
        Wend
    End If
    
    If RecordsAffected = 1 Then
        txtResults.Text = RecordsAffected & " record affected"
    Else
        txtResults.Text = RecordsAffected & " records affected"
    End If
End Sub
Private Sub Form_Activate()
    InitialWidth = Me.Width
    InitialHeight = Me.Height
End Sub
Private Sub Form_Resize()
    Dim NewFrameWidth As Double
    
    NewFrameWidth = Me.ScaleWidth - cmdExecute.Width - (3 * MarginTwips)
    
    If Me.Width < InitialWidth Or Me.Height < InitialHeight Then
        'Debug.Print "Initial Width x Height: " & InitialWidth & "x" & InitialHeight
        'Debug.Print "Me.Width x Height: " & Me.Width & "x" & Me.Height
        Me.Move Me.Left, Me.Top, InitialWidth, InitialHeight
        Exit Sub
    End If
    cmdExecute.Move Me.ScaleWidth - cmdExecute.Width - MarginTwips, MarginTwips
    cmdClose.Move Me.ScaleWidth - cmdClose.Width - MarginTwips, cmdExecute.Top + cmdExecute.Height + MarginTwips
    frameSQL.Move MarginTwips, MarginTwips, NewFrameWidth, Me.ScaleHeight / 3
    txtSQL.Move MarginTwips, 3 * MarginTwips, frameSQL.Width - (2 * MarginTwips), frameSQL.Height - (4 * MarginTwips)
    frameResults.Move MarginTwips, frameSQL.Top + frameSQL.Height + MarginTwips, NewFrameWidth, Me.ScaleHeight - frameSQL.Height - (3 * MarginTwips)
    txtResults.Move MarginTwips, 3 * MarginTwips, frameResults.Width - (2 * MarginTwips), frameResults.Height - (4 * MarginTwips)
End Sub
