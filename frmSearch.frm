VERSION 5.00
Begin VB.Form frmSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search"
   ClientHeight    =   2556
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2556
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
    'Parse fields on screen and build a Search for the recordset...
    
    strSearch = vbNullString
    For i = 0 To RS.Fields.Count - 1
        If Len(txtFields(i).Text) > 0 Then
            If Mid(txtFields(i).Text, 1, 1) = "=" Or _
                Mid(txtFields(i).Text, 1, 2) = "<=" Or _
                Mid(txtFields(i).Text, 1, 2) = ">=" Or _
                Mid(txtFields(i).Text, 1, 1) = "<" Or _
                Mid(txtFields(i).Text, 1, 1) = ">" Or _
                UCase(Mid(txtFields(i).Text, 1, 5)) = "NOT " Then
                'Take what the user said literally...
                strSearch = strSearch & RS.Fields(i).Name & " " & txtFields(i).Text & " and "
            ElseIf UCase(Mid(txtFields(i).Text, 1, 4)) = "LIKE" Then
                'Force the "like" to uppercase for parsing later...
                strSearch = strSearch & RS.Fields(i).Name & " LIKE" & Mid(txtFields(i).Text, 5) & " and "
            ElseIf UCase(Mid(txtFields(i).Text, 1, 7)) = "BETWEEN" Then
                'Force the "between" to uppercase for parsing later...
                strSearch = strSearch & RS.Fields(i).Name & " BETWEEN" & Mid(txtFields(i).Text, 8) & " and "
            ElseIf UCase(Mid(txtFields(i).Text, 1, 2)) = "IN" Then
                'Force the "in" to uppercase for parsing later...
                strSearch = strSearch & RS.Fields(i).Name & " IN" & Mid(txtFields(i).Text, 3) & " and "
            Else
                Select Case RS.Fields(i).Type
                    Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar
                        strSearch = strSearch & RS.Fields(i).Name & " = '" & SQLQuote(txtFields(i).Text) & "' and "
                    Case adDate, adDBDate, adDBTime, adDBTimeStamp
                        strSearch = strSearch & RS.Fields(i).Name & " = #" & txtFields(i).Text & "# and "
                    Case Else
                        strSearch = strSearch & RS.Fields(i).Name & " = " & SQLQuote(txtFields(i).Text) & " and "
                End Select
            End If
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
Private Sub Form_Activate()
    Dim i As Integer
    Dim iTop As Integer
    Dim iLeft As Integer
    Dim NewHeight As Integer
    Dim NewTop As Integer
    Dim NewWidth As Integer
    Dim NewLeft As Integer
    
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
    
    For i = 0 To RS.Fields.Count - 1
        If i > 31 Then
            MsgBox "Warning: Only the first 32 fields can be used to Search your data.", vbInformation
            Exit For
        End If
        txtFields(i).Visible = True
        txtFields(i).Top = iTop
        
        lblFields(i).Alignment = lblFields(0).Alignment 'Right
        lblFields(i).Visible = True
        lblFields(i).Caption = RS.Fields(i).Name & ":"
        lblFields(i).Top = txtFields(i).Top + (txtFields(i).Height / 2) - (lblFields(i).Height / 2)
        lblFields(i).Left = iLeft
        
        txtFields(i).Left = lblFields(i).Left + lblFields(i).Width + hSpace
        txtFields(i).Width = Me.ScaleWidth - lblFields(i).Left - lblFields(i).Width - hSpace - hSpace - cmdApply.Width - hSpace
        txtFields(i).TabIndex = i
        
        iTop = iTop + txtFields(i).Height + vSpace
    Next i
    
    NewHeight = (txtFields(i - 1).Top + txtFields(i - 1).Height + vSpace) + Me.Height - Me.ScaleHeight
    If NewHeight > Me.Height Then
        Me.Top = Me.Top + ((NewHeight - Me.Height) / 2)
    Else
        Me.Top = Me.Top - ((NewHeight - Me.Height) / 2)
    End If
    Me.Height = NewHeight
    txtFields(0).SetFocus
End Sub
Private Sub Form_Load()
    Dim i As Integer
    
    For i = 0 To 31
        txtFields(i).Visible = False
        lblFields(i).Visible = False
    Next
End Sub
Private Sub txtFields_GotFocus(Index As Integer)
    TextSelected
End Sub
Private Sub txtFields_Validate(Index As Integer, Cancel As Boolean)
    txtFields(Index).Text = Trim(txtFields(Index).Text)
End Sub

