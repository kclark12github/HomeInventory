Attribute VB_Name = "SendRecycle"
Option Explicit

Private Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String
End Type

Private Declare Function SHFileOperation Lib "shell32.dll" Alias _
"SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Private Const FO_DELETE = &H3
Private Const FOF_ALLOWUNDO = &H40
Private Const FOF_NOCONFIRMATION = &H10
Private Const FOF_SILENT = &H4

Public Function SendToRecycleBin(ByVal FileName As String, _
 Progress As Boolean, Confirm As Boolean) As Boolean
    Dim SHFileOp As SHFILEOPSTRUCT
    Dim RetVal As Long

    With SHFileOp
        .wFunc = FO_DELETE
        .pFrom = FileName
        .fFlags = FOF_ALLOWUNDO
         If Not Progress And Confirm Then
              'Hide progress dialog box
             .fFlags = FOF_ALLOWUNDO Or FOF_SILENT
         ElseIf Not Confirm Then
             'Hide confirm delete dialog box. This also hides
             'the progress dialog box
             .fFlags = FOF_ALLOWUNDO Or FOF_NOCONFIRMATION
         Else
             .fFlags = FOF_ALLOWUNDO
         End If
    End With
    'Send the file to recycle bin
    RetVal = SHFileOperation(SHFileOp)
    'Check if operation was a success
    If RetVal > 0 Then
    'File does not exist
        SendToRecycleBin = False
    ElseIf SHFileOp.fAnyOperationsAborted Then
        'Operation aborted, file not sent to recycle bin
        SendToRecycleBin = False
    Else
        SendToRecycleBin = True
    End If
End Function
