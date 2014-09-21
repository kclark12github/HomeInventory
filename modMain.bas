Attribute VB_Name = "modMain"
Option Explicit
Global Const fmtDate As String = "dd-MMM-yyyy hh:nn AMPM"
Global Const gstrProvider As String = "Microsoft.Jet.OLEDB.4.0"
'Global Const gstrProvider As String = "Microsoft.Jet.OLEDB.3.51"
Global Const gstrConnectionString As String = "DBQ=F:\Program Files\Home Inventory\Database\Ken's Stuff.mdb;DefaultDir=F:\Program Files\Home Inventory\Database;Driver={Microsoft Access Driver (*.mdb)};DriverId=281;FIL=MS Access;FILEDSN=C:\Program Files\Common Files\ODBC\Data Sources\Ken's Stuff.dsn;MaxBufferSize=2048;MaxScanRows=8;PageTimeout=5;SafeTransactions=0;Threads=3;UID=admin;UserCommitSync=Yes;"
'Global Const gstrConnectionString As String = "DBQ=F:\Program Files\Home Inventory\Database\Ken's Stuff.mdb;"
Global Const gstrRunTimeUserName As String = "admin"
Global Const gstrRunTimePassword As String = vbNullString
'Global Const gstrDefaultImage As String = "EarthRise.jpg"
Global Const gstrDefaultImage As String = "F14_102.jpg"
Global Const iMinWidth As Single = 2184
Global Const iMinHeight As Single = 1440
Global Const ResizeWindow As Single = 36
Global Const gfUseFilterMethod As Boolean = True
Public Enum ActionMode
    modeDisplay = 0
    modeAdd = 1
    modeModify = 2
    modeDelete = 3
End Enum

Public adoConn As ADODB.Connection
Public DBcollection As New DataBaseCollection
Public frmReport As Form
Public fTransaction As Boolean
Public gstrFileDSN As String
Public gstrDefaultImagePath As String
Public gstrImagePath As String
Public gstrODBCFileDSNDir As String
Public MinHeight As Integer
Public MinWidth As Integer
Public mode As ActionMode
Public rdcReport As CRAXDRT.Report
Public SQLmain As String
Public SQLfilter As String
Public SQLkey As String
Public Sub BindField(ctl As Control, DataField As String, DataSource As ADODB.Recordset, Optional RowSource As ADODB.Recordset, Optional BoundColumn As String, Optional ListField As String)
    Dim DateTimeFormat As StdDataFormat
    Select Case TypeName(ctl)
        Case "CheckBox", "Label", "PictureBox", "RichTextBox", "TextBox"
            Set ctl.DataSource = Nothing
            ctl.DataField = DataField
            Set ctl.DataSource = DataSource
            If DataSource(DataField).Type = adDate Then
                If ctl.DataFormat.Format = vbNullString Then
                    Set DateTimeFormat = New StdDataFormat
                    DateTimeFormat.Format = fmtDate
                    Set ctl.DataFormat = DateTimeFormat
                End If
            End If
        Case "DataCombo"
            Set ctl.DataSource = Nothing
            ctl.DataField = DataField
            Set ctl.DataSource = DataSource
            Set ctl.RowSource = Nothing
            ctl.BoundColumn = BoundColumn
            ctl.ListField = ListField
            Set ctl.RowSource = RowSource
    End Select

    Select Case TypeName(ctl)
        Case "TextBox"
            Select Case DataSource(DataField).Type
                Case adDate, adDBDate, adDBTime, adDBTimeStamp
                Case Else
                    ctl.MaxLength = DataSource(DataField).DefinedSize
            End Select
        Case "DataCombo"
    End Select
End Sub
'Public Sub BindFieldDAO(ctl As Control, DataField As String, DataSource As DAO.Recordset, Optional RowSource As DAO.Recordset, Optional BoundColumn As String, Optional ListField As String)
'    Dim DateTimeFormat As StdDataFormat
'    Select Case TypeName(ctl)
'        Case "CheckBox", "Label", "PictureBox", "RichTextBox", "TextBox"
'            Set ctl.DataSource = Nothing
'            ctl.DataField = DataField
'            'Set ctl.DataSource = DataSource
'            If DataSource(DataField).Type = adDate Then
'                If ctl.DataFormat.Format = vbNullString Then
'                    Set DateTimeFormat = New StdDataFormat
'                    DateTimeFormat.Format = fmtDate
'                    Set ctl.DataFormat = DateTimeFormat
'                End If
'            End If
'    End Select
'End Sub
Public Sub CancelCommand(frm As Form, RS As ADODB.Recordset)
    Select Case mode
        Case modeDisplay
            Unload frm
        Case modeAdd, modeModify
            RS.CancelUpdate
            If mode = modeAdd And Not RS.EOF Then RS.MoveLast
            adoConn.RollbackTrans
            fTransaction = False
            ProtectFields frm
            mode = modeDisplay
            frm.adodcMain.Enabled = True
            
            frm.mnuFile.Enabled = True
            frm.mnuRecords.Enabled = True
            frm.tbMain.Enabled = True
    End Select
End Sub
Public Function CloseConnection(frm As Form) As Integer
    Dim DBinfo As DataBaseInfo
    
    If fTransaction Then
        MsgBox "Please complete the current operation before closing the window.", vbExclamation, frm.Caption
        CloseConnection = 1
        Exit Function
    End If
    
    For Each DBinfo In DBcollection
        CloseRecordset DBinfo.Recordset, True
    Next
    DBcollection.Clear
    
    On Error Resume Next
    adoConn.Close
    If Err.Number = 3246 Then
        adoConn.RollbackTrans
        fTransaction = False
        adoConn.Close
    End If
    Set adoConn = Nothing
    CloseConnection = 0
End Function
Public Sub DeleteCommand(frm As Form, RS As ADODB.Recordset)
    mode = modeDelete
    If MsgBox("Are you sure you want to permanently delete this record...?", vbYesNo, frm.Caption) = vbYes Then
        RS.Delete
        RS.MoveNext
        If RS.EOF Then RS.MoveLast
    End If
    mode = modeDisplay
End Sub
Public Sub EstablishConnection(cn As ADODB.Connection)
    If Not cn Is Nothing Then
        If (cn.State And adStateOpen) = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
    Set cn = New ADODB.Connection
    'cn.IsolationLevel = adXactCursorStability
    'cn.mode = adModeShareDenyNone
    cn.CursorLocation = adUseClient
    'cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=F:\Program Files\Home Inventory\Database\Ken's Stuff.mdb;;"
    cn.Open "FileDSN=" & gstrFileDSN
End Sub
Public Sub FilterCommand(frm As Form, RS As ADODB.Recordset, ByVal Key As String)
    Dim FieldList As String
    Dim TableList As String
    Dim WhereClause As String
    Dim OrderByClause As String
    Dim SQLstatement As String
    
    Load frmFilter
    frmFilter.Caption = frm.Caption & " Filter"
    If frmMain.Width > frm.Width And frmMain.Height > frm.Height Then
        frmFilter.Top = frmMain.Top
        frmFilter.Left = frmMain.Left
        frmFilter.Width = frmMain.Width
        frmFilter.Height = frmMain.Height
    Else
        frmFilter.Top = frm.Top
        frmFilter.Left = frm.Left
        frmFilter.Width = frm.Width
        frmFilter.Height = frm.Height
    End If
    
    Set frmFilter.RS = RS
    frmFilter.strFilter = SQLfilter
    frmFilter.Show vbModal
    SQLfilter = frmFilter.strFilter
    Unload frmFilter
    If SQLfilter <> vbNullString Then
        'ParseSQLSelect RS.Source, FieldList, TableList, WhereClause, OrderByClause
        ParseSQLSelect SQLmain, FieldList, TableList, WhereClause, OrderByClause
        If WhereClause <> vbNullString Then
            WhereClause = WhereClause & " And " & SQLfilter
        Else
            WhereClause = SQLfilter
        End If
        SQLstatement = "Select " & FieldList & " From " & TableList
        If WhereClause <> vbNullString Then SQLstatement = SQLstatement & " Where " & WhereClause
        If OrderByClause <> vbNullString Then SQLstatement = SQLstatement & " Order By " & OrderByClause
        
        frm.sbStatus.Panels("Message").Text = "Filter: " & SQLfilter
    Else
        SQLstatement = SQLmain
        frm.sbStatus.Panels("Message").Text = vbNullString
    End If
    
    If gfUseFilterMethod Then
        RS.Filter = 0
        If SQLfilter <> vbNullString Then
            RS.Filter = SQLfilter
        Else
            RefreshCommand RS
        End If
    Else
        CloseRecordset RS, False
        RS.Open SQLstatement, adoConn, adOpenKeyset, adLockOptimistic
        'I may need to change this later, but I didn't want to go through
        'all the screens' List commands and add the argument to ListCommand()
        '(i.e. frmList supports a Filter command, but hasn't been passed a Key)...
        RefreshCommand RS, Key
    End If
End Sub
Public Sub ListCommand(frm As Form, RS As ADODB.Recordset, Optional AllowUpdate As Boolean = True)
    Dim vRS As ADODB.Recordset
    
    Load frmList
    frmList.Caption = frm.Caption & " List"
    If frmMain.Width > frm.Width And frmMain.Height > frm.Height Then
        frmList.Top = frmMain.Top
        frmList.Left = frmMain.Left
        frmList.Width = frmMain.Width
        frmList.Height = frmMain.Height
    Else
        frmList.Top = frm.Top
        frmList.Left = frm.Left
        frmList.Width = frm.Width
        frmList.Height = frm.Height
    End If
    
    If AllowUpdate Then
        Set frmList.rsList = RS
        adoConn.BeginTrans
        fTransaction = True
    Else
        If Not MakeVirtualRecordset(adoConn, RS, vRS, "Junk") Then
            MsgBox "MakeVirtualRecordset failed.", vbExclamation, frm.Caption
            Exit Sub
        End If
        Set frmList.vrsList = vRS
    End If
    
    frmList.Show vbModal
    frm.sbStatus.Panels("Message").Text = vbNullString
    If RS.Filter <> vbNullString And RS.Filter <> 0 Then
        frm.sbStatus.Panels("Message").Text = "Filter: " & RS.Filter
    End If
        
    If AllowUpdate Then
        adoConn.CommitTrans
        fTransaction = False
    Else
        CloseRecordset vRS, True
    End If
End Sub
Public Sub ModifyCommand(frm As Form)
    mode = modeModify
    OpenFields frm
    frm.mnuFile.Enabled = False
    frm.mnuRecords.Enabled = False
    frm.tbMain.Enabled = False
    frm.adodcMain.Enabled = False
    adoConn.BeginTrans
    fTransaction = True
End Sub
Public Sub NewCommand(frm As Form, RS As ADODB.Recordset)
    mode = modeAdd
    OpenFields frm
    frm.mnuFile.Enabled = False
    frm.mnuRecords.Enabled = False
    frm.tbMain.Enabled = False
    frm.adodcMain.Enabled = False
    RS.AddNew
    adoConn.BeginTrans
    fTransaction = True
End Sub
Public Sub OKCommand(frm As Form, RS As ADODB.Recordset)
    Select Case mode
        Case modeDisplay
            Unload frm
        Case modeAdd, modeModify
            'Why we need to do this is buggy...
            'rsMain("Manufacturer") = dbcManufacturer.BoundText
            'rsMain("Catalog") = dbcCatalog.BoundText
            
            'Ignore errors because more than likely they're caused by exceeding
            'a field length. This is handled for TextBoxes, but cannot be easily
            'done for DataCombo controls (no .MaxLength property)...
            On Error Resume Next
            If RS.LockType = adLockBatchOptimistic Then
                RS.UpdateBatch
            Else
                RS.Update
            End If
            adoConn.CommitTrans
            fTransaction = False
            ProtectFields frm
            mode = modeDisplay
            frm.adodcMain.Enabled = True
            
            RefreshCommand RS, SQLkey
            
            frm.mnuFile.Enabled = True
            frm.mnuRecords.Enabled = True
            frm.tbMain.Enabled = True
    End Select
End Sub
Public Sub OpenFields(pForm As Form)
    Dim ctl As Control
    For Each ctl In pForm.Controls
        Select Case TypeName(ctl)
            Case "ComboBox", "DataCombo", "DataGrid", "RichTextBox", "TextBox"
                'ctl.Locked = False
                ctl.Enabled = True
                ctl.BackColor = vbWindowBackground
            Case "CheckBox", "PictureBox"
                ctl.Enabled = True
        End Select
    Next ctl
    pForm.sbStatus.Panels("Status").Text = "Edit Mode"
    pForm.cmdCancel.Caption = "Cancel"
    pForm.cmdOK.Visible = True
End Sub
Public Sub ProtectFields(pForm As Form)
    Dim ctl As Control
    For Each ctl In pForm.Controls
        Select Case TypeName(ctl)
            Case "ComboBox", "DataCombo", "DataGrid", "RichTextBox", "TextBox"
                'ctl.Locked = True
                ctl.Enabled = False
                ctl.BackColor = vbButtonFace
            Case "CheckBox", "PictureBox"
                ctl.Enabled = False
        End Select
    Next ctl

    pForm.sbStatus.Panels("Status").Text = ""
    pForm.cmdCancel.Caption = "&Exit"
    pForm.cmdOK.Visible = False
End Sub
Public Sub RefreshCommand(RS As ADODB.Recordset, Optional Key As Variant)
    Dim SaveBookmark As String
    Dim DBinfo As DataBaseInfo
    
    On Error Resume Next
    If IsMissing(Key) Then
        SaveBookmark = RS(0)
    Else
        SaveBookmark = RS(Key)
    End If
    RS.Requery
    If IsMissing(Key) Then
        RS.Find RS(0).Name & "='" & SQLQuote(SaveBookmark) & "'"
    Else
        RS.Find Key & "='" & SQLQuote(SaveBookmark) & "'"
    End If
    
    For Each DBinfo In DBcollection
        If Not (DBinfo.Recordset Is RS) Then DBinfo.Recordset.Requery
    Next
End Sub
Public Sub ReportCommand(frm As Form, RS As ADODB.Recordset, ByVal ReportPath As String)
    Dim scrApplication As New CRAXDRT.Application
    Dim Report As New CRAXDRT.Report
    Dim vRS As ADODB.Recordset
    
    MakeVirtualRecordset adoConn, RS, vRS
    
    Load frmViewReport
    frmViewReport.Caption = frm.Caption & " Report"
    If frmMain.Width > frm.Width And frmMain.Height > frm.Height Then
        frmViewReport.Top = frmMain.Top
        frmViewReport.Left = frmMain.Left
        frmViewReport.Width = frmMain.Width
        frmViewReport.Height = frmMain.Height
    Else
        frmViewReport.Top = frm.Top
        frmViewReport.Left = frm.Left
        frmViewReport.Width = frm.Width
        frmViewReport.Height = frm.Height
    End If
    frmViewReport.WindowState = vbMaximized
    
    Set Report = scrApplication.OpenReport(ReportPath, crOpenReportByTempCopy)
    Report.Database.SetDataSource vRS, 3, 1
    Report.ReadRecords
    
    frmViewReport.scrViewer.ReportSource = Report
    frmViewReport.Show vbModal
    
    Set scrApplication = Nothing
    Set Report = Nothing
    vRS.Close
    Set vRS = Nothing
End Sub
Public Sub SearchCommand(frm As Form, RS As ADODB.Recordset, ByVal Key As String)
    Dim FieldList As String
    Dim TableList As String
    Dim WhereClause As String
    Dim OrderByClause As String
    Dim SQLstatement As String
    
    Load frmSearch
    frmSearch.Caption = frm.Caption & " Search"
    If frmMain.Width > frm.Width And frmMain.Height > frm.Height Then
        frmSearch.Top = frmMain.Top
        frmSearch.Left = frmMain.Left
        frmSearch.Width = frmMain.Width
        frmSearch.Height = frmMain.Height
    Else
        frmSearch.Top = frm.Top
        frmSearch.Left = frm.Left
        frmSearch.Width = frm.Width
        frmSearch.Height = frm.Height
    End If
    
    Set frmSearch.RS = RS
    frmSearch.Show vbModal
    'RefreshCommand RS, Key
End Sub
Public Sub SQLCommand(ByVal TableName As String)
    Load frmSQL
    Set frmSQL.cnSQL = adoConn
    If ParsePath(gstrFileDSN, DrvDirNoSlash) = gstrODBCFileDSNDir Then
        frmSQL.txtDatabase.Text = ParsePath(gstrFileDSN, FileNameBase)
    Else
        frmSQL.txtDatabase.Text = ParsePath(gstrFileDSN, DrvDirFileNameBase)
    End If
    frmSQL.dbcTables.BoundText = TableName
    frmSQL.Show vbModal
End Sub
Public Sub UpdatePosition(frm As Form, ByVal Caption As String, RS As ADODB.Recordset)
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    If RS.BOF And RS.EOF Then
        Caption = "No Records"
    ElseIf RS.EOF Then
        Caption = "EOF"
    ElseIf RS.BOF Then
        Caption = "BOF"
    Else
        i = InStr(Caption, "&")
        If i > 0 Then Caption = Left(Caption, i) & "&" & Mid(Caption, i + 1)
        If RS.Filter <> vbNullString And RS.Filter <> 0 Then
            frm.sbStatus.Panels("Message").Text = "Filter: " & RS.Filter
        End If
        frm.sbStatus.Panels("Position").Text = "Record " & RS.Bookmark & " of " & RS.RecordCount
    End If
    
    frm.adodcMain.Caption = Caption
    Exit Sub

ErrorHandler:
    MsgBox Err.Description & " (Error " & Err.Number & ")", vbExclamation, frm.Caption
    Resume Next
End Sub
Public Sub dbcValidate(fld As ADODB.Field, ctl As DataCombo)
    If Len(ctl.Text) > fld.DefinedSize Then ctl.Text = Mid(ctl.Text, 1, fld.DefinedSize)
End Sub
