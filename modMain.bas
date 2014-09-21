Attribute VB_Name = "modMain"
Option Explicit
'Global Const fmtDate As String = "dd-MMM-yyyy hh:nn AMPM"
'Global Const gstrProvider As String = "Microsoft.Jet.OLEDB.4.0"
Global Const gstrProvider As String = "Microsoft.Jet.OLEDB.3.51"
Global Const gstrConnectionString As String = "DBQ=C:\My Documents\Home Inventory\Database\Ken's Stuff.mdb;DefaultDir=C:\My Documents\Home Inventory\Database;Driver={Microsoft Access Driver (*.mdb)};DriverId=281;FIL=MS Access;FILEDSN=C:\Program Files\Common Files\ODBC\Data Sources\Ken's Stuff.dsn;MaxBufferSize=2048;MaxScanRows=8;PageTimeout=5;SafeTransactions=0;Threads=3;UID=admin;UserCommitSync=Yes;"
'Global Const gstrConnectionString As String = "DBQ=C:\My Documents\Home Inventory\Database\Ken's Stuff.mdb;"
Global Const gstrRunTimeUserName As String = "admin"
Global Const gstrRunTimePassword As String = vbNullString
'Global Const gstrDefaultImage As String = "EarthRise.jpg"
Global Const gstrDefaultImage As String = "F14_102.jpg"
Global Const iMinWidth As Single = 2184
Global Const iMinHeight As Single = 1440
Global Const ResizeWindow As Single = 36

Private Const LOCALE_SSHORTDATE = &H1F
Private Const WM_SETTINGCHANGE = &H1A
'same as the old WM_WININICHANGE
Private Const HWND_BROADCAST = &HFFFF&

'Private Declare Function SetLocaleInfo Lib "kernel32" Alias _
    "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Boolean
Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias _
    "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

Public fmtDate As String
Public fmtShortDate As String
Public fmtLongDate As String
Public fmtFullDateTime As String

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
Public gfUseFilterMethod As Boolean
Public gstrFileDSN As String
Public gstrDefaultImagePath As String
Public gstrImagePath As String
Public gstrODBCFileDSNDir As String
Public MinHeight As Integer
Public MinWidth As Integer
Public mode As ActionMode
'Public rdcReport As CRAXDRT.Report
Public SQLmain As String
Public SQLfilter As String
Public SQLkey As String
Public Sub BindField(ctl As Control, DataField As String, DataSource As ADODB.Recordset, Caption As String, Optional RowSource As ADODB.Recordset, Optional BoundColumn As String, Optional ListField As String)
    Dim DateTimeFormat As StdDataFormat
    
    Call Trace(trcEnter, "BindField(""" & ctl.Name & """, """ & DataField & """, DataSource, RowSource, """ & BoundColumn & """, """ & ListField & """)")
    Select Case TypeName(ctl)
        Case "CheckBox", "Label", "PictureBox", "RichTextBox", "TextBox"
            ctl.Tag = Caption
            Set ctl.DataSource = Nothing
            ctl.DataField = DataField
            Set ctl.DataSource = DataSource
            If DataSource(DataField).Type = adDate Then
                If ctl.DataFormat.Format = vbNullString Then
                    ctl.DataFormat.Format = fmtDate
'                    Set DateTimeFormat = New StdDataFormat
'                    DateTimeFormat.Format = fmtDate
'                    Set ctl.DataFormat = DateTimeFormat
                End If
            End If
        Case "DataCombo"
            ctl.Tag = Caption
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
                Case adLongVarChar
                Case Else
                    ctl.MaxLength = DataSource(DataField).DefinedSize
            End Select
        Case "DataCombo"
    End Select
    Call Trace(trcExit, "BindField")
End Sub
Public Sub CancelCommand(frm As Form, RS As ADODB.Recordset)
    Call Trace(trcEnter, "CancelCommand(""" & frm.Name & """, RS)")
    Select Case mode
        Case modeDisplay
            Unload frm
        Case modeAdd, modeModify
            Call Trace(trcBody, "RS.CancelUpdate")
            RS.CancelUpdate
            If mode = modeAdd And Not RS.EOF Then
                Call Trace(trcBody, "RS.MoveLast")
                RS.MoveLast
            End If
            Call Trace(trcBody, "adoConn.RollbackTrans")
            adoConn.RollbackTrans
            fTransaction = False
            ProtectFields frm
            mode = modeDisplay
            frm.adodcMain.Enabled = True
            
            frm.mnuFile.Enabled = True
            frm.mnuRecords.Enabled = True
            frm.tbMain.Enabled = True
    End Select
    Call Trace(trcExit, "CancelCommand")
End Sub
Public Function CloseConnection(frm As Form) As Integer
    Dim DBinfo As DataBaseInfo
    
    Call Trace(trcEnter, "CloseConnection(""" & frm.Name & """)")
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
    Call Trace(trcBody, "adoConn.Close")
    adoConn.Close
    If Err.Number = 3246 Then
        Call Trace(trcBody, "adoConn.RollbackTrans")
        adoConn.RollbackTrans
        fTransaction = False
        Call Trace(trcBody, "adoConn.Close")
        adoConn.Close
    End If
    Set adoConn = Nothing
    CloseConnection = 0
    Call Trace(trcExit, "CloseConnection")
End Function
Public Sub DeleteCommand(frm As Form, RS As ADODB.Recordset)
    Call Trace(trcEnter, "DeleteCommand(""" & frm.Name & """, RS)")
    mode = modeDelete
    If MsgBox("Are you sure you want to permanently delete this record...?", vbYesNo, frm.Caption) = vbYes Then
        Call Trace(trcBody, "RS.Delete")
        RS.Delete
        Call Trace(trcBody, "RS.MoveNext")
        RS.MoveNext
        If RS.EOF Then
            Call Trace(trcBody, "RS.MoveLast")
            RS.MoveLast
        End If
    End If
    mode = modeDisplay
    Call Trace(trcExit, "DeleteCommand")
End Sub
Public Sub EstablishConnection(cn As ADODB.Connection)
    Call Trace(trcEnter, "EstablishConnection")
    If Not cn Is Nothing Then
        If (cn.State And adStateOpen) = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
    Set cn = New ADODB.Connection
    'cn.IsolationLevel = adXactCursorStability
    'cn.mode = adModeShareDenyNone
    cn.CursorLocation = adUseClient
    'cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=c:\My Documents\Home Inventory\Database\Ken's Stuff.mdb;;"
    'cn.Open "Provider=MSDASQL;FileDSN=" & gstrFileDSN
    cn.Open "FileDSN=" & gstrFileDSN
    Call Trace(trcExit, "EstablishConnection - """ & cn.ConnectionString & """")
End Sub
Public Sub FilterCommand(frm As Form, RS As ADODB.Recordset, ByVal Key As String)
    Dim FieldList As String
    Dim TableList As String
    Dim WhereClause As String
    Dim OrderByClause As String
    Dim SQLstatement As String
    
    Call Trace(trcEnter, "FilterCommand(""" & frm.Name & """, RS, """ & Key & """)")
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
        Call Trace(trcBody, "RS.Filter = 0")
        RS.Filter = 0
        If SQLfilter <> vbNullString Then
            Call Trace(trcBody, "RS.Filter = """ & SQLfilter & """")
            RS.Filter = SQLfilter
        Else
            RefreshCommand RS
        End If
    Else
        CloseRecordset RS, False
        Call Trace(trcBody, "RS.Open """ & SQLstatement & """, adoConn, adOpenKeyset, adLockOptimistic)")
        RS.Open SQLstatement, adoConn, adOpenKeyset, adLockOptimistic

        'I may need to change this later, but I didn't want to go through
        'all the screens' List commands and add the argument to ListCommand()
        '(i.e. frmList supports a Filter command, but hasn't been passed a Key)...
        RefreshCommand RS, Key
    End If
    Call Trace(trcExit, "FilterCommand")
End Sub
Public Function GetRegionalShortDateFormat() As String
    Dim dwLCID As Long
    Dim dataLen As Integer
    Dim Buffer As String * 100
    
    Call Trace(trcEnter, "GetRegionalShortDateFormat")
    dwLCID = GetSystemDefaultLCID()
    dataLen = GetLocaleInfo(dwLCID, LOCALE_SSHORTDATE, Buffer, 100)
    GetRegionalShortDateFormat = Left$(Buffer, dataLen - 1)
    Call Trace(trcExit, "GetRegionalShortDateFormat = """ & GetRegionalShortDateFormat & """")
End Function
Public Sub ListCommand(frm As Form, RS As ADODB.Recordset, Optional AllowUpdate As Boolean = True)
    Dim vRS As ADODB.Recordset
    
    Call Trace(trcEnter, "ListCommand(""" & frm.Name & """, RS, " & AllowUpdate & ")")
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
        Call Trace(trcBody, "adoConn.BeginTrans")
        adoConn.BeginTrans
        fTransaction = True
        frmList.dgdList.BackColor = vbWindowBackground
    Else
        If Not MakeVirtualRecordset(adoConn, RS, vRS, "Junk") Then
            MsgBox "MakeVirtualRecordset failed.", vbExclamation, frm.Caption
            Exit Sub
        End If
        Set frmList.vrsList = vRS
        frmList.dgdList.BackColor = vbButtonFace
    End If
    
    frmList.Show vbModal
    frm.sbStatus.Panels("Message").Text = vbNullString
    If SQLfilter <> vbNullString Then
        frm.sbStatus.Panels("Message").Text = "Filter: " & SQLfilter
    End If
        
    If AllowUpdate Then
        Call Trace(trcBody, "adoConn.CommitTrans")
        adoConn.CommitTrans
        fTransaction = False
    Else
        CloseRecordset vRS, True
    End If
    Call Trace(trcExit, "ListCommand")
End Sub
Public Sub ModifyCommand(frm As Form)
    Call Trace(trcEnter, "ModifyCommand(""" & frm.Name & """)")
    mode = modeModify
    OpenFields frm
    frm.mnuFile.Enabled = False
    frm.mnuRecords.Enabled = False
    frm.tbMain.Enabled = False
    frm.adodcMain.Enabled = False
    Call Trace(trcBody, "adoConn.BeginTrans")
    adoConn.BeginTrans
    fTransaction = True
    Call Trace(trcExit, "ModifyCommand")
End Sub
Public Sub NewCommand(frm As Form, RS As ADODB.Recordset)
    Call Trace(trcEnter, "NewCommand(""" & frm.Name & """, RS)")
    mode = modeAdd
    OpenFields frm
    frm.mnuFile.Enabled = False
    frm.mnuRecords.Enabled = False
    frm.tbMain.Enabled = False
    frm.adodcMain.Enabled = False
    Call Trace(trcBody, "RS.AddNew")
    RS.AddNew
    Call Trace(trcBody, "adoConn.BeginTrans")
    adoConn.BeginTrans
    fTransaction = True
    Call Trace(trcExit, "NewCommand")
End Sub
Public Sub OKCommand(frm As Form, RS As ADODB.Recordset)
    Call Trace(trcEnter, "OKCommand(""" & frm.Name & """, RS)")
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
                Call Trace(trcBody, "RS.UpdateBatch")
                RS.UpdateBatch
            Else
                Call Trace(trcBody, "RS.Update")
                RS.Update
            End If
            Call Trace(trcBody, "adoConn.CommitTrans")
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
    Call Trace(trcExit, "OKCommand")
End Sub
Public Sub OpenFields(pForm As Form)
    Dim ctl As Control
    
    Call Trace(trcEnter, "OpenFields(""" & pForm.Name & """)")
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
    Call Trace(trcExit, "OpenFields")
End Sub
Public Sub ProtectFields(pForm As Form)
    Dim ctl As Control
    
    Call Trace(trcEnter, "ProtectFields(""" & pForm.Name & """)")
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
    Call Trace(trcExit, "ProtectFields")
End Sub
Public Sub RefreshCommand(RS As ADODB.Recordset, Optional Key As Variant)
    Dim SaveBookmark As String
    Dim DBinfo As DataBaseInfo
    
    Call Trace(trcEnter, "RefreshCommand(RS, """ & Key & """)")
    On Error Resume Next
    If IsMissing(Key) Or IsNull(RS(Key)) Then
        SaveBookmark = RS(0)
    Else
        SaveBookmark = RS(Key)
        If Err.Number <> 0 Then
            Err.Clear
            SaveBookmark = RS(0)
        End If
    End If
    RS.Requery
    If IsMissing(Key) Then
        Call Trace(trcBody, "RS.Find " & RS(0).Name & "='" & SQLQuote(SaveBookmark) & "'")
        RS.Find RS(0).Name & "='" & SQLQuote(SaveBookmark) & "'"
    Else
        Call Trace(trcBody, "RS.Find " & Key & "='" & SQLQuote(SaveBookmark) & "'")
        RS.Find Key & "='" & SQLQuote(SaveBookmark) & "'"
    End If
    
    For Each DBinfo In DBcollection
        If Not (DBinfo.Recordset Is RS) Then
            Call Trace(trcBody, "DBinfo.Recordset.Requery")
            DBinfo.Recordset.Requery
        End If
    Next
    Call Trace(trcExit, "RefreshCommand")
End Sub
Public Sub ReportCommand(frm As Form, RS As ADODB.Recordset, ByVal ReportPath As String)
    Dim scrApplication As New CRAXDRT.Application
    Dim Report As New CRAXDRT.Report
    Dim vRS As ADODB.Recordset
'    Dim scrApplication As New CRPEAuto.Application
'    Dim Report As New CRPEAuto.Report
    
    Call Trace(trcEnter, "ReportCommand(""" & frm.Name & """, RS, """ & ReportPath & """)")
    On Error GoTo ErrorHandler
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
    
    If Dir(ReportPath, vbNormal) = vbNullString Then
        Call MsgBox(ReportPath & " not found.", vbExclamation, App.FileDescription)
        GoTo ExitSub
    End If
    
    Call Trace(trcBody, "Set Report = scrApplication.OpenReport(""" & ReportPath & """, crOpenReportByTempCopy)")
    Set Report = scrApplication.OpenReport(ReportPath, crOpenReportByTempCopy)
    Call Trace(trcBody, "Report.Database.SetDataSource vRS, 3, 1")
    Report.Database.SetDataSource vRS, 3, 1
    Call Trace(trcBody, "Report.ReadRecords")
    Report.ReadRecords
    
    frmViewReport.scrViewer.ReportSource = Report
    frmViewReport.Show vbModal
    
ExitSub:
    Set scrApplication = Nothing
    Set Report = Nothing
    vRS.Close
    Set vRS = Nothing
    Call Trace(trcExit, "ReportCommand")
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description & " (Error " & Err.Number & ")", vbExclamation, frm.Caption
    GoTo ExitSub
    Resume Next
End Sub
Public Sub SearchCommand(frm As Form, RS As ADODB.Recordset, ByVal Key As String)
    Dim FieldList As String
    Dim TableList As String
    Dim WhereClause As String
    Dim OrderByClause As String
    Dim SQLstatement As String
    
    Call Trace(trcEnter, "SearchCommand(""" & frm.Name & """, RS, """ & Key & """)")
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
    Call Trace(trcExit, "SearchCommand")
End Sub
Public Sub SetDateFormats()
    Dim iPos As Integer
    
    Call Trace(trcEnter, "SetDateFormats")
    fmtShortDate = GetRegionalShortDateFormat()
    fmtLongDate = fmtShortDate
    If InStr(LCase(fmtLongDate), "yyyy") = 0 Then
        iPos = InStr(LCase(fmtLongDate), "yy")
        fmtLongDate = Mid(fmtLongDate, 1, iPos - 1) & "yyyy"
    End If
    fmtFullDateTime = fmtLongDate & " hh:nn:ss AMPM"
    Call Trace(trcExit, "SetDateFormats")
End Sub
Public Sub SQLCommand(ByVal TableName As String)
    Call Trace(trcEnter, "SQLCommand(""" & TableName & """")
    Load frmSQL
    Set frmSQL.cnSQL = adoConn
    If ParsePath(gstrFileDSN, DrvDirNoSlash) = gstrODBCFileDSNDir Then
        frmSQL.txtDatabase.Text = ParsePath(gstrFileDSN, FileNameBase)
    Else
        frmSQL.txtDatabase.Text = ParsePath(gstrFileDSN, DrvDirFileNameBase)
    End If
    frmSQL.dbcTables.BoundText = TableName
    frmSQL.Show vbModal
    Call Trace(trcExit, "SQLCommand")
End Sub
Public Sub UpdatePosition(frm As Form, ByVal Caption As String, RS As ADODB.Recordset)
    Dim i As Integer
    
    Call Trace(trcEnter, "UpdatePosition(""" & frm.Name & """, """ & Caption & """, RS)")
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
        If SQLfilter <> vbNullString Then
            frm.sbStatus.Panels("Message").Text = "Filter: " & SQLfilter
        End If
        frm.sbStatus.Panels("Position").Text = "Record " & RS.BookMark & " of " & RS.RecordCount
    End If
    
    frm.adodcMain.Caption = Caption
    Call Trace(trcExit, "UpdatePosition")
    Exit Sub

ErrorHandler:
    MsgBox Err.Description & " (Error " & Err.Number & ")", vbExclamation, frm.Caption
    Resume Next
End Sub
Public Function dbcValidate(fld As ADODB.Field, ctl As DataCombo) As Integer
    Dim adoRS As ADODB.Recordset
    Dim pDataSource As ADODB.Recordset
    Dim RecordsAffected As Long
    Dim SQLstring As String
    Dim FieldList As String
    Dim TableList As String
    Dim WhereClause As String
    Dim OrderByClause As String
    
    Call Trace(trcEnter, "dbcValidate(""" & fld.Name & """, """ & ctl.Name & """)")
    dbcValidate = 1
    Set pDataSource = ctl.DataSource
    If IsNull(ctl.SelectedItem) Then
        Call ParseSQLSelect(pDataSource.Source, FieldList, TableList, WhereClause, OrderByClause)
        If WhereClause <> vbNullString Then WhereClause = WhereClause & " And "
        WhereClause = WhereClause & " " & fld.Name & " like '" & SQLQuote(ctl.Text) & "%'"
        SQLstring = "select " & FieldList & " from " & TableList & " where " & WhereClause
        If OrderByClause <> vbNullString Then SQLstring = SQLstring & " order by " & OrderByClause
        
        Set adoRS = New ADODB.Recordset
        adoRS.Open SQLstring, adoConn, adOpenKeyset, adLockReadOnly
        dbcValidate = adoRS.RecordCount
        If Not adoRS.EOF Then
            ctl.BoundText = adoRS(fld.Name)
            If adoRS.RecordCount > 1 Then
                'Raise it's click event to give the user the list...
            End If
        Else
            If MsgBox("""" & ctl.Text & """ isn't in the list... Do you want it added...?", vbYesNo) = vbNo Then
                ctl.BoundText = vbNullString
                Exit Function
            Else
                dbcValidate = 1 '...to denote that it will be added...
            End If
        End If
        CloseRecordset adoRS, True
    End If
    If Len(ctl.BoundText) > fld.DefinedSize Then ctl.BoundText = Mid(ctl.BoundText, 1, fld.DefinedSize)
    If Len(ctl.Text) > fld.DefinedSize Then ctl.Text = Mid(ctl.Text, 1, fld.DefinedSize)
    
    'Sometimes the data binding doesn't get the recordset updated...
    'Why? I don't know...
    If ctl.BoundText <> pDataSource(fld.Name) Then
        pDataSource(fld.Name) = ctl.BoundText
    End If
    Call Trace(trcExit, "dbcValidate")
End Function
Public Sub dbcKeyPress(fld As ADODB.Field, ctl As DataCombo, KeyAscii As Integer)
    Dim adoRS As ADODB.Recordset
    Dim pDataSource As ADODB.Recordset
    Dim RecordsAffected As Long
    Dim SQLstring As String
    Dim FieldList As String
    Dim TableList As String
    Dim WhereClause As String
    Dim OrderByClause As String
    
    Call Trace(trcEnter, "dbcKeyPress(""" & fld.Name & """, """ & ctl.Name & """, """ & KeyAscii & """)")
    Set pDataSource = ctl.DataSource
    If IsNull(ctl.SelectedItem) Then
        Call ParseSQLSelect(pDataSource.Source, FieldList, TableList, WhereClause, OrderByClause)
        If WhereClause <> vbNullString Then WhereClause = WhereClause & " And "
        WhereClause = WhereClause & " " & fld.Name & " like '" & ctl.Text & "%'"
        SQLstring = "select " & FieldList & " from " & TableList & " where " & WhereClause
        If OrderByClause <> vbNullString Then SQLstring = SQLstring & " order by " & OrderByClause
        
        Set adoRS = New ADODB.Recordset
        adoRS.Open SQLstring, adoConn, adOpenKeyset, adLockReadOnly
        If Not adoRS.EOF Then ctl.BoundText = adoRS(fld.Name)
        CloseRecordset adoRS, True
    End If
    If Len(ctl.BoundText) > fld.DefinedSize Then ctl.BoundText = Mid(ctl.BoundText, 1, fld.DefinedSize)
    If Len(ctl.Text) > fld.DefinedSize Then ctl.Text = Mid(ctl.Text, 1, fld.DefinedSize)
    
    'Sometimes the data binding doesn't get the recordset updated...
    'Why? I don't know...
    If ctl.BoundText <> pDataSource(fld.Name) Then
        pDataSource(fld.Name) = ctl.BoundText
    End If
    Call Trace(trcExit, "dbcKeyPress")
End Sub

