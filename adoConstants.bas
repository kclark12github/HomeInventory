Attribute VB_Name = "libADOconstants"
'adoConstants - adoConstants.bas
'   ADO Conversion Module...
'   Copyright © 2000, SunGard Shareholder Systems Inc.
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Problem:    Programmer:     Description:
'   02/12/00    None        Ken Clark       Added adoDumpRecords() to display records in a recordset (a-la SQLTalk);
'   03/11/99    None        Ken Clark       Added additional command line arguments to suppress detail;
'   03/10/99    None        Ken Clark       Cleaned-up collection outputs;
'   03/04/99    None        Ken Clark       Completed (almost) all ADO objects with their related Enums;
'   03/03/99    None        Ken Clark       Created;
'=================================================================================================================================
Option Explicit
Private fUnit As Integer
Public Function adoAffect(Code As ADODB.AffectEnum) As String
    Select Case Code
        Case adAffectAllChapters
            adoAffect = "adAffectAllChapters"
        Case adAffectCurrent
            adoAffect = "adAffectCurrent"
        Case adAffectGroup
            adoAffect = "adAffectGroup"
        Case Else
            adoAffect = "Unknown code specified: " & Code
    End Select
End Function
Public Function adoBookmark(Code As ADODB.BookmarkEnum) As String
    Select Case Code
        Case adBookmarkCurrent
            adoBookmark = "adBookmarkCurrent"
        Case adBookmarkFirst
            adoBookmark = "adBookmarkFirst"
        Case adBookmarkLast
            adoBookmark = "adBookmarkLast"
        Case Else
            adoBookmark = "Unknown code specified: " & Code
    End Select
End Function
Public Function adoCommandType(Code As ADODB.CommandTypeEnum) As String
    Select Case Code
        Case adCmdFile
            adoCommandType = "adCmdFile"
        Case adCmdStoredProc
            adoCommandType = "adCmdStoredProc"
        Case adCmdTable
            adoCommandType = "adCmdTable"
        Case adCmdTableDirect
            adoCommandType = "adCmdTableDirect"
        Case adCmdText
            adoCommandType = "adCmdText"
        Case adCmdUnknown
            adoCommandType = "adCmdUnknown"
        Case Else
            adoCommandType = "Unknown code specified: " & Code
    End Select
End Function
Public Function adoCompare(Code As ADODB.CompareEnum) As String
    Select Case Code
        Case adCompareEqual
            adoCompare = "adCompareEqual"
        Case adCompareGreaterThan
            adoCompare = "adCompareGreaterThan"
        Case adCompareLessThan
            adoCompare = "adCompareLessThan"
        Case adCompareNotComparable
            adoCompare = "adCompareNotComparable"
        Case adCompareNotEqual
            adoCompare = "adCompareNotEqual"
        Case Else
            adoCompare = "Unknown code specified: " & Code
    End Select
End Function
Public Function adoConnectMode(Code As ADODB.ConnectModeEnum) As String
    Select Case Code
        Case adModeRead
            adoConnectMode = "adModeRead"
        Case adModeReadWrite
            adoConnectMode = "adModeReadWrite"
        Case adModeShareDenyNone
            adoConnectMode = "adModeShareDenyNone"
        Case adModeShareDenyRead
            adoConnectMode = "adModeShareDenyRead"
        Case adModeShareDenyWrite
            adoConnectMode = "adModeShareDenyWrite"
        Case adModeShareExclusive
            adoConnectMode = "adModeShareExclusive"
        Case adModeUnknown
            adoConnectMode = "adModeUnknown"
        Case adModeWrite
            adoConnectMode = "adModeWrite"
        Case Else
            adoConnectMode = "Unknown code specified: " & Code
    End Select
End Function
Public Function adoConnectOption(Code As ADODB.ConnectOptionEnum) As String
    Select Case Code
        Case adAsyncConnect
            adoConnectOption = "adAsyncConnect"
        Case Else
            adoConnectOption = "Unknown code specified: " & Code
    End Select
End Function
Public Function adoConnectPrompt(Code As ADODB.ConnectPromptEnum) As String
    Select Case Code
        Case adPromptAlways
            adoConnectPrompt = "adPromptAlways"
        Case adPromptComplete
            adoConnectPrompt = "adPromptComplete"
        Case adPromptCompleteRequired
            adoConnectPrompt = "adPromptCompleteRequired"
        Case adPromptNever
            adoConnectPrompt = "adPromptNever"
        Case Else
            adoConnectPrompt = "Unknown code specified: " & Code
    End Select
End Function
Public Function adoCursorLocation(Code As ADODB.CursorLocationEnum) As String
    Select Case Code
        Case adUseClient
            adoCursorLocation = "adUseClient"
        Case adUseServer
            adoCursorLocation = "adUseServer"
        Case Else
            adoCursorLocation = "Unknown code specified: " & Code
    End Select
End Function
Public Function adoCursorOption(Code As ADODB.CursorOptionEnum) As String
    Select Case Code
        Case adAddNew
            adoCursorOption = "adAddNew"
        Case adApproxPosition
            adoCursorOption = "adApproxPosition"
        Case adBookmark
            adoCursorOption = "adBookmark"
        Case adDelete
            adoCursorOption = "adDelete"
        Case adFind
            adoCursorOption = "adFind"
        Case adHoldRecords
            adoCursorOption = "adHoldRecords"
        Case adMovePrevious
            adoCursorOption = "adMovePrevious"
        Case adNotify
            adoCursorOption = "adNotify"
        Case adResync
            adoCursorOption = "adResync"
        Case adUpdate
            adoCursorOption = "adUpdate"
        Case adUpdateBatch
            adoCursorOption = "adUpdateBatch"
        Case Else
            adoCursorOption = "Unknown code specified: " & Code
    End Select
End Function
Public Function adoCursorType(Code As ADODB.CursorTypeEnum) As String
    Select Case Code
        Case adOpenDynamic
            adoCursorType = "adOpenDynamic"
        Case adOpenForwardOnly
            adoCursorType = "adOpenForwardOnly"
        Case adOpenKeyset
            adoCursorType = "adOpenKeyset"
        Case adOpenStatic
            adoCursorType = "adOpenStatic"
        Case Else
            adoCursorType = "Unknown code specified: " & Code
    End Select
End Function
Public Function adoDataType(Code As ADODB.DataTypeEnum) As String
    Select Case Code
        Case adBigInt
            adoDataType = "adBigInt"
        Case adBinary
            adoDataType = "adBinary"
        Case adBoolean
            adoDataType = "adBoolean"
        Case adBSTR
            adoDataType = "adBSTR"
        Case adChapter
            adoDataType = "adChapter"
        Case adChar
            adoDataType = "adChar"
        Case adCurrency
            adoDataType = "adCurrency"
        Case adDate
            adoDataType = "adDate"
        Case adDBDate
            adoDataType = "adDBDate"
        Case adDBFileTime
            adoDataType = "adDBFileTime"
        Case adDBTime
            adoDataType = "adDBTime"
        Case adDBTimeStamp
            adoDataType = "adDBTimeStamp"
        Case adDecimal
            adoDataType = "adDecimal"
        Case adDouble
            adoDataType = "adDouble"
        Case adEmpty
            adoDataType = "adEmpty"
        Case adError
            adoDataType = "adError"
        Case adFileTime
            adoDataType = "adFileTime"
        Case adGUID
            adoDataType = "adGUID"
        Case adIDispatch
            adoDataType = "adIDispatch"
        Case adInteger
            adoDataType = "adInteger"
        Case adIUnknown
            adoDataType = "adIUnknown"
        Case adLongVarBinary
            adoDataType = "adLongVarBinary"
        Case adLongVarChar
            adoDataType = "adLongVarChar"
        Case adLongVarWChar
            adoDataType = "adLongVarWChar"
        Case adNumeric
            adoDataType = "adNumeric"
        Case adPropVariant
            adoDataType = "adPropVariant"
        Case adSingle
            adoDataType = "adSingle"
        Case adSmallInt
            adoDataType = "adSmallInt"
        Case adTinyInt
            adoDataType = "adTinyInt"
        Case adUnsignedBigInt
            adoDataType = "adUnsignedBigInt"
        Case adUnsignedInt
            adoDataType = "adUnsignedInt"
        Case adUnsignedSmallInt
            adoDataType = "adUnsignedSmallInt"
        Case adUnsignedTinyInt
            adoDataType = "adUnsignedTinyInt"
        Case adUserDefined
            adoDataType = "adUserDefined"
        Case adVarBinary
            adoDataType = "adVarBinary"
        Case adVarChar
            adoDataType = "adVarChar"
        Case adVariant
            adoDataType = "adVariant"
        Case adVarWChar
            adoDataType = "adVarWChar"
        Case adWChar
            adoDataType = "adWChar"
        Case Else
            adoDataType = "Unknown code specified: " & Code
    End Select
End Function
Public Function adoEditMode(Code As ADODB.EditModeEnum) As String
    Select Case Code
        Case adEditAdd
            adoEditMode = "adEditAdd"
        Case adEditDelete
            adoEditMode = "adEditDelete"
        Case adEditInProgress
            adoEditMode = "adEditInProgress"
        Case adEditNone
            adoEditMode = "adEditNone"
        Case Else
            adoEditMode = "Unknown code specified: " & Code
    End Select
End Function
Public Function adoEventStatus(Code As ADODB.EventStatusEnum) As String
    Select Case Code
        Case adStatusOK
            adoEventStatus = "adStatusOK"
        Case adStatusErrorsOccurred
            adoEventStatus = "adStatusErrorsOccurred"
        Case adStatusCantDeny
            adoEventStatus = "adStatusCantDeny"
        Case adStatusCancel
            adoEventStatus = "adStatusCancel"
        Case adStatusUnwantedEvent
            adoEventStatus = "adStatusUnwantedEvent"
        Case Else
            adoEventStatus = "Unknown code specified: " & Code
    End Select
End Function
Public Function adoEventReason(Code As ADODB.EventReasonEnum) As String
    Select Case Code
        Case adRsnAddNew
            adoEventReason = "adRsnAddNew"
        Case adRsnDelete
            adoEventReason = "adRsnDelete"
        Case adRsnUpdate
            adoEventReason = "adRsnUpdate"
        Case adRsnUndoUpdate
            adoEventReason = "adRsnUndoUpdate"
        Case adRsnUndoAddNew
            adoEventReason = "adRsnUndoAddNew"
        Case adRsnUndoDelete
            adoEventReason = "adRsnUndoDelete"
        Case adRsnRequery
            adoEventReason = "adRsnRequery"
        Case adRsnResynch
            adoEventReason = "adRsnResynch"
        Case adRsnClose
            adoEventReason = "adRsnClose"
        Case adRsnMove
            adoEventReason = "adRsnMove"
        Case adRsnFirstChange
            adoEventReason = "adRsnFirstChange"
        Case adRsnMoveFirst
            adoEventReason = "adRsnMoveFirst"
        Case adRsnMoveNext
            adoEventReason = "adRsnMoveNext"
        Case adRsnMovePrevious
            adoEventReason = "adRsnMovePrevious"
        Case adRsnMoveLast
            adoEventReason = "adRsnMoveLast"
        Case Else
            adoEventReason = "Unknown code specified: " & Code
    End Select
End Function
Public Function adoExecuteOption(Code As ADODB.ExecuteOptionEnum) As String
    Select Case Code
        Case adAsyncExecute
            adoExecuteOption = "adAsyncExecute"
        Case adAsyncFetch
            adoExecuteOption = "adAsyncFetch"
        Case adAsyncFetchNonBlocking
            adoExecuteOption = "adAsyncFetchNonBlocking"
        Case adExecuteNoRecords
            adoExecuteOption = "adExecuteNoRecords"
        Case Else
            adoExecuteOption = "Unknown code specified: " & Code
    End Select
End Function
Public Function adoFieldAttribute(Code As ADODB.FieldAttributeEnum) As String
    adoFieldAttribute = ""
    If CBool(Code And adFldCacheDeferred) Then adoFieldAttribute = adoFieldAttribute & " + adFldCacheDeferred"
    If CBool(Code And adFldFixed) Then adoFieldAttribute = adoFieldAttribute & " + adFldFixed"
    If CBool(Code And adFldIsNullable) Then adoFieldAttribute = adoFieldAttribute & " + adFldIsNullable"
    If CBool(Code And adFldKeyColumn) Then adoFieldAttribute = adoFieldAttribute & " + adFldKeyColumn"
    If CBool(Code And adFldLong) Then adoFieldAttribute = adoFieldAttribute & " + adFldLong"
    If CBool(Code And adFldMayBeNull) Then adoFieldAttribute = adoFieldAttribute & " + adFldMayBeNull"
    If CBool(Code And adFldMayDefer) Then adoFieldAttribute = adoFieldAttribute & " + adFldMayDefer"
    If CBool(Code And adFldNegativeScale) Then adoFieldAttribute = adoFieldAttribute & " + adFldNegativeScale"
    If CBool(Code And adFldRowID) Then adoFieldAttribute = adoFieldAttribute & " + adFldRowID"
    If CBool(Code And adFldRowVersion) Then adoFieldAttribute = adoFieldAttribute & " + adFldRowVersion"
    If CBool(Code And adFldUnknownUpdatable) Then adoFieldAttribute = adoFieldAttribute & " + adFldUnknownUpdatable"
    If CBool(Code And adFldUpdatable) Then adoFieldAttribute = adoFieldAttribute & " + adFldUpdatable"
        
    If Left(adoFieldAttribute, 3) = " + " Then
        adoFieldAttribute = Mid(adoFieldAttribute, 4)
    Else
        adoFieldAttribute = "None"
    End If
End Function
Public Function adoFilterGroup(Code As ADODB.FilterGroupEnum) As String
    Select Case Code
        Case adFilterAffectedRecords
            adoFilterGroup = "adFilterAffectedRecords"
        Case adFilterConflictingRecords
            adoFilterGroup = "adFilterConflictingRecords"
        Case adFilterFetchedRecords
            adoFilterGroup = "adFilterFetchedRecords"
        Case adFilterNone
            adoFilterGroup = "adFilterNone"
        Case adFilterPendingRecords
            adoFilterGroup = "adFilterPendingRecords"
        Case Else
            adoFilterGroup = "Unknown code specified: " & Code
    End Select
End Function
Public Function adoGetRowsOption(Code As ADODB.GetRowsOptionEnum) As String
    Select Case Code
        Case adGetRowsRest
            adoGetRowsOption = "adGetRowsRest"
        Case Else
            adoGetRowsOption = "Unknown code specified: " & Code
    End Select
End Function
Public Function adoIsolationLevel(Code As ADODB.IsolationLevelEnum) As String
    Select Case Code
        Case adXactBrowse
            adoIsolationLevel = "adXactBrowse"
        Case adXactChaos
            adoIsolationLevel = "adXactChaos"
        Case adXactCursorStability
            adoIsolationLevel = "adXactCursorStability"
        Case adXactIsolated
            adoIsolationLevel = "adXactIsolated"
        Case adXactReadCommitted
            adoIsolationLevel = "adXactReadCommitted"
        Case adXactReadUncommitted
            adoIsolationLevel = "adXactReadUncommitted"
        Case adXactRepeatableRead
            adoIsolationLevel = "adXactRepeatableRead"
        Case adXactSerializable
            adoIsolationLevel = "adXactSerializable"
        Case adXactUnspecified
            adoIsolationLevel = "adXactUnspecified"
        Case Else
            adoIsolationLevel = "Unknown code specified: " & Code
    End Select
End Function
Public Function adoLockType(Code As ADODB.LockTypeEnum) As String
    Select Case Code
        Case adLockBatchOptimistic
            adoLockType = "adLockBatchOptimistic"
        Case adLockOptimistic
            adoLockType = "adLockOptimistic"
        Case adLockPessimistic
            adoLockType = "adLockPessimistic"
        Case adLockReadOnly
            adoLockType = "adLockReadOnly"
        Case Else
            adoLockType = "Unknown code specified: " & Code
    End Select
End Function
Public Function adoMarshalOptions(Code As ADODB.MarshalOptionsEnum) As String
    Select Case Code
        Case adMarshalAll
            adoMarshalOptions = "adMarshalAll"
        Case adMarshalModifiedOnly
            adoMarshalOptions = "adMarshalModifiedOnly"
        Case Else
            adoMarshalOptions = "Unknown code specified: " & Code
    End Select
End Function
Public Function adoObjectState(Code As ADODB.ObjectStateEnum) As String
    Select Case Code
        Case adStateOpen
            adoObjectState = "adStateOpen"
        Case adStateClosed
            adoObjectState = "adStateClosed"
        Case adStateConnecting
            adoObjectState = "adStateConnecting"
        Case adStateExecuting
            adoObjectState = "adStateExecuting"
        Case adStateFetching
            adoObjectState = "adStateFetching"
        Case Else
            adoObjectState = "Unknown code specified: " & Code
    End Select
End Function
Public Function adoParameterAttribute(Code As ADODB.ParameterAttributesEnum) As String
    adoParameterAttribute = ""
    If CBool(Code And adParamSigned) Then adoParameterAttribute = adoParameterAttribute & " + adParamSigned"
    If CBool(Code And adParamNullable) Then adoParameterAttribute = adoParameterAttribute & " + adParamNullable"
    If CBool(Code And adParamLong) Then adoParameterAttribute = adoParameterAttribute & " + adParamLong"
        
    If Left(adoParameterAttribute, 3) = " + " Then
        adoParameterAttribute = Mid(adoParameterAttribute, 4)
    Else
        adoParameterAttribute = "None"
    End If
End Function
Public Function adoPosition(Code As ADODB.PositionEnum) As String
    Select Case Code
        Case adPosBOF
            adoPosition = "adPosBOF"
        Case adPosEOF
            adoPosition = "adPosEOF"
        Case adPosUnknown
            adoPosition = "adPosUnknown"
        Case Else
            adoPosition = Code  'Actual Page/Record position...
    End Select
End Function
Public Function adoPropertyAttribute(Code As ADODB.PropertyAttributesEnum) As String
    adoPropertyAttribute = ""
    If CBool(Code And adPropNotSupported) Then adoPropertyAttribute = adoPropertyAttribute & " + adPropNotSupported"
    If CBool(Code And adPropOptional) Then adoPropertyAttribute = adoPropertyAttribute & " + adPropOptional"
    If CBool(Code And adPropRead) Then adoPropertyAttribute = adoPropertyAttribute & " + adPropRead"
    If CBool(Code And adPropRequired) Then adoPropertyAttribute = adoPropertyAttribute & " + adPropRequired"
    If CBool(Code And adPropWrite) Then adoPropertyAttribute = adoPropertyAttribute & " + adPropWrite"
        
    If Left(adoPropertyAttribute, 3) = " + " Then
        adoPropertyAttribute = Mid(adoPropertyAttribute, 4)
    Else
        adoPropertyAttribute = "None"
    End If
End Function
Public Function adoRecordStatus(Code As ADODB.RecordStatusEnum) As String
    Select Case Code
        Case adRecCanceled
            adoRecordStatus = "adRecCanceled"
        Case adRecCantRelease
            adoRecordStatus = "adRecCantRelease"
        Case adRecConcurrencyViolation
            adoRecordStatus = "adRecConcurrencyViolation"
        Case adRecDBDeleted
            adoRecordStatus = "adRecDBDeleted"
        Case adRecDeleted
            adoRecordStatus = "adRecDeleted"
        Case adRecIntegrityViolation
            adoRecordStatus = "adRecIntegrityViolation"
        Case adRecInvalid
            adoRecordStatus = "adRecInvalid"
        Case adRecMaxChangesExceeded
            adoRecordStatus = "adRecMaxChangesExceeded"
        Case adRecModified
            adoRecordStatus = "adRecModified"
        Case adRecMultipleChanges
            adoRecordStatus = "adRecMultipleChanges"
        Case adRecNew
            adoRecordStatus = "adRecNew"
        Case adRecObjectOpen
            adoRecordStatus = "adRecObjectOpen"
        Case adRecOK
            adoRecordStatus = "adRecOK"
        Case adRecOutOfMemory
            adoRecordStatus = "adRecOutOfMemory"
        Case adRecPendingChanges
            adoRecordStatus = "adRecPendingChanges"
        Case adRecPermissionDenied
            adoRecordStatus = "adRecPermissionDenied"
        Case adRecSchemaViolation
            adoRecordStatus = "adRecSchemaViolation"
        Case adRecUnmodified
            adoRecordStatus = "adRecUnmodified"
        Case Else
            adoRecordStatus = "Unknown code specified: " & Code
    End Select
End Function
Public Function adoResync(Code As ADODB.ResyncEnum) As String
    Select Case Code
        Case adResyncAllValues
            adoResync = "adResyncAllValues"
        Case adResyncUnderlyingValues
            adoResync = "adResyncUnderlyingValues"
        Case Else
            adoResync = "Unknown code specified: " & Code
    End Select
End Function
Public Function adoSchema(Code As ADODB.SchemaEnum) As String
    adoSchema = "Not coded yet."
End Function
Public Function adoSearchDirection(Code As ADODB.SearchDirectionEnum) As String
    Select Case Code
        Case adSearchBackward
            adoSearchDirection = "adSearchBackward"
        Case adSearchForward
            adoSearchDirection = "adSearchForward"
        Case Else
            adoSearchDirection = "Unknown code specified: " & Code
    End Select
End Function
Public Function adoStringFormat(Code As ADODB.StringFormatEnum) As String
    Select Case Code
        Case adClipString
            adoStringFormat = "adClipString"
        Case Else
            adoStringFormat = "Unknown code specified: " & Code
    End Select
End Function
Public Function adoXactAttribute(Code As ADODB.XactAttributeEnum) As String
    adoXactAttribute = ""
    If CBool(Code And adXactAbortRetaining) Then adoXactAttribute = adoXactAttribute & " + adXactAbortRetaining"
    If CBool(Code And adXactCommitRetaining) Then adoXactAttribute = adoXactAttribute & " + adXactCommitRetaining"
        
    If Left(adoXactAttribute, 3) = " + " Then
        adoXactAttribute = Mid(adoXactAttribute, 4)
    Else
        adoXactAttribute = "None"
    End If
End Function
Private Sub PrintOut(pString As String)
    If fUnit <> 0 Then
        Print #fUnit, pString
    Else
        Debug.Print pString
    End If
End Sub
Public Sub adoDumpErrors(ByVal pErrors As ADODB.Errors, indent As Integer)
    Dim i As Integer
    Dim e As ADODB.Error
    Dim Tabs As String
    
    On Error Resume Next
    For i = 1 To indent
        Tabs = Tabs & vbTab
    Next
    
    For i = 0 To pErrors.Count - 1
        Set e = pErrors(i)
        PrintOut Tabs & ".Errors(" & i & ").Description:" & vbTab & e.Description
        PrintOut Tabs & vbTab & ".HelpContext: " & e.HelpContext
        PrintOut Tabs & vbTab & ".HelpFile:    " & e.HelpFile
        PrintOut Tabs & vbTab & ".NativeError: " & e.NativeError
        PrintOut Tabs & vbTab & ".Number:      " & e.Number
        PrintOut Tabs & vbTab & ".Source:      " & e.Source
        PrintOut Tabs & vbTab & ".SQLState:    " & e.SQLState
    Next
End Sub
Public Sub adoDumpRSField(ByVal fld As ADODB.Field, indent As Integer, Optional strArgs As String, Optional ByVal pRS As ADODB.Recordset)
    Dim i As Integer
    Dim Tabs As String
    Dim Args As String
    
    If Not IsMissing(strArgs) Then
        Args = strArgs
    End If
    
    On Error Resume Next
    For i = 1 To indent
        Tabs = Tabs & vbTab
    Next
    If Not IsMissing(pRS) Then
        For i = 0 To pRS.Fields.Count - 1
            If fld.Name = pRS.Fields(i).Name Then Exit For
        Next i
        PrintOut Tabs & ".Fields(" & i & ").Name:" & vbTab & fld.Name
    Else
        PrintOut Tabs & ".Field.Name:" & vbTab & fld.Name
    End If
    PrintOut Tabs & vbTab & ".ActualSize:      " & fld.ActualSize
    PrintOut Tabs & vbTab & ".Attributes:      " & adoFieldAttribute(fld.Attributes) & " (" & fld.Attributes & ")"
    PrintOut Tabs & vbTab & ".DefinedSize:     " & fld.DefinedSize
    PrintOut Tabs & vbTab & ".NumericScale:    " & fld.NumericScale
    PrintOut Tabs & vbTab & ".OriginalValue:   " & fld.OriginalValue
    PrintOut Tabs & vbTab & ".Precision:       " & fld.Precision
    If InStr(UCase(Args), "NOPROP") = 0 Then
        PrintOut Tabs & vbTab & ".Properties:      "
        adoDumpProperties fld.Properties, indent + 1
    End If
    PrintOut Tabs & vbTab & ".Type:            " & adoDataType(fld.Type) & " (" & fld.Type & ")"
    PrintOut Tabs & vbTab & ".UnderlyingValue: " & fld.UnderlyingValue
    PrintOut Tabs & vbTab & ".Value:           " & fld.Value
End Sub
Public Sub adoDumpFields(ByVal pFields As ADODB.Fields, indent As Integer, Optional ByVal pRS As ADODB.Recordset)
    Dim i As Integer
    Dim fld As ADODB.Field
    Dim Tabs As String
    
    On Error Resume Next
    For i = 1 To indent
        Tabs = Tabs & vbTab
    Next
    
    For i = 0 To pFields.Count - 1
        Set fld = pFields(i)
        If IsMissing(pRS) Then
            adoDumpRSField fld, indent
        Else
            adoDumpRSField fld, indent, "", pRS
        End If
    Next
End Sub
Public Sub adoDumpParameters(ByVal pParam As ADODB.Parameters, indent As Integer)
    Dim i As Integer
    Dim p As ADODB.Parameter
    Dim Tabs As String
    
    On Error Resume Next
    For i = 1 To indent
        Tabs = Tabs & vbTab
    Next
    
    For i = 0 To pParam.Count - 1
        Set p = pParam(i)
        PrintOut Tabs & ".Parameters(" & i & ").Name:" & vbTab & p.Name
        PrintOut Tabs & vbTab & ".Attributes:      " & adoParameterAttribute(p.Attributes) & " (" & p.Attributes & ")"
        PrintOut Tabs & vbTab & ".Direction:       " & p.Direction
        PrintOut Tabs & vbTab & ".NumericScale:    " & p.NumericScale
        PrintOut Tabs & vbTab & ".Precision:       " & p.Precision
        PrintOut Tabs & vbTab & ".Properties:      "
        adoDumpProperties p.Properties, indent + 1
        PrintOut Tabs & vbTab & ".Size:            " & p.Size
        PrintOut Tabs & vbTab & ".Type:            " & adoDataType(p.Type) & " (" & p.Type & ")"
        PrintOut Tabs & vbTab & ".Value:           " & p.Value
    Next
End Sub
Public Sub adoDumpProperties(ByVal pProperties As ADODB.Properties, indent As Integer)
    Dim i As Integer
    Dim prop As ADODB.Property
    Dim Tabs As String
    
    On Error Resume Next
    For i = 1 To indent
        Tabs = Tabs & vbTab
    Next
    
    For i = 0 To pProperties.Count - 1
        Set prop = pProperties(i)
        PrintOut Tabs & ".Properties(" & i & ").Name:" & vbTab & prop.Name
        PrintOut Tabs & vbTab & ".Attributes:      " & adoPropertyAttribute(prop.Attributes) & " (" & prop.Attributes & ")"
        PrintOut Tabs & vbTab & ".Type:            " & adoDataType(prop.Type) & " (" & prop.Type & ")"
        PrintOut Tabs & vbTab & ".Value:           " & prop.Value
    Next
End Sub
Public Sub adoDumpCommand(ByVal pCMD As ADODB.Command, Optional FileName As String, Optional strArgs As String)
    Dim Args As String
    
    If Not IsMissing(strArgs) Then
        Args = strArgs
    End If
    If Not IsMissing(FileName) And FileName <> "" Then
        fUnit = FreeFile
        Open FileName For Append As #fUnit
    End If
    On Error Resume Next
    PrintOut String(132, "=")
    PrintOut "Command.ActiveConnection: " & pCMD.ActiveConnection.ConnectionString
    PrintOut "Command.CommandText:      " & pCMD.CommandText
    PrintOut "Command.CommandTimeout:   " & pCMD.CommandTimeout
    PrintOut "Command.CommandType:      " & adoCommandType(pCMD.CommandType)
    PrintOut "Command.Name:             " & pCMD.Name
    PrintOut "Command.Parameters:       "
    If InStr(UCase(Args), "NOPARAM") = 0 Then adoDumpParameters pCMD.Parameters, 1
    PrintOut "Command.Prepared:         " & pCMD.Prepared
    PrintOut "Command.Properties:       "
    If InStr(UCase(Args), "NOPROP") = 0 Then adoDumpProperties pCMD.Properties, 1
    PrintOut "Command.State:            " & adoObjectState(pCMD.State)
    If fUnit <> 0 Then Close #fUnit
End Sub
Public Sub adoDumpConnection(ByVal pConnection As ADODB.Connection, Optional FileName As String, Optional strArgs As String)
    Dim Args As String
    
    If Not IsMissing(strArgs) Then
        Args = strArgs
    End If
    If Not IsMissing(FileName) And FileName <> "" Then
        fUnit = FreeFile
        Open FileName For Append As #fUnit
    End If
    On Error Resume Next
    PrintOut String(132, "=")
    PrintOut "Connection.Attributes:        " & adoXactAttribute(pConnection.Attributes)
    PrintOut "Connection.CommandTimeout:    " & pConnection.CommandTimeout
    PrintOut "Connection.ConnectionString:  " & pConnection.ConnectionString
    PrintOut "Connection.ConnectionTimeout: " & pConnection.ConnectionTimeout
    PrintOut "Connection.CursorLocation:    " & adoCursorLocation(pConnection.CursorLocation)
    PrintOut "Connection.DefaultDatabase:   " & pConnection.DefaultDatabase
    PrintOut "Connection.Errors: "
    If InStr(UCase(Args), "NOERR") = 0 Then adoDumpErrors pConnection.Errors, 1
    PrintOut "Connection.IsolationLevel:    " & adoIsolationLevel(pConnection.IsolationLevel)
    PrintOut "Connection.Mode:              " & adoConnectMode(pConnection.mode)
    PrintOut "Connection.Properties: "
    If InStr(UCase(Args), "NOPROP") = 0 Then adoDumpProperties pConnection.Properties, 1
    PrintOut "Connection.Provider:          " & pConnection.Provider
    PrintOut "Connection.State:             " & adoObjectState(pConnection.State)
    PrintOut "Connection.Version:           " & pConnection.Version
    If fUnit <> 0 Then Close #fUnit
End Sub
Private Sub CloseRS(adoRS As ADODB.Recordset, Destroy As Boolean)
    On Error Resume Next
    If Not adoRS Is Nothing Then
        If (adoRS.State And adStateOpen) = adStateOpen Then
            adoRS.CancelUpdate
            adoRS.Close
        End If
        If Destroy Then Set adoRS = Nothing
    End If
End Sub
Private Function FormatField(fld As ADODB.Field, fDelimitted As Boolean) As String
    Dim strVal As String
    If IsNull(fld.Value) Then
        strVal = "Null"
    Else
        Select Case fld.Type
            Case adBigInt
                strVal = fld.Value
            Case adBinary
                strVal = "Binary Data"
            Case adBoolean
                strVal = fld.Value
            Case adBSTR
                strVal = fld.Value
            Case adChapter
                strVal = "Chapter"
            Case adChar
                If fDelimitted Then
                    strVal = """" & fld.Value & """"
                Else
                    strVal = fld.Value
                End If
            Case adCurrency
                strVal = Format(fld.Value, "Currency")
            Case adDate
                strVal = fld.Value
            Case adDBDate
                strVal = fld.Value
            Case adDBFileTime
                strVal = fld.Value
            Case adDBTime
                strVal = fld.Value
            Case adDBTimeStamp
                strVal = fld.Value
            Case adDecimal
                strVal = fld.Value
            Case adDouble
                strVal = fld.Value
            Case adEmpty
                strVal = "Empty"
            Case adError
                strVal = "Error"
            Case adFileTime
                strVal = fld.Value
            Case adGUID
                strVal = "GUID"
            Case adIDispatch
                strVal = "IDispatch"
            Case adInteger
                strVal = fld.Value
            Case adIUnknown
                strVal = "IUnknown"
            Case adLongVarBinary
                strVal = fld.Value
            Case adLongVarChar
                If fDelimitted Then
                    strVal = """" & fld.Value & """"
                Else
                    strVal = fld.Value
                End If
            Case adLongVarWChar
                If fDelimitted Then
                    strVal = """" & fld.Value & """"
                Else
                    strVal = fld.Value
                End If
            Case adNumeric
                strVal = fld.Value
            Case adPropVariant
                strVal = fld.Value
            Case adSingle
                strVal = fld.Value
            Case adSmallInt
                strVal = fld.Value
            Case adTinyInt
                strVal = fld.Value
            Case adUnsignedBigInt
                strVal = fld.Value
            Case adUnsignedInt
                strVal = fld.Value
            Case adUnsignedSmallInt
                strVal = fld.Value
            Case adUnsignedTinyInt
                strVal = fld.Value
            Case adUserDefined
                strVal = fld.Value
            Case adVarBinary
                strVal = "Binary Data"
            Case adVarChar
                If fDelimitted Then
                    strVal = """" & fld.Value & """"
                Else
                    strVal = fld.Value
                End If
            Case adVariant
                strVal = fld.Value
            Case adVarWChar
                If fDelimitted Then
                    strVal = """" & fld.Value & """"
                Else
                    strVal = fld.Value
                End If
            Case adWChar
                If fDelimitted Then
                    strVal = """" & fld.Value & """"
                Else
                    strVal = fld.Value
                End If
            Case Else
                strVal = fld.Value
        End Select
    End If
    FormatField = strVal
End Function
Public Sub adoDumpRecord(ByVal pRS As ADODB.Recordset, Optional FileName As String)
    Dim MaxLen As Integer
    Dim fld As ADODB.Field
    Dim i As Integer
    Dim strOut As String
    
    If Not IsMissing(FileName) And FileName <> "" Then
        fUnit = FreeFile
        Open FileName For Append As #fUnit
    End If
    
    On Error Resume Next
    
    For Each fld In pRS.Fields
        If Len(fld.Name) > MaxLen Then MaxLen = Len(fld.Name)
    Next fld
    If pRS.BOF And pRS.EOF Then
        PrintOut "Recordset is Empty..."
        GoTo ExitSub
    ElseIf pRS.BOF Then
        PrintOut "BOF"
        GoTo ExitSub
    ElseIf pRS.EOF Then
        PrintOut "EOF"
        GoTo ExitSub
    End If

    PrintOut String(132, "=")
    PrintOut "Record #" & pRS.BookMark & " (bookmark) of " & pRS.RecordCount & " records..."
    PrintOut "Field: Name:" & String(MaxLen - Len("Field Name:"), " ") & "Value:"
    PrintOut String(132, "-")
    For i = 0 To pRS.Fields.Count - 1
        Set fld = pRS.Fields(i)
        strOut = "[" & Format(i + 1, "0000") & "] " & fld.Name & String(MaxLen - Len(fld.Name), " ") & vbTab & FormatField(fld, True)
        PrintOut strOut
    Next i
    
ExitSub:
    If fUnit <> 0 Then Close #fUnit
    fUnit = 0
End Sub
Public Sub adoDumpRecords(pRS As ADODB.Recordset, Optional Caption As String, Optional FileName As String)
    Dim BookMark As Variant
    Dim MaxLen As Integer
    Dim ColLen() As Integer
    Dim TotalLen As Integer
    Dim fld As ADODB.Field
    Dim i As Integer
    Dim strOut As String
    Dim strVal As String
    Dim vRS As ADODB.Recordset
    
    If Not IsMissing(FileName) And FileName <> "" Then
        fUnit = FreeFile
        Open FileName For Output As #fUnit
    End If
    
    If pRS Is Nothing Then
        PrintOut "Recordset is Nothing..."
        GoTo ExitSub
    ElseIf (pRS.State And adStateOpen) <> adStateOpen Then
        PrintOut "Recordset is Closed..."
        GoTo ExitSub
    ElseIf pRS.BOF And pRS.EOF Then
        If pRS.Filter <> 0 Then PrintOut "Filter: " & pRS.Filter
        PrintOut "Recordset is Empty..."
        GoTo ExitSub
    End If
    
    'On Error Resume Next
    If pRS.EOF Then
        BookMark = "EOF"
    ElseIf pRS.BOF Then
        BookMark = "BOF"
    Else
        BookMark = pRS.BookMark
    End If
    Set vRS = pRS.Clone
    vRS.Filter = pRS.Filter
    
    ReDim ColLen(vRS.Fields.Count)
    For i = 0 To vRS.Fields.Count - 1
        Set fld = vRS.Fields(i)
        ColLen(i) = Len(fld.Name) + 1
    Next i
    
    vRS.MoveFirst
    While Not vRS.EOF
        For i = 0 To vRS.Fields.Count - 1
            Set fld = vRS.Fields(i)
            strVal = FormatField(fld, False)
            If Len(strVal) > ColLen(i) Then ColLen(i) = Len(strVal) + 1
        Next i
        vRS.MoveNext
    Wend
    TotalLen = 0
    For i = 0 To vRS.Fields.Count - 1
        TotalLen = TotalLen + ColLen(i) + 1
    Next i
    
    strOut = "Record   "
    TotalLen = TotalLen + Len(strOut)
    PrintOut String(TotalLen, "=")
    If Not IsMissing(Caption) Then PrintOut Caption
    If vRS.Filter <> 0 Then PrintOut "Filter: " & vRS.Filter
    PrintOut "Record #" & BookMark & " (bookmark) of " & vRS.RecordCount & " records..."
    For i = 0 To vRS.Fields.Count - 1
        Set fld = vRS.Fields(i)
        strOut = strOut & fld.Name & String(ColLen(i) - Len(fld.Name), " ") & " "
    Next i
    PrintOut strOut
    PrintOut String(TotalLen, "-")
    vRS.MoveFirst
    While Not vRS.EOF
        If vRS.BookMark = BookMark Then
            strOut = "[->" & Format(vRS.BookMark, "0000") & "] "
        Else
            strOut = "[" & Format(vRS.BookMark, "000000") & "] "
        End If
        For i = 0 To vRS.Fields.Count - 1
            strVal = FormatField(vRS.Fields(i), False)
            strOut = strOut & strVal & String(ColLen(i) - Len(strVal), " ") & " "
        Next i
        PrintOut strOut
        vRS.MoveNext
    Wend
    PrintOut vbCr & vRS.RecordCount & " Records"

ExitSub:
    Call CloseRS(vRS, True)
    If fUnit <> 0 Then Close #fUnit
    fUnit = 0
End Sub
Public Sub adoDumpRecordset(ByVal pRS As ADODB.Recordset, Optional FileName As String, Optional strArgs As String)
    Dim Args As String
    
    If Not IsMissing(strArgs) Then
        Args = strArgs
    End If
    If Not IsMissing(FileName) And FileName <> "" Then
        fUnit = FreeFile
        Open FileName For Append As #fUnit
    End If
    
    On Error Resume Next
    
    PrintOut String(132, "=")
    PrintOut "Recordset.AboslutePage:     " & adoPosition(pRS.AbsolutePage)
    PrintOut "Recordset.AbsolutePosition: " & adoPosition(pRS.AbsolutePosition)
    PrintOut "Recordset.ActiveCommand:    " & pRS.ActiveCommand.CommandText
    PrintOut "Recordset.ActiveConnection: " & pRS.ActiveConnection.ConnectionString
    PrintOut "Recordset.BOF:              " & pRS.BOF
    PrintOut "Recordset.Bookmark:         " & pRS.BookMark
    PrintOut "Recordset.CacheSize:        " & pRS.CacheSize
    PrintOut "Recordset.CursorLocation:   " & adoCursorLocation(pRS.CursorLocation)
    PrintOut "Recordset.CursorType:       " & adoCursorType(pRS.CursorType)
    PrintOut "Recordset.DataMember:       " & pRS.DataMember
    If pRS.DataSource Is Nothing Then
        PrintOut "Recordset.DataSource:       Nothing"
    Else
        PrintOut "Recordset.DataSource:       Not Nothing"
    End If
    PrintOut "Recordset.EditMode:         " & adoEditMode(pRS.EditMode)
    PrintOut "Recordset.EOF:              " & pRS.EOF
    PrintOut "Recordset.Fields:           "
    PrintOut vbTab & ".Fields.Count:     " & pRS.Fields.Count
    If InStr(UCase(Args), "NOFIELD") = 0 Then adoDumpFields pRS.Fields, 1, pRS
    PrintOut "Recordset.Filter:           " & pRS.Filter
    PrintOut "Recordset.LockType:         " & adoLockType(pRS.LockType)
    PrintOut "Recordset.MarshalOptions:   " & adoMarshalOptions(pRS.MarshalOptions)
    PrintOut "Recordset.MaxRecords:       " & pRS.MaxRecords
    PrintOut "Recordset.PageCount:        " & pRS.PageCount
    PrintOut "Recordset.PageSize:         " & pRS.PageSize
    PrintOut "Recordset.Properties:       "
    If InStr(UCase(Args), "NOPROP") = 0 Then adoDumpProperties pRS.Properties, 1
    PrintOut "Recordset.RecordCount:      " & pRS.RecordCount
    PrintOut "Recordset.Sort:             " & pRS.Sort
    PrintOut "Recordset.Source:           " & pRS.Source
    PrintOut "Recordset.State:            " & adoObjectState(pRS.State)
    PrintOut "Recordset.Status:           " & adoRecordStatus(pRS.Status)
    PrintOut "Recordset.StayInSync:       " & pRS.StayInSync
    If fUnit <> 0 Then Close #fUnit
End Sub
Public Sub adoDumpField(ByVal pField As ADODB.Field, Optional FileName As String, Optional strArgs As String)
    Dim Args As String
    
    If Not IsMissing(strArgs) Then
        Args = strArgs
    End If
    If Not IsMissing(FileName) And FileName <> "" Then
        fUnit = FreeFile
        Open FileName For Append As #fUnit
    End If
    
    On Error Resume Next
    
    PrintOut String(132, "=")
    adoDumpRSField pField, 1, Args
    If fUnit <> 0 Then Close #fUnit
End Sub



