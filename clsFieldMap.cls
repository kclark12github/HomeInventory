VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FieldMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'FieldMap - clsFieldMap.cls
'   Screen Control to Database Field Mapping Object...
'   Copyright © 1999-2002, Ken Clark
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Description:
'   08/20/02    Added DataType and OriginalValue properties;
'=================================================================================================================================
Option Explicit
Private mScreenControl As Control
Private mDataFieldName As String
Private mDataType As ADODB.DataTypeEnum
Private mLabelControl As Control
Private mCaption As String
Private mFormat As String
Private mTag As String
Private mOriginalValue As Variant
Public Property Get Caption() As String
    Caption = mCaption
End Property
Public Property Let Caption(ByVal vNewValue As String)
    mCaption = vNewValue
End Property
Public Property Get DataField() As String
    DataField = mDataFieldName
End Property
Public Property Let DataField(ByVal vNewValue As String)
    mDataFieldName = vNewValue
End Property
Public Property Get DataType() As ADODB.DataTypeEnum
    DataType = mDataType
End Property
Public Property Let DataType(ByVal vNewValue As ADODB.DataTypeEnum)
    mDataType = vNewValue
End Property
Public Property Get Format() As String
    Format = mFormat
End Property
Public Property Let Format(ByVal vNewValue As String)
    mFormat = vNewValue
End Property
Public Property Get LabelControl() As Control
    Set LabelControl = mLabelControl
End Property
Public Property Set LabelControl(ByRef vNewValue As Control)
    Set mLabelControl = vNewValue
End Property
Public Property Get OriginalValue() As Variant
    OriginalValue = mOriginalValue
End Property
Public Property Let OriginalValue(vNewValue As Variant)
    mOriginalValue = vNewValue
End Property
Public Property Get ScreenControl() As Control
    Set ScreenControl = mScreenControl
End Property
Public Property Set ScreenControl(ByRef vNewValue As Control)
    Set mScreenControl = vNewValue
End Property
Public Property Get Tag() As String
    Tag = mTag
End Property
Public Property Let Tag(ByVal vNewValue As String)
    mTag = vNewValue
End Property
Private Sub Class_Initialize()
    Set mScreenControl = Nothing
    mDataFieldName = vbNullString
    mDataType = adVarChar
    Set mLabelControl = Nothing
    mCaption = vbNullString
    mFormat = vbNullString
    mTag = vbNullString
    mOriginalValue = Empty
End Sub
Private Sub Class_Terminate()
    Set mScreenControl = Nothing
    mDataFieldName = vbNullString
    mDataType = adVarChar
    Set mLabelControl = Nothing
    mCaption = vbNullString
    mFormat = vbNullString
    mTag = vbNullString
    mOriginalValue = Empty
End Sub
