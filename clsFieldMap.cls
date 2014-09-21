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
'FieldMap - FieldMap.cls
'   Field Map Collection...
'   Copyright © 1998-2002, SunGard Investor Accounting Systems
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Problem:    Programmer:     Description:
'   02/19/00    None        Ken Clark       Added Class Initialize and Terminate event handlers (cause of memory leak in
'                                           frmACReview);
'   01/14/00    None        Ken Clark       Added .LabelControl, .Format, and .Tag;
'   02/17/99    None        Ken Clark       Added .Caption property;
'   11/18/98    None        Ken Clark       Added History;
'=================================================================================================================================
Option Explicit
Private mScreenControl As Control
Private mDataFieldName As String
Private mLabelControl As Control
Private mCaption As String
Private mFormat As String
Private mTag As String
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
    Set mLabelControl = Nothing
    mCaption = vbNullString
    mFormat = vbNullString
    mTag = vbNullString
End Sub
Private Sub Class_Terminate()
    Set mScreenControl = Nothing
    mDataFieldName = vbNullString
    Set mLabelControl = Nothing
    mCaption = vbNullString
    mFormat = vbNullString
    mTag = vbNullString
End Sub
