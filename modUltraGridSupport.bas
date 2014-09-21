Attribute VB_Name = "modUltraGridSupport"
'modUltraGridSupport - modUltraGridSupport.bas
'   Utility Module Handling Column Sizing for Infragistics UltraGrid Control...
'   Copyright © 2001, SunGard Investor Accounting Systems
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Problem:    Programmer:     Description:
'   11/06/01    None        Ken Clark       Taken from Infragistics Sample Code;
'=================================================================================================================================
Option Explicit
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type VBPOINT
    X As Single
    Y As Single
End Type

Private Const SM_CXHSCROLL = 21

Private Declare Function GetMessagePos Lib "user32" () As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Function ScaledPointsFromMessagePos(ByVal hwnd As Long, ByVal ScalingObj As Object) As VBPOINT
    Dim sngX As Single
    Dim sngY As Single
    Dim pt As POINTAPI
    Dim UIElement As UltraGrid.SSUIElement
    Dim lMessagePos As Long
    
    On Error Resume Next
    
    lMessagePos = GetMessagePos()
    
    'extract the screen coordinates from the results of GetMessagePos
    pt.X = LoWord(lMessagePos)
    pt.Y = HiWord(lMessagePos)
    
    Call ScreenToClient(hwnd, pt)
    
    'scale the client coordinates from pixels to scale units
    sngX = ScaleXCoor(ScalingObj, pt.X, vbPixels, ScalingObj.ScaleMode)
    sngY = ScaleYCoor(ScalingObj, pt.Y, vbPixels, ScalingObj.ScaleMode)
    
    ScaledPointsFromMessagePos.X = sngX
    ScaledPointsFromMessagePos.Y = sngY
End Function
Public Sub AutoSizeColFromMessagePos(ByVal Grd As UltraGrid.SSUltraGrid, _
        Optional ByVal IncludeColHeader As Boolean = True, _
        Optional ByVal ScaleMode As VBRUN.ScaleModeConstants = vbPixels, _
        Optional ByVal ScalingObj As Object, _
        Optional ByVal AllowSizeDecrease As Boolean = True)
    Dim sngX As Single
    Dim sngY As Single
    Dim pt As POINTAPI
    Dim UIElement As UltraGrid.SSUIElement
    Dim lMessagePos As Long
    
    On Error Resume Next
    
    lMessagePos = GetMessagePos()
    
    'extract the screen coordinates from the results of GetMessagePos
    pt.X = LoWord(lMessagePos)
    pt.Y = HiWord(lMessagePos)
    
    Call ScreenToClient(Grd.hwnd, pt)
    
    'scale the client coordinates from pixels to scale units
    sngX = ScaleXCoor(ScalingObj, pt.X, vbPixels, ScaleMode)
    sngY = ScaleYCoor(ScalingObj, pt.Y, vbPixels, ScaleMode)
    
    Set UIElement = ColHeaderSizeArea(sngX, sngY, Grd, ScalingObj, ScaleMode)
    AutoSizeColumn Grd, UIElement.Header.Column, IncludeColHeader, ScaleMode, ScalingObj, AllowSizeDecrease
End Sub
Public Sub AutoSizeColumn(ByVal Grid As UltraGrid.SSUltraGrid, _
        ByVal Column As UltraGrid.SSColumn, _
        Optional ByVal IncludeColHeader As Boolean = True, _
        Optional ByVal ScaleMode As VBRUN.ScaleModeConstants = vbPixels, _
        Optional ByVal ScalingObj As Object, _
        Optional ByVal AllowSizeDecrease As Boolean = True, _
        Optional ByVal RowInBand As UltraGrid.SSRow)
    Dim objRow As UltraGrid.SSRow
    Dim objCell As UltraGrid.SSCell
    Dim RowScrollRegion As UltraGrid.SSRowScrollRegion
    Dim uiHeader As UltraGrid.SSUIElement
    Dim uiText As UltraGrid.SSUIElement
    Dim str As String
    Dim sngMax As Single
    Dim sngExtra As Single
    Dim sngHeader As Single
    Dim bHasButton As Boolean
    
    If Column Is Nothing Then Exit Sub
    
    'in order for TextWidth to return a value in the correct scaleunits...
    If Not ScalingObj Is Nothing Then
        ScalingObj.ScaleMode = ScaleMode
    End If
    
    'determine if we need to accomodate a button in the cell
    Select Case Column.Style
        Case UltraGrid.ssStyleCheckBox
        Case UltraGrid.ssStyleButton
        Case UltraGrid.ssStyleEdit, UltraGrid.ssStyleHTML
        Case UltraGrid.ssStyleDefault
            Select Case Column.DataType
                Case UltraGrid.ssDataTypeDate
                    bHasButton = True
                Case Else
                    If Not Column.ValueList Is Nothing Then bHasButton = True
            End Select
        Case Else
            bHasButton = True
    End Select
    
    'Iterate through all of the rows in the band?
    If Not RowInBand Is Nothing Then
        'if so, grab the first row in this band and walk down
        Set objRow = RowInBand.GetSibling(ssSiblingRowFirst)
        Do
            Set objCell = objRow.Cells(Column.Key)  'get the cell object in this row
            'Resolve its appearance - this way we have exact values to work with and not defaults
            With objCell.ResolveAppearance
                'Set the font of the scaling obj so that textwidth will work correctly
                Set ScalingObj.Font = .Font
                
                If Not IsEmpty(.Picture) Then
                    If IsObject(.Picture) Then
                        'Add the width of the picture
                        If Not .Picture Is Nothing Then sngExtra = ScaleXCoor(ScalingObj, .Picture.Width, vbHimetric, ScaleMode)
                    Else
                        'Can only get size of picture if it is in the internal images collection
                        If Not Grid.UseImageList Then sngExtra = ScaleXCoor(ScalingObj, Grid.Images(.Picture).Picture.Width, vbHimetric, ScaleMode)
                    End If
                End If
            End With
            
            'Get the text...
            str = objCell.GetText(Column.MaskDisplayMode)

            'Store only the maximum width
            sngMax = Max(sngMax, ScalingObj.TextWidth(str) + sngExtra)
            
            'Move to the next row if there is one
            If objRow.HasNextSibling(False) Then
                Set objRow = objRow.GetSibling(ssSiblingRowNext)
            Else
                Set objRow = Nothing
            End If
            
            'Reset the variable used to hold extra sizing info
            sngExtra = 0
        Loop Until objRow Is Nothing
    Else
        For Each RowScrollRegion In Grid.RowScrollRegions
            For Each objRow In RowScrollRegion.VisibleRows
                If objRow.Band.IsSameAs(Column.Band) Then
                Set objCell = objRow.Cells(Column.Key)
                
                With objCell.ResolveAppearance
                    Set ScalingObj.Font = .Font
                
                    If Not IsEmpty(.Picture) Then
                        If IsObject(.Picture) Then
                            If Not .Picture Is Nothing Then sngExtra = ScaleXCoor(ScalingObj, .Picture.Width, vbHimetric, ScaleMode)
                        Else
                            'Can only get size of picture if it is in the internal images collection
                            If Not Grid.UseImageList Then sngExtra = ScaleXCoor(ScalingObj, Grid.Images(.Picture).Picture.Width, vbHimetric, ScaleMode)
                        End If
                    End If
                End With
                
                str = objCell.GetText(Column.MaskDisplayMode)
    
                sngMax = Max(sngMax, ScalingObj.TextWidth(str))
                End If
            Next objRow
        Next RowScrollRegion
    End If
    
    If bHasButton Then
        'Button size is based on the system's scrollbar size with 4 extra pixels for "padding"
        sngExtra = ScaleXCoor(ScalingObj, 2 + GetSystemMetrics(SM_CXHSCROLL), vbPixels, ScaleMode)
    End If
    
    'Include spacing set aside for cellpadding and cellspacing
    With Column.Band.ResolveOverride
        sngExtra = sngExtra + Abs(.CellPadding) * 2!
        sngExtra = sngExtra + Abs(.CellSpacing) * 2!
    End With
    
    'Set aside space for column border
    sngExtra = sngExtra + ScaleXCoor(ScalingObj, 2, vbPixels, ScaleMode)
    
    'Some extra "padding" from edit window itself
    sngExtra = sngExtra + ScaleXCoor(ScalingObj, 4, vbPixels, ScaleMode)

    sngMax = sngMax + sngExtra
    
    If IncludeColHeader Then
        Set ScalingObj.Font = Column.Header.ResolveAppearance.Font  'Take into account header caption?
        str = Column.Header.Caption
        
        sngHeader = ScalingObj.TextWidth(str)
        Set uiHeader = Column.Header.GetUIElement                   'Try to get the UIElement for the header
        sngExtra = ScaleXCoor(ScalingObj, 4, vbPixels, ScaleMode)   'Header not visible so assume a 4 pixel padding?
        With Column.Band.ResolveOverride
            Select Case .HeaderClickAction
                Case ssHeaderClickActionSortMulti, ssHeaderClickActionSortSingle
                    Select Case Column.SortIndicator
                        Case ssSortIndicatorDescending, ssSortIndicatorAscending
                            sngExtra = sngExtra + ScaleXCoor(ScalingObj, 11, vbPixels, ScaleMode)
                    End Select
            End Select
            
            'If column swapping is allow then we must add a couple of pixels for the column swapping button
            Select Case .AllowColSwapping
                Case ssAllowColSwappingWithinBand, ssAllowColSwappingWithinGroup
                    sngExtra = sngExtra + ScaleXCoor(ScalingObj, 11, vbPixels, ScaleMode)
            End Select
        End With
            
        If Not uiHeader Is Nothing Then
            'Take away any extra space in the header from the recordselectors
            sngExtra = sngExtra - (ScaleXCoor(ScalingObj, uiHeader.Rect.Width, vbPixels, ScaleMode) - Column.Width)
            sngExtra = Max(sngExtra, 0!)
        End If
        
        sngHeader = sngHeader + sngExtra
    
        'Only store the maximum width
        sngMax = Max(sngMax, sngHeader)
    End If
    
    If Not AllowSizeDecrease Then
        'If the calculated size is smaller than the current size then only resize if allowed to do so
        If sngMax < Column.Width Then Exit Sub
    End If
    
    Column.Width = sngMax
End Sub
Public Function ColHeaderSizeArea(ByVal X As Single, ByVal Y As Single, ByVal Grd As UltraGrid.SSUltraGrid, _
        ByVal ScalingObj As Object, ByVal ScaleMode As VBRUN.ScaleModeConstants) As UltraGrid.SSUIElement
    Dim sngX As Single
    Dim sngY As Single
    Dim pt As POINTAPI
    Dim UIElement As UltraGrid.SSUIElement
    
    'Determines if you are over the area in between column headers that
    'is used to size columns. If so, it returns the uielement of the
    'leftmost header - i.e. the one that would be sized.
    
    sngX = X
    sngY = Y
    
    Set UIElement = Grd.UIElementFromPoint(sngX, sngY)
    
    pt.X = ScaleXCoor(ScalingObj, X, ScaleMode, vbPixels)
    pt.Y = ScaleYCoor(ScalingObj, Y, ScaleMode, vbPixels)
    
    'If no uielement at those coordinates then exit
    If UIElement Is Nothing Then Exit Function
    
    'Make sure we can get to a header object
    If Not UIElement.CanResolveUIElement(ssUIElementHeader) Then
        'Check to see if they are doubleclicking on the last column
        'header... shift 2 pixels to the left and try again?
        sngX = ScaleXCoor(ScalingObj, pt.X - 3, vbPixels, ScaleMode)
        'sngY = ScaleYCoor(ScalingObj, pt.Y, vbPixels, ScaleMode)
        
        Set UIElement = Grd.UIElementFromPoint(sngX, sngY)
        
        If UIElement Is Nothing Then Exit Function
        If UIElement.Header Is Nothing Then Exit Function
    End If
            
    Set UIElement = UIElement.ResolveUIElement(ssUIElementHeader)
    
    'If there is no header object available then exit
    If UIElement.Header Is Nothing Then Exit Function
    
    'If the header is not for a column then exit
    If UIElement.Header.Type <> ssHeaderTypeColumn Then Exit Function
    
    If pt.X >= UIElement.Rect.Right - 2 And pt.X <= UIElement.Rect.Right Then
        'If the pt is within 2 pixels of the right edge of the column header, do nothing
    ElseIf pt.X >= UIElement.Rect.Left And pt.X <= UIElement.Rect.Left + 2 Then
        'when you are within two pixels of the left border rect, then shift to the column to the left
        sngX = ScaleXCoor(ScalingObj, pt.X - 3, vbPixels, ScaleMode)
        'sngY = ScaleYCoor(ScalingObj, pt.Y, vbPixels, ScaleMode)
        
        Set UIElement = Grd.UIElementFromPoint(sngX, sngY)
    Else
        Exit Function
    End If
        
    If UIElement Is Nothing Then Exit Function
    If UIElement.Header Is Nothing Then Exit Function
    Set UIElement = UIElement.ResolveUIElement(ssUIElementHeader)
    
    Set ColHeaderSizeArea = UIElement
End Function
Private Function ScaleXCoor(ByVal Obj As Object, ByVal Value As Single, ByVal ScaleFrom As VBRUN.ScaleModeConstants, ByVal ScaleTo As VBRUN.ScaleModeConstants) As Single
    If ScaleFrom = ScaleTo Then
        ScaleXCoor = Value
    ElseIf ScaleFrom = vbPixels And ScaleTo = vbTwips Then
        ScaleXCoor = Value * CSng(Screen.TwipsPerPixelX)
    ElseIf ScaleTo = vbPixels And ScaleFrom = vbTwips Then
        ScaleXCoor = Value / CSng(Screen.TwipsPerPixelX)
    ElseIf Not Obj Is Nothing Then
        On Error Resume Next
        ScaleXCoor = Obj.ScaleX(Value, ScaleFrom, ScaleTo)
    Else
        'do nothing
    End If
End Function
Private Function ScaleYCoor(ByVal Obj As Object, ByVal Value As Single, ByVal ScaleFrom As VBRUN.ScaleModeConstants, ByVal ScaleTo As VBRUN.ScaleModeConstants) As Single
    If ScaleFrom = ScaleTo Then
        ScaleYCoor = Value
    ElseIf ScaleFrom = vbPixels And ScaleTo = vbTwips Then
        ScaleYCoor = Value * CSng(Screen.TwipsPerPixelY)
    ElseIf ScaleTo = vbPixels And ScaleFrom = vbTwips Then
        ScaleYCoor = Value / CSng(Screen.TwipsPerPixelY)
    ElseIf Not Obj Is Nothing Then
        On Error Resume Next
        ScaleYCoor = Obj.ScaleY(Value, ScaleFrom, ScaleTo)
    Else
        'do nothing
    End If
End Function
Private Function LoWord(DWord As Long) As Integer
   If DWord And &H8000& Then
      LoWord = DWord Or &HFFFF0000
   Else
      LoWord = DWord And &HFFFF&
   End If
End Function
Private Function HiWord(DWord As Long) As Integer
   HiWord = (DWord And &HFFFF0000) \ &H10000
End Function
Private Function Max(First As Single, Second As Single) As Single
    If First > Second Then
        Max = First
    Else
        Max = Second
    End If
End Function
Private Function Min(First As Single, Second As Single) As Single
    If First < Second Then
        Min = First
    Else
        Min = Second
    End If
End Function



