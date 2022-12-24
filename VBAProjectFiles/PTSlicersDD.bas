Attribute VB_Name = "PTSlicersDD"
Option Explicit

Sub Worksheet_PivotTableUpdate_Slicers()
'code by Debra Dalgleish

Dim Wb As Workbook
Dim scShort As SlicerCache
Dim scLong As SlicerCache

Dim siShort As SlicerItem
Dim siLong As SlicerItem

On Error GoTo errHandler
Application.ScreenUpdating = False
Application.EnableEvents = False

Set Wb = ThisWorkbook
Set scShort = Wb.SlicerCaches("Slicer_City")
Set scLong = Wb.SlicerCaches("Slicer_City1")

scLong.ClearManualFilter

For Each siLong In scLong.VisibleSlicerItems
    Set siLong = scLong.SlicerItems(siLong.name)
    Set siShort = Nothing
    
    On Error Resume Next
    Set siShort = scShort.SlicerItems(siLong.name)
    
    On Error GoTo errHandler
    If Not siShort Is Nothing Then
    
        If siShort.Selected = True Then
            siLong.Selected = True
        ElseIf siShort.Selected = False Then
            siLong.Selected = False
        End If
    Else
        siLong.Selected = False
    End If
Next siLong

exitHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

errHandler:
    MsgBox "Could not update pivot table"
    Resume exitHandler

End Sub


