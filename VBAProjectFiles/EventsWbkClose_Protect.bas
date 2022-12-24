Attribute VB_Name = "EventsWbkCloseProtect"
Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)

Dim Wks As Worksheet
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'adds protection and hides sheets

For Each Wks In ThisWorkbook.Worksheets
    'drawing objects = False allows the editing of Slicers
    If Wks.name = "Slicers" Or Wks.name = "Targets" Or Wks.name = "QTD" Or Wks.name = "data" Then
    
        Wks.Protect ("inhead"), DrawingObjects:=False, Contents:=True, _
            Scenarios:=True, AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:=True
        
        Wks.Visible = xlSheetVeryHidden
    End If
    
    'drawing objects = False allows the editing of Slicers
    If Wks.Visible = True Then Wks.Protect ("inhead"), DrawingObjects:=False, Contents:=True, _
        Scenarios:=True, AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:=True

Next Wks

End Sub

