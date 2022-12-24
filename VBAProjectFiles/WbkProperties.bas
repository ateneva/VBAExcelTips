Attribute VB_Name = "WbkProperties"
Option Explicit

Sub GetFileSize()

MsgBox "The current workbook size is " & Round(FileLen(ActiveWorkbook.FullName) / 1048576, 2) & " MB"

End Sub

Sub GetRidOfUnUsedRange()

Dim Wks As Worksheet
Dim i As Integer
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva, September 2016
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each Wks In ActiveWorkbook.Worksheets
    If Wks.Index > 18 Then Wks.Activate
    
        With ActiveSheet
            'reaches the last populated cell and goes to the next row
            Range("A2").End(xlDown).Offset(1, 0).Select
            
            'uses the curently active cell and goes to the last one of the range
            Range(ActiveCell, ActiveCell.SpecialCells(xlLastCell)).EntireRow.Delete
        End With
    
Next

ActiveWorkbook.Save
End Sub
