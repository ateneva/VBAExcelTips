Attribute VB_Name = "CellInR_CleanUpString"
Option Explicit

Sub ReplaceACharInString()

Dim Cell As Range

With ActiveSheet
For Each Cell In Range("A2:A" & ActiveSheet.UsedRange.Rows.Count).SpecialCells(xlCellTypeVisible)

Cell.Activate
 'blank character was in the middle
 'VBA built-in would not work
 'however the WS function wouild

ActiveCell.Value = Trim(ActiveCell)
ActiveCell.Value = Application.WorksheetFunction.Trim(ActiveCell)

If InStr(Cell, "n") = 1 Then Cell.Replace "n", ""

Next Cell

End With
End Sub
