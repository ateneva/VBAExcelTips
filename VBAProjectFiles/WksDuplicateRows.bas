Attribute VB_Name = "WksDuplicateRows"
Option Explicit

Sub DuplicateRows()
Dim Cell As Range

'~inserts a row after each value currently present
'1st cell with number of tickets

Set Cell = Range("B2")
Do While Not IsEmpty(Cell)

    If Cell > 1 Then
    Range(Cell.Offset(1, 0), Cell.Offset(Cell.Value - 1, 0)).EntireRow.Insert
    Range(Cell, Cell.Offset(Cell.Value - 1, 1)).EntireRow.FillDown
    End If

Set Cell = Cell.Offset(Cell.Value, 0)

Loop
End Sub
