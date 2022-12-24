Attribute VB_Name = "CellInR_PreventScientificFormat"
Option Explicit

Sub preventscientific()

Dim Cell As Range
Dim textid As Variant
Dim newid As String
'----------------------------

For Each Cell In ActiveSheet.Range("B2:B" & ActiveSheet.UsedRange.Rows.Count).SpecialCells(xlCellTypeVisible)
    
    textid = Cell.Value
    newid = "'" & textid
    Cell.Value = newid

Next Cell

End Sub
