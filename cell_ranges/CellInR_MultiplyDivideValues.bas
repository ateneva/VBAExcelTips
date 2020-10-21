
Sub MultiplyValues()
Dim Cell As Range
Dim prv As Double

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva
'https://datageeking.wordpress.com/2017/05/11/quickly-restate-values-with-vba/
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'multiplies by a 1000
For Each Cell In ActiveSheet.Range("J2:AX" & ActiveSheet.UsedRange.Rows.Count).SpecialCells(xlCellTypeVisible)
  prv = Cell.Value
  Cell.Value = prv * 1000 '*(-1)
Next Cell

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'make values negative
For Each Cell In ActiveSheet.Range("AV:AV", "AX:AX")

  If IsNumeric(Cell) And Cell.Value <> 0 Then
    prv = Cell.Value

    Cell.Formula = prv * -1
  End If
Next Cell

End Sub
