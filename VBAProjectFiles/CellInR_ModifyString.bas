Attribute VB_Name = "CellInR_ModifyString"
Option Explicit

Sub ReplaceACharInString()

Dim Cell As Range

With ActiveSheet
For Each Cell In Range("I2:I14")

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

Sub AddARandomNumberAfterString()
Dim Cell As Range
Dim prv As String
'~~~~~~~~~~~~~~~~~~~~~~
For Each Cell In ActiveSheet.Range("D2:D" & ActiveSheet.UsedRange.Rows.Count)

        prv = Cell.Value
        Cell.Value = prv & Chr(44) & Application.WorksheetFunction.RandBetween(0, 9)

Next Cell
End Sub

Sub CompileNewStringfromExistingOnes()

Dim Cell As Range
For Each Cell In ActiveSheet.Range("S3:S" & ActiveSheet.UsedRange.Rows.Count)
    
    Cell.Value = Cell.Offset(0, -6).Value & ", " & Left(Cell.Offset(0, -7), 1) & "."

Next Cell
End Sub

Sub ExtractString()

Dim Cell As Range
'~~~~~~~~~~~~~~~~~~~~~~~~~~~

'extract the first 8 characters of each cell located in column 14
For Each Cell In ActiveSheet.Range("U3:U" & ActiveSheet.UsedRange.Rows.Count)
    Cell.Value = Left(Cells(Cell.row, 14), 8)
Next Cell

'extract the last 8 characters of each cell located in column 14
For Each Cell In ActiveSheet.Range("U3:U" & ActiveSheet.UsedRange.Rows.Count)
    Cell.Value = Right(Cells(Cell.row, 14), 8)
Next Cell

End Sub

Sub StringCase()

Dim Cell As Range

For Each Cell In ActiveSheet.Range("T3:T" & ActiveSheet.UsedRange.Rows.Count)
    Cell.Value = LCase(Cell)                                    'lower case
    Cell.Value = UCase(Cell)                                    'upper case
    Cell.Value = Application.WorksheetFunction.Proper(Cell)     'proper case Proper does not exist as VBA function
Next Cell

End Sub
