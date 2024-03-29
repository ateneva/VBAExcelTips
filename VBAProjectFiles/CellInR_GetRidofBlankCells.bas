Attribute VB_Name = "CellInR_GetRidofBlankCells"
Option Explicit

Sub FillInEmptyCellWithConstant()

Dim i As Long
Dim Col As Integer
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Col = 7 To 68
    For i = 1 To Cells(Rows.Count, Col).End(xlUp).row
        If IsEmpty(Cells(i, Col)) Then Cells(i, Col).Value = "empty"
    Next i
Next Col

End Sub

Sub FillInEmptyCellWithPrevious()

Dim Cell As Range
Dim Region As Range
Set Region = Worksheets("Extract").Range(Cells(3, 1), Cells(ActiveSheet.UsedRange.Rows.Count, 3, 3))
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each Cell In Region

If IsEmpty(Cell.Value) Then Cell.Offset(-1, 0).Select  'if cell is empty then put the value from the previous cell
    Selection.Copy ActiveCell.Offset(1, 0)

    'If Cell.Value = "" Then Cell.Offset(-1, 0).Select
        'Selection.Copy ActiveCell.Offset(1, 0) -----> formulated in this way, the code will take into account
        '"IFERROR(VLOOKUP(),"")" generated results into account

Next Cell

End Sub

Sub FillInEmptyCellWithAdjacent()

Dim Cell As Range

'copy to column
For Each Cell In ActiveSheet.Range("AB2:AB" & ActiveSheet.UsedRange.Rows.Count)

Cell.Activate
On Error Resume Next

    If ActiveCell.Value <> 0 And Cells(ActiveCell.row, 25) = 0 Then ActiveCell.Copy Cells(ActiveCell.row, 25)
    If IsEmpty(ActiveCell) Or ActiveCell.text = "" And Cells(ActiveCell.row, 44).text <> "#N/A" Then ActiveCell.Value = Cells(ActiveCell.row, 44)

Next Cell

End Sub

Sub GetRidOfUnunsedRange()
'~written by Angelina Teneva
'get rid of unused range
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ThisWorkbook.Worksheets("data").Activate

With ActiveSheet
    Rows("1:1").AutoFilter
    
    'reaches the last populated cell and goes to the next row
     Range("A2").End(xlDown).Offset(1, 0).Select
    
    'uses the curently active cell and goes to the last one of the range
    Range(ActiveCell, ActiveCell.SpecialCells(xlLastCell)).EntireRow.Delete
    
End With

ThisWorkbook.Save

End Sub

Sub DeleteEmptyRows()
Dim LastRow As Long
Dim r As Long
Dim Counter As Long

'written by John Waleknbach

Application.ScreenUpdating = False
LastRow = ActiveSheet.UsedRange.Rows.Count + ActiveSheet.UsedRange.Rows(1).row - 1

For r = LastRow To 1 Step -1
    If Application.WorksheetFunction.CountA(Rows(r)) = 0 Then
    Rows(r).Delete
    Counter = Counter + 1
End If
Next r

Application.ScreenUpdating = True
MsgBox Counter & " empty rows were deleted."
End Sub

Sub ReplacePreviousValueWithNextOne()

Dim Cell As Range

'copy to column
For Each Cell In ActiveSheet.Range("H2:H40")

Cell.Activate
On Error Resume Next

    If IsDate(Cell) = True Then ActiveCell.Copy Cells(ActiveCell.row - 1, 8)
    
        'OR (both produce the same result
        
    If IsDate(Cell) = True Then ActiveCell.Copy Cell.Offset(-1, 0)

Next Cell

End Sub
