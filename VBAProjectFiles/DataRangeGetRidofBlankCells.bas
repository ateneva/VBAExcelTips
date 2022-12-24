Attribute VB_Name = "DataRangeGetRidofBlankCells"
Option Explicit

Sub GetRidofBlankCells()

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

Sub GetRidOfUnunsedRange()
'~written by Angelina Teneva
'get rid of unused range
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ThisWorkbook.Worksheets("CATIS").Activate

With ActiveSheet
    Rows("1:1").AutoFilter
    Range("A2").End(xlDown).Offset(1, 0).Select                        'reaches the last populated cell and goes to the next row
    Range(ActiveCell, ActiveCell.SpecialCells(xlLastCell)).Rows.Delete 'uses the curently active cell and goes to the last one of the range
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


