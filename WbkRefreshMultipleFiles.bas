
Sub RefreshMonthly()

Dim Cell As Range
Dim Source As String
Dim Target As String

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'The code below assumes that the workbook has active database connections
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Application.ScreenUpdating = False
Application.DisplayAlerts = False

For Each Cell In ThisWorkbook.Worksheets("monthly").Range("A2:A" & Worksheets("monthly").UsedRange.Rows.count)
    Source = Cell.Value
    Target = Cell.Offset(0, 2).Value
    Workbooks.Open FileName:=Source, ReadOnly:=False, UpdateLinks:=False

    With ActiveWorkbook
        .RefreshAll
        .saveas Target
        .Close
    End With
Next Cell

MsgBox ("Refresh Completed")
Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub
