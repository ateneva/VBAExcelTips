Attribute VB_Name = "CellInR_ModifyDates"
Option Explicit

Sub ReplaceActiveCellYear()
Dim Cell As Range

'replace active cell year but retain the rest of the date
Dim d As Integer
Dim m As Integer
Dim y As Integer

For Each Cell In Range("E2:E" & ActiveSheet.UsedRange.Rows.Count)
Cell.Activate

        d = Day(ActiveCell)
        m = Month(ActiveCell)
        y = Year(ActiveCell)

        If y > 2015 Then ActiveCell.Formula = DateValue(d & "/" & m & "/2015")

Next Cell
End Sub

Sub ReplaceDate1MonthBack()
Dim Cell As Range

For Each Cell In Range("E2:E" & ActiveSheet.UsedRange.Rows.Count)
        If Cell.Value > Date Then Cell.Value = Application.WorksheetFunction.EoMonth(Date, -1)

Next Cell
End Sub
