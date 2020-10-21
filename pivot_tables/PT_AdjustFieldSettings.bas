
Sub FieldSettings()

Dim Wks As Worksheet
Dim PT As PivotTable
Dim PF As PivotField
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva, Oct 2016
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each Wks In ActiveWorkbook.Worksheets

Wks.Activate

    For Each PT In ActiveSheet.PivotTables
        If PT.Name Like "date_stats" Or PT.Name Like "chart_dates" Then
            Set PF = PT.PivotFields("Date")

            PF.ShowDetail = True                'expands the field; if applied on a PivotItem = double click
            PF.ShowAllItems = True              'shows items with no data
            PF.RepeatLabels = True              'repeat item labels
            PF.LayoutBlankLine = True           'inserts a blank line
            PF.IncludeNewItemsInFilter = True   'ensure the pivot field filter picks up new values

        End If
    Next PT
Next Wks

End Sub
