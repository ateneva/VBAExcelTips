Attribute VB_Name = "PTFieldSettings"
Option Explicit

Sub FieldSettings()

Dim PT As PivotTable
Dim PF As PivotField

For Each PT In ActiveSheet.PivotTables

    'Set PF = PT1.PivotFields("Date")

'****************************PT.PF field formats*************************************************************************
PT.PivotFields("Country").ShowDetail = True         'expands the field; if applied on a PivotItem = double click
PT.PivotFields("Country").ShowAllItems = True       'shows items with no data
PT.PivotFields("Element").RepeatLabels = True       'repeat item labels
PT.PivotFields("Element").LayoutBlankLine = True    'inserts a blank line
PT.PivotFields("Element").IncludeNewItemsInFilter = True

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~extracting dataset in a new tab~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
monthly = Format(Application.WorksheetFunction.EoMonth(Date, -2) + 1, "dd/mm/yyyy") 'returns the 1st day of a month in "dd/mmm/yyy" format
Set PI = PF.PivotItems(monthly)

PI.DataRange.End(xlDown).Select
Selection.ShowDetail = True                         'when applied on a pivotitem = double-click
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

PT.PF.Subtotals(1) = True                           'show subtotals for a field
PT.PF.Subtotals(1) = False
PT.PivotFields("Location").Subtotals = Array(True, False, False, False, False, False, False, False, False, False, False, False)

'.DataBodyRange ---------> Object in the PivotTable
'.DataRange -------------> Object in the PivotField and PivotItems

PT.PivotFields("Activity Hours").NumberFormat = "0,0;;" 'formats a single field
PT.PivotFields("Activity Hours").NumberFormat = "0.00%"

'formats only numbers associated with a PivotField
PT.PivotFields("Region").DataRange.NumberFormat = "$#,##0"
PT.PivotFields("Region").DataRange.NumberFormat = "$#,##0"
PT.PivotFields("Region").DataRange.NumberFormat = "0.00%"

'formats only numbers associated with a DataItem
PT.PivotFields("Region").PivotItems("Europe").DataRange.NumberFormat = "$#,##0"
PT.PivotFields("Region").PivotItems("North America").DataRange.NumberFormat = "$#,##0"
PT.PivotFields("Region").PivotItems("Europe/N America").DataRange.NumberFormat = "0.00%"


Next PT
End Sub



