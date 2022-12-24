Attribute VB_Name = "PTEvents"
Option Explicit

Private Sub Worksheet_PivotTableChangeSync(ByVal Target As PivotTable)
 
'The PivotTableChangeEvent occurs during most changes to a PivotTable; '
'code responds to user actions, such as clearing, grouping, re-naming or refreshing items in the PivotTable.

Set Target = ActiveSheet.PivotTables(1)
Range("G1").Value = Target.RefreshDate
Range("H1").Value = Target.RefreshName
Range("I1").Value = Now() + 1

End Sub

 Private Sub Worksheet_PivotTableAfterValueChange(TargetPivotTable, TargetRange)
 
 'The PivotTableAfterValueChange event does not occur under any conditions other than editing or recalculating cells.
 'For example, it will not occur when the PivotTable is refreshed, sorted, filtered, or drilled down on,
 'even though those operations move cells and potentially retrieve new values from the OLAP data source.
 
 End Sub
 
Private Sub Worksheet_PivotTableUpdate(ByVal Target As PivotTable)
  
'runs after a pivot table has been changed, refreshed or re-filtered
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Set Target = ActiveSheet.PivotTables(1)

Range("G1").Value = Target.RefreshDate
Range("H1").Value = Target.RefreshName
  
End Sub
