Attribute VB_Name = "EventsWks_ChangeSpecificCell"
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)

Dim PT As PivotTable
Dim PF As PivotField

'**************************************************************************
'Very handy way of filtering all PivotTables on a Worksheet via dropdown
'**************************************************************************

'filter for network benchmark
'*******************************
If Target.Address = "$B$5" Then

For Each PT In ActiveSheet.PivotTables
    Set PF = PT.PivotFields("CASE")
        PF.ClearAllFilters
        PF.PivotFilters.Add Type:=xlCaptionContains, Value1:=Range("B5").Value
    Next PT
End If

'filter for platform
'*******************************
If Target.Address = "$D$5" Then

For Each PT In ActiveSheet.PivotTables
    If Range("D5").Value <> "All" Then
        Set PF = PT.PivotFields("Platform")
            PF.ClearAllFilters
            PF.PivotFilters.Add Type:=xlCaptionContains, Value1:=Range("D5").Value
    Else
    
        Set PF = PT.PivotFields("Platform")
        PF.ClearAllFilters
    End If
Next PT

End If

End Sub

