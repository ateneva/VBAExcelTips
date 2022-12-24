Attribute VB_Name = "EventsWks_BeforeRightClick"
Option Explicit

Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)

'a handy way of adding an URL to a PivotTable
    ActiveWorkbook.FollowHyperlink Address:=Target.Value, NewWindow:=True

End Sub
