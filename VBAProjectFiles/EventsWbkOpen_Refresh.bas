Attribute VB_Name = "EventsWbkOpenRefresh"
Option Explicit

Private Sub Workbook_Open_Refresh()

If Application.UserName = "Angelina" Then
    ThisWorkbook.RefreshAll
    ThisWorkbook.saveas "\\VBOXSVR\Virtual_Machine_\Dashboards\Plarium_Weekly\plarium_conversion_source_" & Format(Date, "yyyymmdd") & ".xlsb"
    ThisWorkbook.saveas "C:\Users\Angelina\Dashboards\weeklyreports\plarium_conversion_source_" & Format(Date, "yyyymmdd") & ".xlsb"
    ThisWorkbook.Close
End If
End Sub


