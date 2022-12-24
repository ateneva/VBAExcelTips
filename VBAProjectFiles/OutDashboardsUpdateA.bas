Attribute VB_Name = "OutDashboardsUpdateA"
Option Explicit

Sub UpdateDashboard()

Dim SourceWbk As String
SourceWbk = "\\VBOXSVR\Virtual_Machine_\Dashboards\Amplify\amplify_dashboard_" & Format(Date, "yyyymmdd") & ".csv"

Dim DayValue As Integer
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'delete the data for the 15th week back
With ThisWorkbook.Worksheets
    Worksheets("data").Activate
    
    With ActiveSheet
        .ListObjects("QTDbyW").Range.AutoFilter Field:=3, Criteria1:=Application.WorksheetFunction.WeekNum(Date) - 16
        
        Rows("2:2").Select
        Range(Selection, Selection.End(xlDown)).SpecialCells(xlCellTypeVisible).Select
        Selection.Delete Shift:=xlUp
        .ShowAllData
    End With
End With

'open latest .csv
Workbooks.Open FileName:=SourceWbk, ReadOnly:=True, UpdateLinks:=False
ActiveWorkbook.Worksheets(1).Activate
    With ActiveSheet
    Range(Cells(2, 1), Cells(ActiveSheet.UsedRange.Rows.Count, 24)).Copy
    End With

'add the data from the latest .csv
ThisWorkbook.Worksheets("data").Activate
    With ActiveSheet
        Cells(1, 1).End(xlDown).Offset(1, 0).Select
        .Paste
    End With
       
'update the day of the month on Targets tab
ThisWorkbook.Worksheets("Targets").Activate
    With ActiveSheet
        If DayValue < 91 Then
           DayValue = Cells(1, 10).Value
           Cells(1, 10).Value = DayValue + 7
        Else: Cells(1, 10).Value = 1
        End If
        
     Cells(16, 1).Value = "amplify_dashboard_" & Format(Date, "yyyymmdd") & ".csv"
    End With
    
Windows(Worksheets("Targets").Range("A16").Value).Activate
Application.CutCopyMode = False
ActiveWorkbook.Close savechanges:=False

'refresh pivottables
ThisWorkbook.RefreshAll
ThisWorkbook.Worksheets("Account View").Activate
    With ActiveSheet
    .PivotTables("PivotTable10").PivotFields("AdvertiserName").EnableMultiplePageItems = False
    .PivotTables("PivotTable10").PivotFields("AdvertiserName").CurrentPage = Worksheets("Amplify Dashboard").Cells(19, 14).Value
    End With
    
ThisWorkbook.Worksheets("Amplify Dashboard").Activate
ThisWorkbook.SaveAs "\\VBOXSVR\Virtual_Machine_\Dashboards\Amplify\Global - Amplify Business Performance Dashboard Q3_" & Format(Date, "yyyymmdd") & ".xlsb"

End Sub
