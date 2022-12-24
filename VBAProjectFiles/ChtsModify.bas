Attribute VB_Name = "ChtsModify"
Option Explicit

Sub ModifyChartSheet()
Dim Cht As Chart    'chart sheets

For Each Cht In ActiveWorkbook.Charts
    Cht.Type = xlArea
    Cht.Type = xlColumnClustered
Next Cht

End Sub

Sub FormatAllCharts()
Dim Cht As Chart
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
For Each Cht In ActiveWorkbook.Charts

With Cht
    .ChartType = xlLineMarkers
    .ApplyLayout 3
    .ChartStyle = 12
    .ClearToMatchStyle
    .SetElement msoElementChartTitleAboveChart
    .SetElement msoElementLegendNone
    .SetElement msoElementPrimaryValueAxisTitleNone
    .SetElement msoElementPrimaryCategoryAxisTitleNone
    .Axes(xlValue).MinimumScale = 0
    .Axes(xlValue).MaximumScale = 1000
End With

Next Cht
End Sub
