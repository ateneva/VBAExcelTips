Attribute VB_Name = "ChtCreate"
Option Explicit

'Application --> Workbook --> Worksheet --> ChartObject --> Chart --> ChartTitle

Sub CreateChartSheet()

Dim ch As Chart
Set ch = ActiveWorkbook.Charts.Add

ch.SetSourceData Source:=Worksheets("Sheet1").Range("B3:F6"), PlotBy:=xlRows
ch.ChartType = xlLineMarkers

End Sub

Sub CreateEmbeddedChart()

Dim co As ChartObject
Dim ch As Chart

Set co = Worksheets(“Sheet1”).ChartObjects.Add(50, 100, 250, 165)
Set ch = co.Chart
ch.SetSourceData Source:=Worksheets("Sheet1").Range("B3:F6"), PlotBy:=xlRows

End Sub

Sub CreateAChartFromSelectedRange()
Dim ChartData As Range
Dim ChartShape As Shape
Dim NewChart As Chart
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Create chart from selected range
Set ChartData = ActiveWindow.RangeSelection
Set ChartShape = ActiveSheet.Shapes.AddChart
Set NewChart = ChartShape.Chart


'Adjust the chart
With NewChart
.ChartType = xlColumnClustered
.SetSourceData Source:=Range(ChartData.Address)
.Legend.Delete
.SeriesCollection(1).Format.Shadow.Type = msoShadow21
End With

End Sub

Sub CreateAChartTypes()
Dim MyChart As Chart
Dim MyChartSht As Chart
Dim DataRange As Range

'creates chart as an embedded object
Set DataRange = ActiveSheet.Range("A1:C7")
Set MyChart = ActiveSheet.Shapes.AddChart.Chart 'creates the chart
MyChart.SetSourceData Source:=DataRange 'adds data to it

'creates a chart as a chart sheet
Set MyChartSht = Charts.Add
MyChartSht.SetSourceData Source:=DataRange
ActiveChart.ChartType = xlColumnClustered

End Sub

Sub AddAScatterChart()

Dim ch As Chart
Set co = Worksheets("Sheet1").ChartObjects.Add(50, 200, 250, 165)
Set ch = co.Chart

'A scatter chart (sometimes called an XY chart) is fundamentally different from other types of Excel charts.
'Most charts plot values against categories.

'In contrast, a scatter chart plots values versus values;
'therefore, both the horizontal and vertical axes have values on them.

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ch.SetSourceData Source:=Worksheets("Sheet1").Range("B3:D10"), PlotBy:=xlColumns
    
    'define scatter type
    ch.ChartType = xlXYScatter
    ch.ChartType = xlXYScatterLines
    ch.ChartType = xlXYScatterLinesNoMarkers
    ch.ChartType = xlXYScatterSmooth
    ch.ChartType = xlXYScatterSmoothNoMarkers
    
    
    'add title
    ch.HasTitle = True
    ch.ChartTitle.text = Worksheets("Sheet1").Range("A1").Value

'Add a category axis title.
With ch.Axes(xlCategory)
    .HasTitle = True
    .AxisTitle.text = "Processor Speed (GHz)"
End With

'Add a value axis title.
With ch.Axes(xlValue)
    .HasTitle = True
    .AxisTitle.text = "Units Sold"
End With

End Sub

