Attribute VB_Name = "ChtAllGeneral"
Option Explicit

Sub MoveChart()

ActiveSheet.ChartObjects(1).Chart.Location xlLocationAsNewSheet, "MyChart" 'move As ChartSheet
Charts("MyChart").Location xlLocationAsObject, "Sheet1" 'move as Embedded Object

End Sub

Private Sub Chart_Select(ByVal ElementID As Long, _
ByVal Arg1 As Long, ByVal Arg2 As Long)
Dim Id As String
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'executed whenever a chart element is selected

Select Case ElementID
Case xlAxis: Id = "Axis"
Case xlAxisTitle: Id = "AxisTitle"
Case xlChartArea: Id = "ChartArea"
Case xlChartTitle: Id = "ChartTitle"
Case xlCorners: Id = "Corners"
Case xlDataLabel: Id = "DataLabel"
Case xlDataTable: Id = "DataTable"
Case xlDownBars: Id = "DownBars"
Case xlDropLines: Id = "DropLines"
Case xlErrorBars: Id = "ErrorBars"
Case xlFloor: Id = "Floor"
Case xlHiLoLines: Id = "HiLoLines"
Case xlLegend: Id = "Legend"
Case xlLegendEntry: Id = "LegendEntry"
Case xlLegendKey: Id = "LegendKey"
Case xlMajorGridlines: Id = "MajorGridlines"
Case xlMinorGridlines: Id = "MinorGridlines"
Case xlNothing: Id = "Nothing"
Case xlPlotArea: Id = "PlotArea"
Case xlRadarAxisLabels: Id = "RadarAxisLabels"
Case xlSeries: Id = "Series"
Case xlSeriesLines: Id = "SeriesLines"
Case xlShape: Id = "Shape"
Case xlTrendline: Id = "Trendline"
Case xlUpBars: Id = "UpBars"
Case xlWalls: Id = "Walls"
Case xlXErrorBars: Id = "XErrorBars"
Case xlYErrorBars: Id = "YErrorBars"
Case Else:: Id = "Some unknown thing"
End Select

MsgBox "Selection type" & Id & vbCrLf & Arg1 & vbCrLf & Arg2
End Sub

Sub SaveChartAsPNG()
Dim fname As String
'~~~~~~~~~~~~~~~~~~~~~~
If ActiveChart Is Nothing Then Exit Sub

fname = ThisWorkbook.path & " \ " & ActiveChart.name & ".png"
ActiveChart.Export FileName:=fname, FilterName:="PNG"

End Sub
