Attribute VB_Name = "ChtObjModify"
Option Explicit

Sub FormatAllChartObjects()

Dim ChtObj As ChartObject
For Each ChtObj In ActiveSheet.ChartObjects
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
With ChtObj.Chart
    .ChartType = xlLine
    .ChartType = xlLineStacked
    .ChartType = xlLineMarkers
    .ChartType = xlArea
    .ChartType = xlBarClustered
    .ChartType = xlBarStacked
    .ChartType = xlBubble
    .ChartType = xlColumnClustered
    .ChartType = xlColumnStacked
    .ChartType = xlDoughnut
    .ChartType = xlPie
    .ChartType = xlRadar
    
    'scatter charts --------------> values on both x and y axes
    .ChartType = xlXYScatter
    .ChartType = xlXYScatterLines
    .ChartType = xlXYScatterLinesNoMarkers
    .ChartType = xlXYScatterSmooth
    .ChartType = xlXYScatterSmoothNoMarkers
        
    .HasTitle = True
    .ChartTitle.text = "YTD Sales"
    .ChartTitle.HorizontalAlignment
                
    .ApplyLayout 3
    .ChartStyle = 12
    .ClearToMatchStyle
    .SetElement msoElementChartTitleAboveChart                  'title above chart
    .SetElement msoElementLegendNone                            'no legend
    .SetElement msoElementPrimaryValueAxisTitleNone             'no value axis title
    .SetElement msoElementPrimaryCategoryAxisTitleNone          'no category axis title
    
    'if you do not set .HasTitle to True before specifiying the AxisTitle object, an error occurs
    .Axes(xlValue).HasTitle = True
    .Axes(xlValue).AxisTitle.text = "Sales by Region"
    .Axes(xlValue).AxisTitle.Font.name = "Arial"
    .Axes(xlValue).AxisTitle.Size = 10
    
    .Axes(xlCategory).HasTitle = True
    .Axes(xlCategory).AxisTitle.text = "Month"
    .Axes(xlCategory).AxisTitle.Font.name = "Arial"
    .Axes(xlCategory).AxisTitle.Size = 10
    
    'setting the minimum scales
    .Axes(xlValue).MinimumScale = 0
    .Axes(xlValue).MaximumScale = 1000
    
    .ChartArea.Font.name = "Calibri"
    .ChartArea.Font.FontStyle = "Regular"
    .ChartArea.Font.Size = 9
           
    .PlotArea.Interior.ColorIndex = xlNone
    .Axes(xlValue).TickLabels.Font.Bold = True
    .Axes(xlCategory).TickLabels.Font.Bold = True
    .Legend.Position = xlBottom
    
End With

With ChtObj
    .BringToFront
    .Copy
    .CopyPicture
    .Cut
    .Delete
    .Duplicate
    .name
    .SendToBack
End With

Next ChtObj
End Sub

Sub ModifyChartObjectExample()
Dim ChtObj As ChartObject                                   'embedded charts
Dim i As Integer
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
For Each ChtObj In ActiveSheet.ChartObjects

    ChtObj.Chart.Type = xlArea                              'changes to area chart
    ChtObj.Chart.Type = xlColumnClustered                   'changes toscolumnclustered chart
    
    'modifies chart Title elements
    With ChtObj.Chart.ChartTitle
        .text = "YTD Sales"
        .Font.name = "Arial"
        .Size = 14
        .Color = RGB(0, 0, 255)
    End With
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'changing the font of the Legend
        With ChtObj.Chart.Legend.Font
            .name = "Calibri"
            .FontStyle = "Bold"
            .Size = 12
        End With

    'modify the Values Axis (y)
    With .Axes(xlValue)
        .HasTitle = True
        .AxisTitle.text = "Sales by Region"
        .AxisTitle.Font.name = "Arial"
        .AxisTitle.Size = 10
    End With
    
    'modify the Categories Axis (x)
    With .Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.text = "Sales by Region"
        .AxisTitle.Font.name = "Arial"
        .AxisTitle.Size = 10
    End With
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'other possible formats
'You might need to activate a chart before executing the ChartMods macro.
'Activate an embedded chart by clicking it.
'To activate a chart on a chart sheet, activate the chart sheet.

'ActiveChart.Type = xlArea
'ActiveChart.ChartArea.Font.name = "Calibri"
'ActiveChart.ChartArea.Font.FontStyle = "Regular"
'ActiveChart.ChartArea.Font.Size = 9
'
'ActiveChart.PlotArea.Interior.ColorIndex = xlNone
'ActiveChart.Axes(xlValue).TickLabels.Font.Bold = True
'ActiveChart.Axes(xlCategory).TickLabels.Font.Bold = True
'ActiveChart.Legend.Position = xlBottom

Next Cht

End Sub

Sub SizeAndAlignChartObjects()

Dim W As Long, H As Long
Dim TopPosition As Long, LeftPosition As Long
Dim ChtObj As ChartObject
Dim i As Long, NumCols As Long
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
If ActiveChart Is Nothing Then
    MsgBox "Select a chart to be used as the base for the sizing"
Exit Sub
End If

'Get Columns
On Error Resume Next
NumCols = InputBox("How many columns of charts?")

If Err.Number <> 0 Then Exit Sub
If NumCols < 1 Then Exit Sub
On Error GoTo 0

'Get size of active chart
W = ActiveChart.Parent.width
H = ActiveChart.Parent.height

'Change starting positions, if necessary
TopPosition = 100
LeftPosition = 20

For i = 1 To ActiveSheet.ChartObjects.count
    With ActiveSheet.ChartObjects(i)
        .width = W
        .height = H
        .Left = LeftPosition + ((i - 1) Mod NumCols) * W
        .Top = TopPosition + Int((i - 1) / NumCols) * H
    End With
Next i

End Sub

Sub CopyEmbeddedChartsToNewSheet(name As String, width As Integer, height As Integer)

'Copies all embedded charts in the current workbook to a new worksheet with the specified name.
'The copied charts have the specified width and height and are arranged in a single column.
'The vertical space between charts.

Const SPACE_BETWEEN_CHARTS = 20
Dim newWS As Worksheet
Dim oldWS As Worksheet
Dim co As ChartObject
Dim yPos As Integer
Dim count As Integer

' Turn screen updating off so screen does not flicker as charts are copied.
Application.ScreenUpdating = False
Set newWS = Worksheets.Add.name = "Dashboard"

For Each oldWS In Worksheets

    'Do not copy from the new worksheet.
    If oldWS.name <> name Then
        For Each co In oldWS.ChartObjects
            co.Copy newWS.Range("A1").Select
            newWS.Paste
        Next co
    End If
    
Next oldWS

'Position and size the charts.
count = 0
For Each co In newWS.ChartObjects
    co.width = width
    co.height = height
    co.Left = 30
    co.Top = count * (height + SPACE_BETWEEN_CHARTS) + SPACE_BETWEEN_CHARTS
    
count = count + 1
Next

' Turn screen updating back on.
Application.ScreenUpdating = True

End Sub
