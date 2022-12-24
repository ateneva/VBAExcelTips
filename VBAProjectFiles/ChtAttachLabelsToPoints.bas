Attribute VB_Name = "ChtAttachLabelsToPoints"

Sub AttachLabelsToPoints()

   'Dimension variables.
   Dim Counter As Integer, ChartName As String, xVals As String

   ' Disable screen updating while the subroutine is run.
   Application.ScreenUpdating = False

   'Store the formula for the first series in "xVals".
   xVals = ActiveChart.SeriesCollection(1).Formula

   'Extract the range for the data from xVals.
   xVals = Mid(xVals, InStr(InStr(xVals, ","), xVals, _
      Mid(Left(xVals, InStr(xVals, "!") - 1), 9)))
   xVals = Left(xVals, InStr(InStr(xVals, "!"), xVals, ",") - 1)
   Do While Left(xVals, 1) = ","
      xVals = Mid(xVals, 2)
   Loop

   'Attach a label to each data point in the chart.
   For Counter = 1 To Range(xVals).Cells.Count
     ActiveChart.SeriesCollection(1).Points(Counter).HasDataLabel = _
         True
      ActiveChart.SeriesCollection(1).Points(Counter).DataLabel.text = _
         Range(xVals).Cells(Counter, 1).Offset(0, -1).Value
   Next Counter

End Sub

Sub DataLabelsFromRange() 'replace DataLabels with items from text

Dim DLRange As Range
Dim Cht As Chart
Dim i As Integer, Pts As Integer
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Specify chart
Set Cht = ActiveSheet.ChartObjects(1).Chart

'Prompt for a range
On Error Resume Next
Set DLRange = Application.InputBox(Prompt:="Range for data labels?", Type:=8)
If DLRange Is Nothing Then Exit Sub
On Error GoTo 0

'Add data labels
Cht.SeriesCollection(1).ApplyDataLabels Type:=xlDataLabelsShowValue, AutoText:=True, LegendKey:=False

'Loop through the Points, and set the data labels
Pts = Cht.SeriesCollection(1).Points.Count

For i = 1 To Pts
    Cht.SeriesCollection(1).Points(i).DataLabel.text = DLRange(i)
Next i

End Sub
