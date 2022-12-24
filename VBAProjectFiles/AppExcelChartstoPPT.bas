Attribute VB_Name = "AppExcelChartstoPPT"
Option Explicit

Sub ExportCharts()

Dim PPApp As PowerPoint.Application
Dim PPpres As PowerPoint.Presentation

Dim pptx As String
pptx = ThisWorkbook.Worksheets("calculated fields").Range("F2")

'Create a PP application and make it visible
Set PPApp = New PowerPoint.Application
PPApp.Visible = msoCTrue

'Open the presentation you wish to copy to
Set PPpres = PPApp.Presentations.Open(pptx)

'''needed if you're using Excel 2013 to prevent PowerPoint from losing focus and returning
'"shapes (unknown member) invalid request. the specified data type is unavailable"
'- Run-time error -2147188160 (80048240):View (unknown member) error
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
PPApp.Activate
PPApp.ActiveWindow.ViewType = ppViewNormal
PPApp.ActiveWindow.Panes(2).Activate

'********************************************************************************
'copy chartsheets ('ThisWorkbook.Chart = when the chart is in a separate sheet)
ThisWorkbook.Charts("EMEA").ChartArea.Copy
PPpres.Slides(3).Shapes.PasteSpecial ppPasteEnhancedMetafile

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'copy charts from Dashboard tab (This is .ChartObjects collection
'and Object must always be activated first)

Worksheets("Dashboard").Activate
With ActiveSheet

Range("C8").Value = "EMEA"
.ChartObjects("all").Activate
ActiveChart.ChartArea.Copy
PPpres.Slides(4).Shapes.PasteSpecial ppPasteEnhancedMetafile

End With
PPpres.Save
End Sub
