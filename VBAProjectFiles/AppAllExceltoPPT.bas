Attribute VB_Name = "AppAllExceltoPPT"
 Option Explicit

Sub ExcelRangeToNewPowerPoint()

Dim PPApp As PowerPoint.Application

'Set PPApp = CreateObject("Powerpoint.Application") 'always creates a new instance of the object
Set PPApp = GetObject(, "Powerpoint.Application")  'uses the instance that is currently active
'Set PPT = GetObject("C:\Myfile.pptx") 'used to access a file that's already loaded

Dim PPpres As PowerPoint.Presentation 'this is an example of Early Binding --> PowerPoint Object Model must be referenced
Set PPpres = PPApp.ActivePresentation

Dim PPSlide As PowerPoint.Slide
Dim Wks As Worksheet

Dim Ans As Integer
'************************************************************
'''needed if you're using Excel 2013 to prevent PowerPoint from losing focus and returning
'"shapes (unknown member) invalid request. the specified data type is unavailable"
'- Run-time error -2147188160 (80048240):View (unknown member) error
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
PPApp.Activate
PPApp.ActiveWindow.ViewType = ppViewNormal
PPApp.ActiveWindow.Panes(2).Activate 'standard ppt view
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Worksheets("YTD").Columns("Y:DP").EntireColumn.Hidden = True
Worksheets("YTD").Range("B1:E31,DQ1:EK31").Copy
PPpres.Slides(73).Shapes.PasteSpecial ppPasteEnhancedMetafile ''paste Excel Range on PowerPoint

Worksheets("EMEA").Range("D1:AK31").Copy
PPpres.Slides(10).Shapes.PasteSpecial ppPasteEnhancedMetafile ''paste Excel Range on PowerPoint

'copy CEE
Range("B4:O21").SpecialCells(xlCellTypeVisible).Copy
PPpres.Slides(36).Shapes.PasteSpecial ppPasteEnhancedMetafile ''paste the visible cells of the Excel Range on PowerPoint

' Save the presentation
PPpres.Save

Application.CutCopyMode = False
End Sub

Sub ExcelToExistingPowerPoint()

Dim PPApp As PowerPoint.Application
Dim PPpres As PowerPoint.Presentation
Dim PPS As Integer

'~~~~~~~~~~~~~~Early binding~~~~~~~requires referencing the PowerPoint Object Model first
'Create a PP application and make it visible
Set PPApp = New PowerPoint.Application
PPApp.Visible = msoCTrue

'Open the presentation you wish to copy to
Set PPpres = PPApp.Presentations.Open("C:\Users\Angelina\Documents\Balance.pptm")
'****************************************************************************************************************************

'''needed if you're using Excel 2013 to prevent PowerPoint from losing focus and returning
'"shapes (unknown member) invalid request. the specified data type is unavailable"
'- Run-time error -2147188160 (80048240):View (unknown member) error
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
PPApp.Activate
PPApp.ActiveWindow.ViewType = ppViewNormal
PPApp.ActiveWindow.Panes(2).Activate
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ThisWorkbook.Worksheets("Balance").Activate
With ActiveSheet
Range("A1:N4").Copy

For PPS = 2 To 12 Step 2
PPpres.Slides(PPS).Shapes.PasteSpecial ppPasteEnhancedMetafile
Next PPS

'''export pivot tables on PowerPoint
.PivotTables("Total").PivotSelect "", xlDataAndLabel, True
Selection.Copy
PPpres.Slides(2).Shapes.PasteSpecial ppPasteEnhancedMetafile 'picture with no background and good resolution
'*************************************************************

.PivotTables("Monthly").PivotSelect "", xlDataAndLabel, True
Selection.Copy
PPpres.Slides(12).Shapes.PasteSpecial ppPasteEnhancedMetafile 'picture with no background and good resolution
'**************************************************************

ActiveSheet.PivotTables(1).PivotSelect "", xlDataAndLabel, True 'the pivot table that Excel considers first on a worksheet
Selection.Copy
PPpres.Slides(1).Shapes.PasteSpecial ppPasteHTML 'pastes the data in a format which allows the selection of data items from PowerPoint

End With
Application.CutCopyMode = False
PPpres.Save
PPpres.Close

End Sub

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

Sub ExportChartOrRange()

Dim PPApp As PowerPoint.Application
Dim PPpres As PowerPoint.Presentation
Dim Ans As Integer
Dim inputdata As String

'Create a PP application and make it visible
Set PPApp = New PowerPoint.Application
PPApp.Visible = msoCTrue

'Open the presentation you wish to copy to
Set PPpres = PPApp.Presentations.Open("C:\Users\Angelina\Documents\Slides.pptm")

'''needed if you're using Excel 2013 to prevent PowerPoint from losing focus and returning
'"shapes (unknown member) invalid request. the specified data type is unavailable"
'- Run-time error -2147188160 (80048240):View (unknown member) error
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
PPApp.Activate
PPApp.ActiveWindow.ViewType = ppViewNormal
PPApp.ActiveWindow.Panes(2).Activate
'****************************************************************************************************************************

Ans = MsgBox("Would you like to copy chart", vbYesNo)

Select Case Ans
Case vbYes
ActiveSheet.ChartObjects("Chart 1").Copy
PPpres.Slides(32).Shapes.PasteSpecial ppPasteEnhancedMetafile

Case vbNo
On Error GoTo notification

inputdata = InputBox("Enter the range you want to copy")
Range(inputdata).Copy
PPpres.Slides(32).Shapes.PasteSpecial ppPasteEnhancedMetafile
End Select

notification:
MsgBox ("Range not selected")

Application.CutCopyMode = False
End Sub

Sub ExportCameraToolPics()

Dim Msg As Integer, Ans As Integer

Dim Sh As Shape
Dim PPApp As PowerPoint.Application
Dim PPpres As PowerPoint.Presentation
Dim PPSlide As PowerPoint.Slide

Dim Country As Range
Set Country = Worksheets("Geography C$").Range("G7")

'Create a PP application and make it visible
Set PPApp = New PowerPoint.Application
PPApp.Visible = msoCTrue

'Open the presentation you wish to copy to
Set PPpres = PPApp.Presentations.Open("C:\Users\tenevaa\Documents\TS EMEA\I am Responsible For\Delivery Packages\TS EMEA Delivery Finance Package.pptm")

'''needed if you're using Excel 2013 to prevent PowerPoint from losing focus and returning
'"shapes (unknown member) invalid request. the specified data type is unavailable"
'- Run-time error -2147188160 (80048240):View (unknown member) error
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
PPApp.Activate
PPApp.ActiveWindow.ViewType = ppViewNormal
PPApp.ActiveWindow.Panes(2).Activate
'**********************************************************************************************
Select Case Country

Case "Europe"
ActiveSheet.Shapes.Range(Array("Picture 1")).Select
Selection.Copy
PPpres.Slides(2).Shapes.PasteSpecial ppPasteEnhancedMetafile

Case "UK & I"
ActiveSheet.Shapes.Range(Array("Picture 1")).Select
Selection.Copy
PPpres.Slides(6).Shapes.PasteSpecial ppPasteEnhancedMetafile

Case "Germany"
ActiveSheet.Shapes.Range(Array("Picture 1")).Select
Selection.Copy
PPpres.Slides(10).Shapes.PasteSpecial ppPasteEnhancedMetafile

Case "France"
ActiveSheet.Shapes.Range(Array("Picture 1")).Select
Selection.Copy
PPpres.Slides(22).Shapes.PasteSpecial ppPasteEnhancedMetafile

Case "GWE"
ActiveSheet.Shapes.Range(Array("Picture 1")).Select
Selection.Copy
PPpres.Slides(26).Shapes.PasteSpecial ppPasteEnhancedMetafile

End Select

PPpres.Save
PPpres.Close

End Sub



