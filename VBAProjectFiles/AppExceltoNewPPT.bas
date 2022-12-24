Attribute VB_Name = "AppExcelToNewPPT"
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
