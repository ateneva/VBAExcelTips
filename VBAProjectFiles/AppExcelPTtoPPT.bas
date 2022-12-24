Attribute VB_Name = "AppExcelPTtoPPT"
Option Explicit

Sub ExcelToPowerPoint_Open()

Dim PPApp As PowerPoint.Application
Dim PPpres As PowerPoint.Presentation
Dim PPS As Integer

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

