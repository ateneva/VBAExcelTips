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

			'export pivot tables on PowerPoint
					.PivotTables("Total").PivotSelect "", xlDataAndLabel, True

					Selection.Copy
					PPpres.Slides(2).Shapes.PasteSpecial ppPasteEnhancedMetafile 'picture with no background and good resolution
				'*************************************************************
	End With

PPpres.Save
PPpres.Close

End Sub
