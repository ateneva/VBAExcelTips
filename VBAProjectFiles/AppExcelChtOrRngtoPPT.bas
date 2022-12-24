Attribute VB_Name = "AppExcelChtOrRngtoPPT"
Option Explicit

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

