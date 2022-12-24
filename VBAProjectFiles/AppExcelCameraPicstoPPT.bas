Attribute VB_Name = "AppExcelCameraPicstoPPT"
Option Explicit

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




