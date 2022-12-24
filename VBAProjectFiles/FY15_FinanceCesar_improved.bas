Attribute VB_Name = "FY15_FinanceCesar_improved"
Option Explicit

Sub Select_Actual()
Attribute Select_Actual.VB_ProcData.VB_Invoke_Func = "S\n14"

Dim Msg As Integer, Ans As Integer

Dim Sh As Shape
Dim PPApp As PowerPoint.Application
Dim PPpres As PowerPoint.Presentation
Dim PPSlide As PowerPoint.Slide

Dim Country As Range
Dim Cell As Range

'Create a PP application and make it visible
Set PPApp = New PowerPoint.Application
PPApp.Visible = msoCTrue

'Open the presentation you wish to copy to
Set PPpres = PPApp.Presentations.Open("C:\Users\hp\Desktop\TS EMEA Delivery Finance Package.pptm")

'**********************************************************************************************
'prevent PowerPoint from losing focus and returning
'"shapes (unknown member) invalid request. the specified data type is unavailable"
'- Run-time error -2147188160 (80048240):View (unknown member) error
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
PPApp.Activate
PPApp.ActiveWindow.ViewType = ppViewNormal
PPApp.ActiveWindow.Panes(2).Activate 'standard ppt view
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Set Country = ActiveSheet.Range("G7")

'***************************
For Each Cell In ActiveSheet.Range("BW8:BW24")
    Country = Cell.Value
    
    ActiveSheet.Shapes.Range(Array("Picture 1")).Select
    Selection.Copy

    Select Case Country
        Case "Europe": PPpres.Slides(2).Shapes.PasteSpecial ppPasteEnhancedMetafile
        Case "UK": PPpres.Slides(6).Shapes.PasteSpecial ppPasteEnhancedMetafile
        Case "Germany": PPpres.Slides(10).Shapes.PasteSpecial ppPasteEnhancedMetafile
        Case "IBERIA": PPpres.Slides(14).Shapes.PasteSpecial ppPasteEnhancedMetafile
        Case "Italy": PPpres.Slides(18).Shapes.PasteSpecial ppPasteEnhancedMetafile
        Case "France": PPpres.Slides(22).Shapes.PasteSpecial ppPasteEnhancedMetafile
        Case "GWE": PPpres.Slides(26).Shapes.PasteSpecial ppPasteEnhancedMetafile
        Case "MEMA": PPpres.Slides(30).Shapes.PasteSpecial ppPasteEnhancedMetafile
        Case "CEE": PPpres.Slides(34).Shapes.PasteSpecial ppPasteEnhancedMetafile
        Case "Russia": PPpres.Slides(2).Shapes.PasteSpecial ppPasteEnhancedMetafile
    End Select

Next Cell

PPpres.Save
PPpres.Close

End Sub

Sub Select_Flash()
Attribute Select_Flash.VB_ProcData.VB_Invoke_Func = "Q\n14"

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

'************************************************************
'''needed if you're using Excel 2013 to prevent PowerPoint from losing focus and returning
'"shapes (unknown member) invalid request. the specified data type is unavailable"
'- Run-time error -2147188160 (80048240):View (unknown member) error
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
PPApp.Activate
PPApp.ActiveWindow.ViewType = ppViewNormal
PPApp.ActiveWindow.Panes(2).Activate 'standard ppt view
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'***********************************************************************
Select Case Country

Case "Europe"
ActiveSheet.Shapes.Range(Array("Picture 1")).Select
Selection.Copy
PPpres.Slides(3).Shapes.PasteSpecial ppPasteEnhancedMetafile

Case "UK & I"
ActiveSheet.Shapes.Range(Array("Picture 1")).Select
Selection.Copy
PPpres.Slides(7).Shapes.PasteSpecial ppPasteEnhancedMetafile

Case "Germany"
ActiveSheet.Shapes.Range(Array("Picture 1")).Select
Selection.Copy
PPpres.Slides(11).Shapes.PasteSpecial ppPasteEnhancedMetafile

Case "IBERIA"
ActiveSheet.Shapes.Range(Array("Picture 1")).Select
Selection.Copy
PPpres.Slides(15).Shapes.PasteSpecial ppPasteEnhancedMetafile

Case "Italy"
ActiveSheet.Shapes.Range(Array("Picture 1")).Select
Selection.Copy
PPpres.Slides(19).Shapes.PasteSpecial ppPasteEnhancedMetafile

Case "France"
ActiveSheet.Shapes.Range(Array("Picture 1")).Select
Selection.Copy
PPpres.Slides(23).Shapes.PasteSpecial ppPasteEnhancedMetafile

Case "GWE"
ActiveSheet.Shapes.Range(Array("Picture 1")).Select
Selection.Copy
PPpres.Slides(27).Shapes.PasteSpecial ppPasteEnhancedMetafile

Case "MEMA"
ActiveSheet.Shapes.Range(Array("Picture 1")).Select
Selection.Copy
PPpres.Slides(31).Shapes.PasteSpecial ppPasteEnhancedMetafile

Case "CEE"
ActiveSheet.Shapes.Range(Array("Picture 1")).Select
Selection.Copy
PPpres.Slides(35).Shapes.PasteSpecial ppPasteEnhancedMetafile

Case "Russia"
ActiveSheet.Shapes.Range(Array("Picture 1")).Select
Selection.Copy
PPpres.Slides(3).Shapes.PasteSpecial ppPasteEnhancedMetafile

End Select

PPpres.Save
PPpres.Close

End Sub
