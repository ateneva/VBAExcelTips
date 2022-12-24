Attribute VB_Name = "FY15_Project_Slides_Delivery"
Option Explicit

Sub Projects_Profile()

Dim Sh As Shape
Dim PPApp As PowerPoint.Application
Dim PPpres As PowerPoint.Presentation
Dim PPSlide As PowerPoint.Slide

'Create a PP application and make it visible
Set PPApp = New PowerPoint.Application
PPApp.Visible = msoCTrue

'Open the presentation you wish to copy to
Set PPpres = PPApp.Presentations.Open("C:\Users\tenevaa\Documents\TS EMEA\I am Responsible For\Delivery Packages\Projects.pptm")

'''needed if you're using Excel 2013 to prevent PowerPoint from losing focus and returning
'"shapes (unknown member) invalid request. the specified data type is unavailable
'- Run-time error -2147188160 (80048240):View (unknown member) error
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
PPApp.Activate
PPApp.ActiveWindow.ViewType = ppViewNormal
PPApp.ActiveWindow.Panes(2).Activate
'********************************************************************************************************************************

Application.Calculation = xlCalculationAutomatic

Worksheets("projects SR profile").Activate

'copy CEE data
ActiveSheet.Range("B44:L81").Copy
PPpres.Slides(2).Shapes.PasteSpecial ppPasteEnhancedMetafile

ActiveSheet.Range("B407:J415").Copy 'Total Activity
PPpres.Slides(1).Shapes.PasteSpecial ppPasteEnhancedMetafile
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'copy FRA data
ActiveSheet.Range("B83:L120").Copy
PPpres.Slides(5).Shapes.PasteSpecial ppPasteEnhancedMetafile

'copy FRA data
ActiveSheet.Range("B417:J425").Copy
PPpres.Slides(4).Shapes.PasteSpecial ppPasteEnhancedMetafile
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'copy GER data
ActiveSheet.Range("B122:L159").Copy
PPpres.Slides(8).Shapes.PasteSpecial ppPasteEnhancedMetafile

'copy GER data
ActiveSheet.Range("B427:J435").Copy
PPpres.Slides(7).Shapes.PasteSpecial ppPasteEnhancedMetafile
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'copy GWE data
ActiveSheet.Range("B161:L198").Copy
PPpres.Slides(11).Shapes.PasteSpecial ppPasteEnhancedMetafile

'copy GWE data
ActiveSheet.Range("B437:J445").Copy
PPpres.Slides(10).Shapes.PasteSpecial ppPasteEnhancedMetafile
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'copy IBE data
ActiveSheet.Range("B200:L237").Copy
PPpres.Slides(14).Shapes.PasteSpecial ppPasteEnhancedMetafile

'copy IBE data
ActiveSheet.Range("B447:J455").Copy
PPpres.Slides(13).Shapes.PasteSpecial ppPasteEnhancedMetafile
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'copy ITA data
ActiveSheet.Range("B239:L276").Copy
PPpres.Slides(17).Shapes.PasteSpecial ppPasteEnhancedMetafile

'copy ITA data
ActiveSheet.Range("B457:J465").Copy
PPpres.Slides(16).Shapes.PasteSpecial ppPasteEnhancedMetafile
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'copy MEMA data
ActiveSheet.Range("B278:L315").Copy
PPpres.Slides(20).Shapes.PasteSpecial ppPasteEnhancedMetafile

'copy MEMA data
ActiveSheet.Range("B467:J475").Copy
PPpres.Slides(19).Shapes.PasteSpecial ppPasteEnhancedMetafile
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'copy RUS data
ActiveSheet.Range("B317:L354").Copy
PPpres.Slides(26).Shapes.PasteSpecial ppPasteEnhancedMetafile

'copy RUS data
ActiveSheet.Range("B477:J485").Copy
PPpres.Slides(25).Shapes.PasteSpecial ppPasteEnhancedMetafile
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'copy UKI data
ActiveSheet.Range("B356:L393").Copy
PPpres.Slides(23).Shapes.PasteSpecial ppPasteEnhancedMetafile

'copy UKI data
ActiveSheet.Range("B487:J495").Copy
PPpres.Slides(22).Shapes.PasteSpecial ppPasteEnhancedMetafile
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Application.CutCopyMode = False
PPpres.Save
PPpres.Close

're-name the HP logo to "Picture 1"
Worksheets("new projects").Activate
ActiveSheet.Shapes.Range(Array("Picture 2")).Select
Selection.name = "Picture 1"

Worksheets("completed projects").Activate
ActiveSheet.Shapes.Range(Array("Picture 2")).Select
Selection.name = "Picture 1"

MsgBox ("Format as Table and create live picture for new projects & completed projects")

End Sub

Sub New_Completed_Projects()

Dim Sh As Shape
Dim PPApp As PowerPoint.Application
Dim PPpres As PowerPoint.Presentation
Dim PPSlide As PowerPoint.Slide

'Create a PP application and make it visible
Set PPApp = New PowerPoint.Application
PPApp.Visible = msoCTrue

'Open the presentation you wish to copy to
Set PPpres = PPApp.Presentations.Open("C:\Users\tenevaa\Documents\TS EMEA\I am Responsible For\Delivery Packages\Projects.pptm")

'''needed if you're using Excel 2013 to prevent PowerPoint from losing focus and returning
'"shapes (unknown member) invalid request. the specified data type is unavailable
'- Run-time error -2147188160 (80048240):View (unknown member) error
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
PPApp.Activate
PPApp.ActiveWindow.ViewType = ppViewNormal
PPApp.ActiveWindow.Panes(2).Activate
'********************************************************************************************************************************

Application.Calculation = xlCalculationAutomatic

'get new projects
Worksheets("new projects").Activate

'copy CEE
ActiveSheet.ListObjects("Table3").Range.AutoFilter Field:=7, Criteria1:="CEE&I"
ActiveSheet.Shapes.Range(Array("Picture 2")).Select
Selection.Copy
PPpres.Slides(3).Shapes.PasteSpecial ppPasteEnhancedMetafile

'copy FRA
ActiveSheet.ListObjects("Table3").Range.AutoFilter Field:=7  'clears previously applied filters in that field

ActiveSheet.ListObjects("Table3").Range.AutoFilter Field:=7, Criteria1:="FRA"
ActiveSheet.Shapes.Range(Array("Picture 2")).Select
Selection.Copy
PPpres.Slides(6).Shapes.PasteSpecial ppPasteEnhancedMetafile

'copy GER
ActiveSheet.ListObjects("Table3").Range.AutoFilter Field:=7  'clears previously applied filters in that field

ActiveSheet.ListObjects("Table3").Range.AutoFilter Field:=7, Criteria1:="GER"
ActiveSheet.Shapes.Range(Array("Picture 2")).Select
Selection.Copy
PPpres.Slides(9).Shapes.PasteSpecial ppPasteEnhancedMetafile

'copy GWE
ActiveSheet.ListObjects("Table3").Range.AutoFilter Field:=7  'clears previously applied filters in that field

ActiveSheet.ListObjects("Table3").Range.AutoFilter Field:=7, Criteria1:="GWE"
ActiveSheet.Shapes.Range(Array("Picture 2")).Select
Selection.Copy
PPpres.Slides(12).Shapes.PasteSpecial ppPasteEnhancedMetafile

'copy IBE
ActiveSheet.ListObjects("Table3").Range.AutoFilter Field:=7  'clears previously applied filters in that field

ActiveSheet.ListObjects("Table3").Range.AutoFilter Field:=7, Criteria1:="IBE"
ActiveSheet.Shapes.Range(Array("Picture 2")).Select
Selection.Copy
PPpres.Slides(15).Shapes.PasteSpecial ppPasteEnhancedMetafile

'copy ITA
ActiveSheet.ListObjects("Table3").Range.AutoFilter Field:=7  'clears previously applied filters in that field

ActiveSheet.ListObjects("Table3").Range.AutoFilter Field:=7, Criteria1:="ITA"
ActiveSheet.Shapes.Range(Array("Picture 2")).Select
Selection.Copy
PPpres.Slides(18).Shapes.PasteSpecial ppPasteEnhancedMetafile

'copy MEMA
ActiveSheet.ListObjects("Table3").Range.AutoFilter Field:=7  'clears previously applied filters in that field

ActiveSheet.ListObjects("Table3").Range.AutoFilter Field:=7, Criteria1:="MEMA"
ActiveSheet.Shapes.Range(Array("Picture 2")).Select
Selection.Copy
PPpres.Slides(21).Shapes.PasteSpecial ppPasteEnhancedMetafile

'copy UKI
ActiveSheet.ListObjects("Table3").Range.AutoFilter Field:=7  'clears previously applied filters in that field

ActiveSheet.ListObjects("Table3").Range.AutoFilter Field:=7, Criteria1:="UKI"
ActiveSheet.Shapes.Range(Array("Picture 2")).Select
Selection.Copy
PPpres.Slides(24).Shapes.PasteSpecial ppPasteEnhancedMetafile

'copy RUS
ActiveSheet.ListObjects("Table3").Range.AutoFilter Field:=7  'clears previously applied filters in that field

ActiveSheet.ListObjects("Table3").Range.AutoFilter Field:=7, Criteria1:="RUS"
ActiveSheet.Shapes.Range(Array("Picture 2")).Select
Selection.Copy
PPpres.Slides(27).Shapes.PasteSpecial ppPasteEnhancedMetafile

PPpres.Save

'*******************************************************************************
'get completed projects
Worksheets("completed projects").Activate

'copy CEE
ActiveSheet.ListObjects("Table4").Range.AutoFilter Field:=7, Criteria1:="CEE&I"
ActiveSheet.Shapes.Range(Array("Picture 2")).Select
Selection.Copy
PPpres.Slides(3).Shapes.PasteSpecial ppPasteEnhancedMetafile

'copy FRA
ActiveSheet.ListObjects("Table4").Range.AutoFilter Field:=7  'clears previously applied filters in that field

ActiveSheet.ListObjects("Table4").Range.AutoFilter Field:=7, Criteria1:="FRA"
ActiveSheet.Shapes.Range(Array("Picture 2")).Select
Selection.Copy
PPpres.Slides(6).Shapes.PasteSpecial ppPasteEnhancedMetafile

'copy GER
ActiveSheet.ListObjects("Table4").Range.AutoFilter Field:=7  'clears previously applied filters in that field

ActiveSheet.ListObjects("Table4").Range.AutoFilter Field:=7, Criteria1:="GER"
ActiveSheet.Shapes.Range(Array("Picture 2")).Select
Selection.Copy
PPpres.Slides(9).Shapes.PasteSpecial ppPasteEnhancedMetafile

'copy GWE
ActiveSheet.ListObjects("Table4").Range.AutoFilter Field:=7  'clears previously applied filters in that field

ActiveSheet.ListObjects("Table4").Range.AutoFilter Field:=7, Criteria1:="GWE"
ActiveSheet.Shapes.Range(Array("Picture 2")).Select
Selection.Copy
PPpres.Slides(12).Shapes.PasteSpecial ppPasteEnhancedMetafile

'copy IBE
ActiveSheet.ListObjects("Table4").Range.AutoFilter Field:=7  'clears previously applied filters in that field

ActiveSheet.ListObjects("Table4").Range.AutoFilter Field:=7, Criteria1:="IBE"
ActiveSheet.Shapes.Range(Array("Picture 2")).Select
Selection.Copy
PPpres.Slides(15).Shapes.PasteSpecial ppPasteEnhancedMetafile

'copy ITA
ActiveSheet.ListObjects("Table4").Range.AutoFilter Field:=7  'clears previously applied filters in that field

ActiveSheet.ListObjects("Table4").Range.AutoFilter Field:=7, Criteria1:="ITA"
ActiveSheet.Shapes.Range(Array("Picture 2")).Select
Selection.Copy
PPpres.Slides(18).Shapes.PasteSpecial ppPasteEnhancedMetafile

'copy MEMA
ActiveSheet.ListObjects("Table4").Range.AutoFilter Field:=7  'clears previously applied filters in that field

ActiveSheet.ListObjects("Table4").Range.AutoFilter Field:=7, Criteria1:="MEMA"
ActiveSheet.Shapes.Range(Array("Picture 2")).Select
Selection.Copy
PPpres.Slides(21).Shapes.PasteSpecial ppPasteEnhancedMetafile

'copy UKI
ActiveSheet.ListObjects("Table4").Range.AutoFilter Field:=7  'clears previously applied filters in that field

ActiveSheet.ListObjects("Table4").Range.AutoFilter Field:=7, Criteria1:="UKI"
ActiveSheet.Shapes.Range(Array("Picture 2")).Select
Selection.Copy
PPpres.Slides(24).Shapes.PasteSpecial ppPasteEnhancedMetafile

'copy RUS
ActiveSheet.ListObjects("Table4").Range.AutoFilter Field:=7  'clears previously applied filters in that field

ActiveSheet.ListObjects("Table4").Range.AutoFilter Field:=7, Criteria1:="RUS"
ActiveSheet.Shapes.Range(Array("Picture 2")).Select
Selection.Copy
PPpres.Slides(27).Shapes.PasteSpecial ppPasteEnhancedMetafile

PPpres.Save

End Sub
