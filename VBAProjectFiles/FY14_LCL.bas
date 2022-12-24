Attribute VB_Name = "FY14_LCL"
Option Explicit

Sub LCL_Actuals_14()

Dim PT As PivotTable
Dim PF As PivotField
'***************************

With ActiveWorkbook
Worksheets("Overview Pivot A Vs T").Activate
Columns("N:AA").ColumnWidth = 10.57

For Each PT In ActiveSheet.PivotTables

On Error Resume Next
PT.PivotFields("Date").PivotItems("(blank)").Visible = False 'excludes blanks
PT.PivotFields("Service Line Description").Position = 1
PT.PivotFields("Location").Position = 2
PT.PivotFields("Year").PivotItems("FY13").ShowDetail = True

PT.PivotFields("Sum of Target Spend in US$").Orientation = xlHidden 'hide Target Spend Field

'Sum of Remaining Target is a calculated field.
'Orientation property does not work for calculated fields. You need to delete them
For Each PF In PT.CalculatedFields
PF.Delete
Next PF
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
If Month(Date) = 12 Or Month(Date) = 1 Or Month(Date) = 2 Then PT.PivotFields("Quarter").PivotItems("Q1FY13").ShowDetail = True
If Month(Date) >= 3 And Month(Date) <= 5 Then PT.PivotFields("Quarter").PivotItems("Q2FY13").ShowDetail = True
If Month(Date) >= 6 And Month(Date) <= 8 Then PT.PivotFields("Quarter").PivotItems("Q3FY13").ShowDetail = True
If Month(Date) >= 9 And Month(Date) <= 11 Then PT.PivotFields("Quarter").PivotItems("Q4FY13").ShowDetail = True

'remove sub-region totals
PT.PivotFields("Subregion").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    
'add subototals for Location
PT.PivotFields("Location").Subtotals = Array(True, False, False, False, False, False, False, False, False, False, _
        False, False)
        
'add product line sub-totals
PT.PivotFields("Service Line Description").Subtotals = Array(True, False, False, False, False, False, False, False, False, False, _
        False, False)

PT.PivotSelect "Farshore", xlDataAndLabel, True  'color farshore items differently
    With Selection.Interior
        .pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    
PT.PivotSelect "", xlDataAndLabel, True
Selection.Copy
Range("C100").Select
ActiveSheet.Paste
   
Next PT

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'create a 2nd pivot table
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each PT In ActiveSheet.PivotTables

If PT.name <> "ActualVsTarget" Then
PT.PivotFields("Subregion").Position = 1
PT.PivotFields("Subregion").ShowAllItems = True
PT.PivotFields("Subregion").Subtotals = _
        Array(True, False, False, False, False, False, False, False, False, False, False, False)

PT.PivotFields("Location").Position = 2
PT.name = "ActualVsTarget2"

End If
Next PT
End With

Application.Run "Prepare_LCL_slide_v02"

End Sub

Sub Prepare_LCL_slide_v02()

Dim PT As PivotTable
Dim PT1 As PivotTable
Dim PF As PivotField
'*******************************
Dim PPApp As PowerPoint.Application
Dim PPpres As PowerPoint.Presentation
Dim PPSlide As PowerPoint.Slide
Dim Today As Date

'Create a PP application and make it visible
Set PPApp = New PowerPoint.Application
PPApp.Visible = msoCTrue

'Open the presentation you wish to copy to
Set PPpres = PPApp.Presentations.Open("C:\Users\TENEVAA\Documents\TS EMEA\I am Responsible For\Scorecards\COS and Residual.pptm")
'****************************************************************************************************************************

With ActiveWorkbook
Worksheets("Overview Tables").Activate
ActiveWindow.DisplayGridlines = False

'transform PT in slide version
For Each PT In ActiveSheet.PivotTables
If PT.name = "ActualVsTarget1" Then

PT.PivotSelect "", xlDataAndLabel, True
Selection.Copy
Range("Y2").Select
ActiveSheet.Paste
Application.CutCopyMode = False

End If
Next PT

're-arrange PT
For Each PT In ActiveSheet.PivotTables
If PT.name <> "ActualVsTarget1" And PT.name <> "ActualVsTarget11" Then 'PT ActualvsTarget2 = PT showing FTEs

PT.PivotFields("Sum of Target Spend in US$").Orientation = xlHidden 'hide Target Spend Field

PT.PivotFields("Subregion").Orientation = xlColumnField
PT.PivotFields("Date").Orientation = xlPageField

PT.PivotFields("Supplier").Orientation = xlRowField
PT.PivotFields("Supplier").Position = 2
PT.PivotFields("Supplier").PivotItems("(blank)").Visible = False

PT.PivotFields("Service Line Description").Orientation = xlRowField
PT.PivotFields("Service Line Description").Position = 3

PT.PivotFields("Supplier").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)

PT.RowAxisLayout xlCompactRow 'changes layout

'filter PT to show data for this month only
PT.PivotFields("Date").ClearAllFilters
PT.PivotFields("Date").EnableMultiplePageItems = False
Today = InputBox("Enter the date for the reported month in the format dd/mm/yyyy")

PT.PivotFields("Date").CurrentPage = Today
Columns("K:K").AutoFit

PT.TableStyle2 = "PivotStyleDark9"  'changes pivot table colours
PT.PivotSelect "Farshore", xlDataAndLabel, True  'color farshore items differently
    With Selection.Interior
        .pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
'
''copy PT on slide
PT.PivotSelect "", xlDataAndLabel, True
Selection.Copy
PPpres.Slides(37).Shapes.PasteSpecial ppPasteEnhancedMetafile

End If
Next PT
End With

PPpres.Save
Application.CutCopyMode = False

End Sub

