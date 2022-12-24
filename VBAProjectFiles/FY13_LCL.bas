Attribute VB_Name = "FY13_LCL"
Option Explicit

Sub LCL_Actuals()
Attribute LCL_Actuals.VB_ProcData.VB_Invoke_Func = " \n14"

Dim PT As PivotTable
Dim PF As PivotField
'***************************

With ActiveWorkbook
Worksheets("Pivot").Activate

For Each PT In ActiveSheet.PivotTables

    PT.PivotFields("Month2").PivotItems("Q3").Visible = False 'excludes targets
    PT.PivotFields("Month2").PivotItems("Q4").Visible = False
    
    PT.PivotFields("NearShore").ShowDetail = True                  'shows details
    PT.RowAxisLayout xlTabularRow                                 'moves in a separate column
    PT.PivotFields("SL").Orientation = xlRowField
    PT.PivotFields("SL").Position = 1
    PT.PivotFields("SLC").Orientation = xlHidden
    PT.PivotFields("NearShore").Position = 2

    'remove sub-region totals
    PT.PivotFields("Subregion").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
            False, False)
        
    'add subototals for Location
    PT.PivotFields("NearShore").Subtotals = Array(True, False, False, False, False, False, False, False, False, False, _
            False, False)
            
    'add product line sub-totals
    PT.PivotFields("SL").Subtotals = Array(True, False, False, False, False, False, False, False, False, False, _
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

If PT.name <> "LCL Spend" Then
PT.PivotFields("Subregion").Orientation = xlRowField
PT.PivotFields("Subregion").Position = 1
PT.PivotFields("Subregion").ShowAllItems = True
PT.PivotFields("Subregion").Subtotals = _
        Array(True, False, False, False, False, False, False, False, False, False, False, False)

PT.PivotFields("NearShore").Position = 2

End If
Next PT

End With
MsgBox ("Include EMEA in Subregion Filter before exporting")
Application.Run "Prepare_LCL_slide"

End Sub

Sub Prepare_LCL_slide()

Dim PT As PivotTable
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
Worksheets("Split FY13").Activate
ActiveWindow.DisplayGridlines = False

For Each PT In ActiveSheet.PivotTables

    PT.PivotFields("Date").ClearAllFilters
    PT.PivotFields("Date").EnableMultiplePageItems = False
    Today = InputBox("Enter the date for the reported month in the format dd/mm/yyyy")
    
    PT.PivotFields("Date").CurrentPage = Today
    PT.PivotFields("Supplier").PivotItems("Target Spend").ShowDetail = False
    
    Columns("B:B").AutoFit
    PT.PivotSelect "", xlDataAndLabel, True
    Selection.Copy
    
    PPpres.Slides(37).Shapes.PasteSpecial ppPasteEnhancedMetafile

Next PT
End With

PPpres.Save
Application.CutCopyMode = False

End Sub
