Attribute VB_Name = "PTEventsRefilter"
Option Explicit

Private Sub Workbook_AfterSave(ByVal Success As Boolean)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'code written by Angelina Teneva in March 2013
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Dim Wks As Worksheet
Dim PT As PivotTable
Dim PF As PivotField
Dim Ans As String
Dim Ans2 As Integer
'~~~~~~~~~~~~~~~~~~~~~~~
Worksheets("Export Costs Analysis").Activate

Ans2 = MsgBox("Would you like to refilter pivot tables", vbYesNo)
Select Case Ans2

'*********************************************************************************
'useful if the pivot tables come from different sources and slicer can't be used
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Case vbYes
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'apply filter for the latest month to all pivot tables as given by user --> might also be a keyword
Ans = InputBox("Please enter latest fiscal period in the format Period nn yyyy")

For Each Wks In ThisWorkbook.Worksheets
If Wks.PivotTables.Count > 0 And Wks.name <> "Presales Costs Trend by SL" And Wks.name <> "Costs Trend" And Wks.name <> "# Details" Then Wks.Activate

For Each PT In ActiveSheet.PivotTables
Set PF = PT.PivotFields("Fiscal year/period")

On Error Resume Next
PF.ClearAllFilters
PF.EnableMultiplePageItems = False
PF.CurrentPage = Ans
Next PT

Next Wks

Case vbNo
MsgBox ("Remember to re-filter before closing")
End Select

'*************************change currency for multiple pivot tables prior slicers*********************************************
Dim PT As PivotTable
Dim PF As PivotField
    
If Target.Address = [g3].Address Then 'currency is stored here, code is triggered via Worksheet Change event
Cur = Sheets("Project Worksheet").Range("g3").Value
        
Application.ScreenUpdating = False
      
For Each Wks In ThisWorkbook.Worksheets
If Wks.PivotTables.Count > 0 Then Wks.Activate

For Each PT In ActiveSheet.PivotTables
Set PF = PT.PivotFields("Currency")

On Error Resume Next
PF.ClearAllFilters
PF.EnableMultiplePageItems = False
PF.CurrentPage = Cur

Next PT
Next Wks
'******************************************************************************************************************************

End Sub
