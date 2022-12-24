Attribute VB_Name = "DataRangeCopy"
Option Explicit

Sub CopyingData()

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'copy/paste special all the data in a (single) column
Range("AQ19:AQ" & .UsedRange.Rows.Count).Copy
Range("AQ19").PasteSpecial xlPasteValuesAndNumberFormats 'same column
Range("AR19").PasteSpecial xlPasteValuesAndNumberFormats 'different column
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'activate sheet and copy/cut/paste visible cells only in another worksheet in ThisWorkbook
Cells.SpecialCells(xlCellTypeVisible).Copy ThisWorkbook.Worksheets("CATIS").Range("A1")                'visible cells in the whole sheet
Range("A3:BY" & ActiveSheet.UsedRange.Rows.Count).SpecialCells(xlCellTypeVisible).Copy FSC.Range("B1") 'used range visible cells only
Range("A3:A" & ActiveSheet.UsedRange.Rows.Count).Cut ThisWorkbook.News.Range("E4")                     'cut all used range cells in another wks in this workbook

'activate activate sheet and copy paste data in a specific cell in another existing workbook
TargetWbk = ThisWorkbook.Worksheets("Control File Locations").Range("A4").Value
'~~~~~~~~~~~~~~~~
Range("A18:AP" & ActiveSheet.UsedRange.Rows.Count).Copy TargetWbk.Worksheets("RMR").Range("A1")
Range(Cells(7, 1), Cells(ActiveSheet.UsedRange.Rows.Count, 71)).Copy TargetWbk.Worksheets("ExportBW").Range("A2")

'activate sheet and copy and paste data at the end of a another dataset in another existing workbook
TargetWbk = ThisWorkbook.Worksheets("Control File Locations").Range("A4").Value
'~~~~~~~~~~~~~~~~~
Range(Cells(15, 1), Cells(ActiveSheet.UsedRange.Rows.Count, 2)).Copy TargetWbk.Worksheets("Default").Range("A15").Rows.End(xlDown).Offset(1, 0)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'copy and paste variable range data in a existing in the same workbook, no sheet activation
    ThisWorkbook.Worksheets("data").Range("A2:L" & Worksheets("data").UsedRange.Rows.Count).Copy
    Worksheets("newdata").Range("A1").PasteSpecial xlValues

'copy and paste variable range data in a newly added sheet in the same workbook, no sheet activation
    ThisWorkbook.Worksheets("data").Range("A2:L" & Worksheets("data").UsedRange.Rows.Count).Copy
    ThisWorkbook.Worksheets.Add.name = "new data"
    Worksheets("new data").Range("A1").PasteSpecial xlValues
    
    ThisWorkbook.Worksheets.Add(after:=Worksheets("newdata")).name = "Open SDM" 'if you want to add the new worksheet at a particular locaion
    Worksheets("Open SDM").Range("A2").PasteSpecial xlValues
    Worksheets("Open SDM").Rows.End(xlDown).Offset(1, 0).PasteSpecial xlValues 'only works if the 1st cell is non-empty
    
'copy and paste all data in a different existing workbook, use workbook activation
    Workbooks.Open FileName:=CL, ReadOnly:=True, UpdateLinks:=False
    With ActiveWorkbook
        Worksheets("existing CLs").Cells.Copy ThisWorkbook.Worksheets("CLs").Range("A1")
    .Close
    End With

'by using xLastCell property
With ActiveWorkbook
    Worksheets("Report 1").Activate
    Range(Cells(5, 2), ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Copy OutputBk.Worksheets("data").Range("B6")
    'Range("B5:V" & .UsedRange.Rows.Count).Copy OutputBk.Worksheets("data").Range("B6") 'for some reason this line of code does not want to work with bluebook extract
End With

Range("E6:E7,E9,E10").Copy 'paste non-adjacent range
Range("X41").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
'paste EMEA data 'fixed adjacent range <---Delivery Scorecards Templates
SourceWbk = ThisWorkbook.Worksheets("Control File Locations").Range("A4").Value
OutputWbk = ThisWorkbook.Worksheets("Control File Locations").Range("A5").Value
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
SourceWbk.Worksheets("EMEA").Range("G7:AH36").Copy
OutputWbk.Worksheets("EMEA").Range("G7").PasteSpecial Paste:=xlPasteValuesAndNumberFormats

End Sub

Sub CopyData_ActivateWbkWks()

'the purpose of this sub is to get the latest data from rmr and vlookup the BW extact against it
Dim Ans As Integer

Dim OutputBk As Workbook
Set OutputBk = ThisWorkbook

Dim subregionpath As String
Dim rmrpath As String

'subregionpath = "C:\Users\TENEVAA\Documents\TS EMEA\I am Responsible For\Headcount Report\CL_Subregion_Country_mapping.xlsx"
'rmrpath = "C:\Users\TENEVAA\Documents\TS EMEA\I am Responsible For\Headcount Report\Headcount_ExportDataSet_FormatTemplate.xlsm"

subregionpath = ThisWorkbook.Worksheets("Control File Locations").Range("A4").Value
rmrpath = ThisWorkbook.Worksheets("Control File Locations").Range("A7").Value

'*********************************************************************
Ans = MsgBox("Did you clear filters", vbYesNo)

Select Case Ans
Case vbYes

'get Subregion Data---> copies the whole sheet
Workbooks.Open FileName:=subregionpath, ReadOnly:=True, UpdateLinks:=False
With ActiveWorkbook
    Worksheets("existing CLs").Cells.Copy
End With

OutputBk.Worksheets("CL").Activate
With ActiveSheet
    Range("A1").Select
    Selection.PasteSpecial xlPasteValues
    Range("A1").Value = Now
End With

'clears previous data
OutputBk.Worksheets("RMR").Activate
ActiveSheet.Cells.Clear
'*********************************************************************
'get RMR data
Workbooks.Open FileName:=rmrpath, ReadOnly:=True, UpdateLinks:=False
ActiveWorkbook.Worksheets("Template").Activate

With ActiveSheet
    'copy variable range and paste formattted dates as values
    Range("AQ19:AQ" & .UsedRange.Rows.Count).Copy
    Range("AC19").PasteSpecial xlPasteValuesAndNumberFormats 'copy data onto the same sheet
      
    'copy all data in another workbook
    Range("A18:AP" & .UsedRange.Rows.Count).Copy OutputBk.Worksheets("RMR").Range("A1")
End With

'close the opened workbook
Windows("CL_Subregion_Country_mapping.xlsx").Activate
ActiveWorkbook.Close

'close the opened workbook
Windows("Headcount_ExportDataSet_FormatTemplate.xlsm").Activate
ActiveWorkbook.Close

ThisWorkbook.Save

Case vbNo: MsgBox ("Please clear filters before proceeeding")
End Select

End Sub
