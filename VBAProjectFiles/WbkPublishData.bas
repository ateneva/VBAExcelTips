Attribute VB_Name = "WbkPublishData"
Option Explicit

Sub PublishData()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'code written by Angelina Teneva, Dec 2012
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Dim OutputBk As Workbook
Set OutputBk = ThisWorkbook

Dim Ans As Integer

Dim sourcedata As String
sourcedata = ThisWorkbook.Worksheets("Control File Locations").Range("A28").Value

Dim access As String
access = ThisWorkbook.Worksheets("Control File Locations").Range("A31").Value

Dim lastrecord As String
lastrecord = ThisWorkbook.Worksheets("Control File Locations").Range("A34").Value

Dim Records As Long
Records = ThisWorkbook.Worksheets("ExportBW").Range("A2:A" & Worksheets("ExportBW").UsedRange.Rows.Count).Count

'*****************************************************************************
Ans = MsgBox("Did you check for commas?", vbYesNo)

Select Case Ans
Case vbYes

Workbooks.Open FileName:=sourcedata, ReadOnly:=False, UpdateLinks:=False
ActiveWorkbook.Worksheets("Export").Activate
ActiveSheet.Range(Cells(2, 1), Cells(ActiveSheet.UsedRange.Rows.Count, 90)).Clear

OutputBk.Worksheets("ExportBW").Activate
With ActiveSheet
ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=1

Range(Cells(2, 1), Cells(ActiveSheet.UsedRange.Rows.Count, 90)).SpecialCells(xlCellTypeVisible).Copy Workbooks("Export.xlsx").Worksheets("Export").Range("A2")
End With

Application.CutCopyMode = False
Workbooks("Export.xlsx").Worksheets("Export").Activate
ActiveWorkbook.Save

'updates the CSV file for pivot tables
ActiveWorkbook.saveas FileName:= _
        "C:\Users\TENEVAA\Documents\TS EMEA\I am Responsible For\Import-Export Balance\Current Month\Export\Export.csv" _
        , FileFormat:=xlCSV, CreateBackup:=False
        
ActiveWorkbook.Close

'******************************************************
'export data for importing into MS Access
Workbooks.Open FileName:=access, ReadOnly:=False, UpdateLinks:=False
ActiveWorkbook.Worksheets("Export_no header").Activate
ActiveSheet.Range(Cells(1, 1), Cells(ActiveSheet.UsedRange.Rows.Count, 90)).Clear  'clear data from previous run
'3 columns less than original spource data as customer columns are  excluded

OutputBk.Worksheets("ExportBW").Activate
With ActiveSheet
Range(Cells(2, 1), Cells(ActiveSheet.UsedRange.Rows.Count, 10)).SpecialCells(xlCellTypeVisible).Copy Workbooks("Export_no header.csv").Worksheets("Export_no header").Range("B1")
Range(Cells(2, 13), Cells(ActiveSheet.UsedRange.Rows.Count, 90)).SpecialCells(xlCellTypeVisible).Copy Workbooks("Export_no header.csv").Worksheets("Export_no header").Range("L1")
'data is copied in two parts because of the need to get rid of the 3 customer fields
Range("A1").Select
End With

OutputBk.Worksheets("Control File Locations").Activate
OutputBk.Save

'open history of latest database records and calculate the number of records currently in the database
Workbooks.Open FileName:=lastrecord, ReadOnly:=False, UpdateLinks:=False
ActiveWorkbook.Worksheets("Records").Activate
ActiveSheet.Range("F2").Formula = WorksheetFunction.Sum(Range(Cells(2, 2), Cells(ActiveSheet.UsedRange.Rows.Count, 2)))
ActiveWorkbook.Save

'update the starting record in the file for access import
Windows("Export_no header.csv").Activate

With ActiveWorkbook
Worksheets("Export_no header").Activate

With ActiveSheet
Range("A1").Formula = Workbooks("Access Records.xlsx").Worksheets("Records").Range("F2").Value + 1
Range("A2").Formula = "=R[-1]C+1"
Range("A2:A" & .UsedRange.Rows.Count).FillDown
End With

ActiveWorkbook.Save
ActiveWorkbook.Close

End With

Windows("Access Records.xlsx").Activate
With ActiveWorkbook
Worksheets("Records").Activate

'update the number of records for this run
Range("D1").Value = Records
Range("D1").NumberFormat = "0"
Range("D1").Copy ActiveSheet.Range("B1").End(xlDown).Offset(1, 0)

ActiveWorkbook.Save
End With

Case vbNo
MsgBox ("Please, check for commas")
Worksheets("ExportBW").Activate

End Select

End Sub
