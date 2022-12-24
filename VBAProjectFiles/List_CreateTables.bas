Attribute VB_Name = "List_CreateTables"
Option Explicit

Sub FormatAsTable()

'assumes data is not fed via .csv connection --> creating the table will break off the connection

'Make a table
Worksheets("SalesData").ListObjects.Add xlSrcRange, Range("A1").CurrentRegion
Worksheets("SalesData").ListObjects.Add xlSrcRange, Range("A1").UsedRange

'format source data as table = 1st syntax
Range(Cells(2, 1), Cells(ActiveSheet.UsedRange.Rows.Count, 61)).Select
ActiveSheet.ListObjects.Add(xlSrcRange, Range(Cells(2, 1), Cells(ActiveSheet.UsedRange.Rows.Count, 61)), , xlYes).name = "Table1"
ActiveSheet.ListObjects("Table1").TableStyle = "TableStyleLight9"

'format source data as table = 2nd syntax
Worksheets("SalesData").ListObjects.Add(xlSrcRange, Range("A1", Worksheets("SalesData").UsedRange), , xlYes).name = "Table1"
Worksheets("SalesData").ListObjects("Table1").TableStyle = "TableStyleLight9"

'creates a PivotCache to be used by existing pivot tables
PT.ChangePivotCache ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, sourcedata:="Table1", VERSION:=xlPivotTableVersion12)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'declare a table
Dim Source As ListObject
Set Source = RMRData.ListObjects("Table1")

Source.Unlist 'removing the list object
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
End Sub




