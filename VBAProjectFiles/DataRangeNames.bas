Attribute VB_Name = "DataRangeNames"
Option Explicit

Sub AddNames()

    With ActiveSheet
    
    ActiveWorkbook.Names.Add name:="PlanRef", RefersToR1C1:="='Lookups'!R2C2:R26C2"
    ActiveWorkbook.Names.Add name:="BankHolidays", RefersToR1C1:="='Lookups'!R2C7:R7C7"
    
    'need to activate Sheet first
    ActiveWorkbook.Names.Add name:="Discretionary_IT_Plans", RefersToR1C1:="=qrytempReportDump!C5"
    ActiveWorkbook.Names.Add name:="RequestType", RefersToR1C1:="=qrytempReportDump!C18"
    
    ActiveWorkbook.Names.Add "latestdata", RefersToR1C1:=Worksheets("SalesData").Range("A1").CurrentRegion
        
    End With

End Sub

Sub DeleteNames()

Dim Wbk As Workbook
Dim n As name
'~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each name In ActiveWorkbook.Names
    name.Delete
Next name

'~~~~~~~~~~~~~~~~~~~~~~deleting avaialable names~~~~~~~~~~~~~~~~~~~~~~~~~~~~
With ActiveWorkbook                                 'deleting specific names
    .Names("BankHolidays").Delete
    .Names("Contract_Type").Delete
    .Names("CurrentWorkRequestStatus").Delete
    .Names("Days").Delete
    .Names("Discretionary_IT_Plans").Delete
    .Names("MoveRequest").Delete
    .Names("PlanRef").Delete
    .Names("RequestType").Delete
    .Names("text_closed").Delete
    .Names("text_launched").Delete

End With

End Sub


