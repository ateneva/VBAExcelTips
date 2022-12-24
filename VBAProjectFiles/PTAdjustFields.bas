Attribute VB_Name = "PTAdjustFields"
Option Explicit

Sub AdjustPTFieldsFunction()

Dim PT As PivotTable
Dim PF As PivotField
Dim Title As String

For Each PT In ActiveSheet.PivotTables

'.DataBodyRange --> Object in the PivotTable
'.DataRange --> Object in the PivotField and PivotItems

PT.DataBodyRange.NumberFormat = "#,#" 'formats all fields in values section
PT.DataBodyRange.NumberFormat = "#,###" 'formats all the fields currenty in the values area

'adjust the datafields
    For Each PF In PT.DataFields
        PF.Function = xlSum
        PF.NumberFormat = "#,##"
        Title = PF.name
        PF.name = Mid(Title, 8, Len(Title) - 7) & " " 'removes the sum of
    Next PF
    
PT.CalculatedFields.Add "CTR", "=PaidClicks/PaidListings*100", True
PT.CalculatedFields.Add "CPC", "=GrossRevenue/PaidClicks", True
   
    'adjust the calculated fields
    For Each PF In PT.CalculatedFields
        PF.Orientation = xlDataField
        PF.Function = xlSum
        PF.NumberFormat = "#,##0.000"
        Title = PF.name
        PF.name = Mid(Title, 8, Len(Title) - 7) & " " 'removes the sum of
    Next PF
Next PT
End Sub



