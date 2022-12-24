Attribute VB_Name = "PTAddShowCalculatedFields"
Option Explicit

Sub AddAndShowPTCalculatedFields()
Attribute AddAndShowPTCalculatedFields.VB_ProcData.VB_Invoke_Func = " \n14"

Dim Wks As Worksheet
Dim PT As PivotTable
Dim PF As PivotField
Dim PI As PivotItem
Dim Title As String
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each Wks In ActiveWorkbook.Worksheets
If Wks.name <> "data" Then Wks.Activate

If ActiveSheet.PivotTables.Count > 0 Then

    For Each PT In ActiveSheet.PivotTables
    
        PT.EnableDrilldown = False
    
        'add calculated fields
        On Error Resume Next 'a calculated field is added to PivotCache, not individual pivot table
        PT.CalculatedFields.Add "CTR", "=Clicks/Impressions", True
        PT.CalculatedFields.Add "CPC", "=Spend/Clicks", True
        PT.CalculatedFields.Add "CPM", "=Spend/Impressions*1000", True
        PT.CalculatedFields.Add "CVR", "=Conversions/Clicks", True
        PT.CalculatedFields.Add "CPA", "=Spend/Conversions", True

        'make pivot tables visible in your pivot table
        For Each PF In PT.CalculatedFields
            PF.Orientation = xlDataField
        Next PF
        
    
        'adjust the retrieved datafields to a presentble format
        For Each PF In PT.DataFields
            PF.Function = xlSum
            Title = PF.name
            PF.name = Mid(Title, 8, Len(Title) - 7) & " "  'removes the sum of
            
            If PF.name Like "*CPM*" Or PF.name Like "*CPC*" Or PF.name Like "*CPA*" Then PF.NumberFormat = "[$$-en-US]0.00"
            If PF.name Like "*CTR*" Or PF.name Like "*CVR*" Then PF.NumberFormat = "0.0%"
            
            If PF.name Like "*Impressions*" Or PF.name Like "*Clicks*" Or PF.name Like "*Spend*" Then PF.NumberFormatFormat = "#,##"
        Next PF
       
    Next PT

End If
Next Wks
End Sub


