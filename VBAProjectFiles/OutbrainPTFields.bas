Attribute VB_Name = "OutbrainPTFields"
Option Explicit

Sub AdjustPivotFieldsSummary()
Attribute AdjustPivotFieldsSummary.VB_ProcData.VB_Invoke_Func = " \n14"

Dim Wks As Worksheet
Dim PT As PivotTable
Dim PF As PivotField
Dim PI As PivotItem
Dim Title As String
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each Wks In ActiveWorkbook.Worksheets
If Wks.name <> "data" Then Wks.Activate

If ActiveSheet.PivotTables.count > 0 Then

    For Each PT In ActiveSheet.PivotTables
    
        PT.EnableDrilldown = False
    
    On Error Resume Next

'        PT.CalculatedFields.Add "CTR", "=Clicks/Impressions", True
'       PT.CalculatedFields.Add "CPC", "=Revenue/Clicks", True
'        PT.CalculatedFields.Add "CTR", "=Clicks/TotalPVs", True
'        PT.CalculatedFields.Add "PaidCTR", "=PaidClicks/PaidPVs", True
'        PT.CalculatedFields.Add "RPM", "=GrossRevenue/PaidPageViews*1000", True
'        PT.CalculatedFields.Add "RPM_CC", "=GrossRevenueCC/PaidPageViews*1000", True
'        PT.CalculatedFields.Add "Paid_Coverage", "=PaidListings/TotalRequests", True 'also referred to as PaidCoverage
'        PT.CalculatedFields.Add "BlockRate", "=BlockedPVs/TotalPVs", True
         
'         PT.CalculatedFields.Add "RPM", "=GrossRevenue/PVs*1000", True

        
        For Each PF In PT.CalculatedFields
            PF.Orientation = xlDataField
        Next PF
    
        For Each PF In PT.DataFields
            PF.Function = xlSum
            Title = PF.name
            PF.name = Mid(Title, 8, Len(Title) - 7) & " "  'removes the sum of
            PF.NumberFormatFormat = "#,##"
        Next PF

    Next PT

End If
Next Wks
End Sub

Sub RemoveALLCalculatedFields()

Dim Wks As Worksheet
Dim PT As PivotTable
Dim PF As PivotField
Dim DF As PivotField
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each Wks In ActiveWorkbook.Worksheets

    For Each PT In Wks.PivotTables
        For Each PF In PT.CalculatedFields
            'trying to change the orientation of the calculated field without going through the data field will result in an error
            For Each DF In PT.DataFields
                If DF.SourceName = PF.name Then DF.Parent.PivotItems(DF.name).Visible = False
            Next DF
        Next PF
    Next PT
    
Next Wks
End Sub
