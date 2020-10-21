
Sub MakeCalcFieldVisible()
Dim Wks As Worksheet
Dim PT As PivotCache
Dim PT As PivotTable
Dim PF As PivotField
Dim Title As String
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva, Oct 2016
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

  For Each Wks In ActiveWorkbook.Worksheets
    For Each PT In Wks.PivotTables

      On Error Resume Next  'in case all PTs share the same cache
      PT.CalculatedFields.Add "CTR", "=Clicks/Impressions", True
      PT.CalculatedFields.Add "CPC", "=AdSpend/Clicks", True

        For Each PF In PT.CalculatedFields
            If PF.SourceName Like "*CTR*" Then
                PF.Orientation = xlDataField
                PF.Function = xlSum
                Title = PF.SourceName & " "
                PF.Caption = Title
                PF.NumberFormat = "0.0%"
            End If
        Next PF
    Next PT

Next Wks
End Sub
