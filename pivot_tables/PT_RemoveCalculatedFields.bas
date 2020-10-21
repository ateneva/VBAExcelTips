
Sub DeleteCalculateFields()

Dim Wks As Worksheet
Dim PT As PivotTable
Dim PF As PivotField
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva, September 2016
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each Wks In ActiveWorkbook.Worksheets

    For Each PT In Wks.PivotTables
        For Each PF In PT.CalculatedFields
           PF.Delete
        Next PF
    Next PT

Next Wks
End Sub

'------------------------------------------------------------------------------
Sub RemovePTFieldsFromLayout()

Dim Wks As Worksheet
Dim PT As PivotTable
Dim PF As PivotField
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva, November 2016
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each Wks In ActiveWorkbook.Worksheets
    For Each PT In Wks.PivotTables
        For Each PF In PT.DataFields
            If PF.SourceName Like "*Paid*" Or PF.SourceName Like "*CPC*" Then PF.Orientation = xlHidden
        Next PF
    Next PT
Next Wks

End Sub
