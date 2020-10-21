
Sub Pivots()

Dim Wks As Worksheet
Dim PT As PivotTable
Dim PF As PivotField

Dim UCost As PivotField
Dim Price As PivotField
Dim Qty As PivotField
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva, Oct 2016
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each Wks In ActiveWorkbook.Worksheets

Wks.Activate
    For Each PT In ActiveSheet.PivotTables
        PT.HasAutoFormat = False
        PT.PivotCache.MissingItemsLimit = xlMissingItemsNone

        Set UCost = PT.PivotFields("UnitPrice")
        Set Price = PT.PivotFields("ListPrice")
        Set Qty = PT.PivotFields("OrderQty")

        On Error Resume Next
        UCost.Orientation = xlDataField
        Price.Orientation = xlDataField
        Qty.Orientation = xlDataField

        'add and format value values
        For Each PF In PT.DataFields
            PF.Function = xlSum
            PF.NumberFormat = "#,###"

            If PF.Position = 1 Then PF.Caption = "Ucost"
            If PF.Position = 2 Then PF.Caption = "Price"
            If PF.Position = 3 Then PF.Caption = "Qty"
         Next PF

        'show calculated fields
        PT.CalculatedFields.Add "Cost", "=Qty*Ucost", True
        PT.CalculatedFields.Add "Revenue", "=Qty*Price", True
        PT.CalculatedFields.Add "Profit", "=Revenue - Cost", True

    Next PT

Next Wks
End Sub
