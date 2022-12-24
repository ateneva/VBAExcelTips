Attribute VB_Name = "PTAdjust"
Option Explicit

Sub FieldFormat()

Dim PT As PivotTable
Dim PF As PivotField
Dim PI As PivotItem
Dim Wks As Worksheet
Dim i As Integer
'
For Each Wks In ActiveWorkbook.Worksheets
Wks.Activate

'For i = 1 To 2
'Worksheets(i).Activate

For Each PT In ActiveSheet.PivotTables

'On Error Resume Next
'PT.PivotFields("Employee Class Desc").Orientation = xlHidden
'PT.PivotFields("Employee Reg / Temp Code").Orientation = xlPageField

'PT.PivotFields("Import Category2").PivotItems("Group1").ShowDetail = False
'
'   For Each PI In PT.PivotFields("Import Category2").PivotItems
'        If PI.name = "Group1" Then PI.name = "Other BU"
'        Next PI

    PT.PivotFields("UniqueRespondent").Orientation = xlPageField
    PT.PivotFields("UniqueRespondent").CurrentPage = "yes"
    
    PT.HasAutoFormat = False

Next PT
'Next i

Next Wks
End Sub

Sub Pivots()

Dim Wks As Worksheet
Dim PT As PivotTable
Dim PF As PivotField

Dim UCost As PivotField
Dim Price As PivotField
Dim Qty As PivotField
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

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
    
        For Each PF In PT.DataFields
            PF.Function = xlSum
            PF.NumberFormat = "#,###"
            
            If PF.Position = 1 Then PF.Caption = "Ucost"
            If PF.Position = 2 Then PF.Caption = "Price"
            If PF.Position = 3 Then PF.Caption = "Qty"
         Next PF
               
        PT.CalculatedFields.Add "Cost", "=Qty*Ucost", True
        PT.CalculatedFields.Add "Revenue", "=Qty*Price", True
        PT.CalculatedFields.Add "Profit", "=Revenue - Cost", True
                 
    Next PT

Next Wks
End Sub
