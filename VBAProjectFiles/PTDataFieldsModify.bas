Attribute VB_Name = "PTDataFieldsModify"
Option Explicit

Sub ModifyDataFields()

Dim PT As PivotTable
Dim PF As PivotField
Dim PI As PivotItem
Dim i As Integer

For i = 1 To 2
Worksheets(i).Activate
Set PT = ActiveSheet.PivotTables(1)

Select Case i

Case 1
ActiveSheet.name = Format(ActiveSheet.Range("K2"), "dd-mmm")
For Each PF In PT.DataFields
'must use DataFields Collection if you are going to change the method of consolidation
    If PF.Position > 4 Then PF.Function = xlCountNums
    If PF.Position <= 4 Then PF.Function = xlSum
    If PF.Position <= 4 Then PF.NumberFormat = "0.0"
Next PF

Case 2
ActiveSheet.name = "weeks" & Format(ActiveSheet.Range("K2"), "dd-mmm")

For Each PF In PT.DataFields
'must use DataFields Collection if you are going to change the method of consolidation
    If PF.Position > 3 Then PF.Function = xlCountNums
    If PF.Position <= 3 Then PF.Function = xlSum
    If PF.Position <= 3 Then PF.NumberFormat = "0.0"
    Next PF

End Select

'PT.PivotSelect "TS IC HC Country[Romania,Croatia]", xlDataAndLabel, True
'Selection.Group
Next i
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

