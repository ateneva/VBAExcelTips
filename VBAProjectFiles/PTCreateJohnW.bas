Attribute VB_Name = "PTCreateJohnW"
Option Explicit

Sub MakePivotTables() 'This procedure creates 28 pivot tables

Dim PTCache As PivotCache
Dim PT As PivotTable
Dim SummarySheet As Worksheet
Dim ItemName As String
Dim row As Long, Col As Long, i As Long
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error Resume Next

'Create Pivot Cache
Set PTCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=Sheets("SurveyData").Range("A1").CurrentRegion)

row = 1
For i = 1 To 14
For Col = 1 To 6 Step 5 '2 columns
    ItemName = Sheets(“SurveyData”).Cells(1, i + 2)
    With Cells(row, Col)
    .Value = ItemName
    .Font.Size = 16
    End With
'Create pivot table
Set PT = ActiveSheet.PivotTables.Add(PivotCache:=PTCache, TableDestination:=SummarySheet.Cells(row + 1, Col))

'Add the fields
If Col = 1 Then 'Frequency tables
    With PT.PivotFields(ItemName)
    .Orientation = xlDataField
    .name = "Frequency"
    .Function = xlCount
    End With
    
Else ' Percent tables
    With PT.PivotFields(ItemName)
    .Orientation = xlDataField
    .name = “Percent”
    .Function = xlCount
    .Calculation = xlPercentOfColumn
    .NumberFormat = "0.0%"
    End With
End If

PT.PivotFields(ItemName).Orientation = xlRowField
PT.PivotFields(“Sex”).Orientation = xlColumnField
PT.TableStyle2 = "PivotStyleMedium2"
PT.DisplayFieldCaptions = False

If Col = 6 Then
'add data bars to the last column
PT.ColumnGrand = False
PT.DataBodyRange.Columns(3).FormatConditions.AddDatabar 'adds conditional formatting to pivottable

    With PT.DataBodyRange.Columns(3).FormatConditions(1)
        .BarFillType = xlDataBarFillSolid
        .MinPoint.Modify newtype:=xlConditionValueNumber, newvalue:=0
        .MaxPoint.Modify newtype:=xlConditionValueNumber, newvalue:=1
    End With
End If

Next Col
row = row + 10

End Sub


