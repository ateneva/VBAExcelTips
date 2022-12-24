Attribute VB_Name = "ForEachWksInWbk"
Option Explicit

Sub RefreshPT() 'Angelina's suggestion

Dim Wks As Worksheet
Dim PT As PivotTable

For Each Wks In ThisWorkbook.Worksheets

If Wks.name = "Backlog" Or Wks.name = "Current Actuals" Or _
Wks.name = "Previous Actuals" Or Wks.name = "WD15" Then

    For Each PT In Wks.PivotTables
        PT.RefreshTable
    Next PT

End If
Next Wks

End Sub

Sub FilterAllTablesInActiveSheet()

Dim Wks As Worksheet
Dim T As ListObject
Dim i As Integer

Dim bU As String
bU = ActiveSheet.Range("C6").Value
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
For Each Wks In ThisWorkbook.Worksheets
Wks.Activate

If ActiveSheet.ListObjects.Count > 0 Then

    For Each T In ActiveSheet.ListObjects
    If bU = "All" Then
        T.Range.AutoFilter Field:=1
    Else
    T.Range.AutoFilter Field:=1, Criteria1:=bU
    End If

Next T

End If
End Sub


