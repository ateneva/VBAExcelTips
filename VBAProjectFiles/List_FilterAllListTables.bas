Attribute VB_Name = "List_FilterAllListTables"
Option Explicit

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

'ActiveSheet.ListObjects("Table3").Range.AutoFilter Field:=64, Criteria1:="3P" 'constant
'ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=3, Criteria1:=Array("D011", "E064", "E066"), Operator:=xlFilterValues 'more than 2 values

    For Each T In ActiveSheet.ListObjects
        If bU = "All" Then
            T.Range.AutoFilter Field:=1
            Else
            T.Range.AutoFilter Field:=1, Criteria1:=bU
        End If

    Next T

End If
End Sub



