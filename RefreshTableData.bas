
Sub FilterAllTablesInActiveSheet()

Dim Wks As Worksheet
Dim T As ListObject
Dim i As Integer

Dim bU As String
bU = ActiveSheet.Range("C6").Value

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~`
'written by Angelina Teneva
'https://datageeking.wordpress.com/2017/09/16/quickly-refilter-all-your-tables-with-vba/
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'

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
