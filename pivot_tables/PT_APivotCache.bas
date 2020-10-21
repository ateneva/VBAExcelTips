Sub RefreshPivotData()

Dim Wbk As Workbook
Dim Wks As Worksheet
Set Wbk = ActiveWorkbook

Dim PC As PivotCache
Dim PT As PivotTable

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva
'https://datageeking.wordpress.com/2017/07/03/how-can-i-check-if-ive-got-different-pivot-caches-in-my-workbook/
'https://datageeking.wordpress.com/2017/07/12/how-to-quickly-refilter-pivot-tables-that-have-different-pivot-caches-with-vba/
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

If Wbk.PivotCaches.Count = 0 Then
    MsgBox "The current workbook has 0 caches"

    Else

    'counts the number of PivotCaches
    MsgBox "The current workbook has " & Wbk.PivotCaches.Count & " caches"_
 _
     & vbNewLine & "The current workbook size is " & Round(FileLen(Wbk.FullName) / 1048576, 2) & " MB"

    '----------------------------------------------------------------------------------------
    'Option1: refresh  pivot caches
    'useful when you have multiple caches '+ listobjects fed through direct DB conections
    'and you only want to refresh the pivot tables but not the listobjects

        For Each PC In Wbk.PivotCaches
            PC.Refresh
            MsgBox PC.RecordCount & " records"  'shows the number of records for each cache

        Next PC
    '-----------------------------------------------------------------------------------------
    'Option 2: refresh individual pivot tables
        For Each Wks In Wbk.Worksheets     'refresh all pivot tables in a workbook
            Wks.Activate

            For Each PT In ActiveSheet.PivotTables
                PT.RefreshTable
                PT.SaveData = True
            Next PT
        Next Wks

End If
End Sub
