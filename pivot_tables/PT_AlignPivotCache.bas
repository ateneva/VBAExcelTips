
Sub AllignSourceData()

Dim Wks As Worksheet
Dim PT As PivotTable

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva
'https://datageeking.wordpress.com/2017/07/06/how-do-i-quickly-align-all-my-pivot-tables-to-the-same-pivot-cache/
'https://datageeking.wordpress.com/2017/06/30/why-should-i-align-my-pivot-caches/
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Application.DisplayAlerts = False

'The code below can also change the pivot table source from interanl (e.g. dataset in wbk)
                                                        'to external (e.g OLEDB, ODBC connection)

For Each Wks In ActiveWorkbook.Worksheets
    Wks.Activate

    For Each PT In ActiveSheet.PivotTables

        PT.CacheIndex = Sheets(1).PivotTables(1).CacheIndex
        '1 in Sheets(1) refers to the position of the sheet in the wbk
        '1 in PivotTables(1) refers to the first pivot table in the active worksheet

        PT.RefreshTable

    Next PT
Next Wks

Application.DisplayAlerts = True
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'The code above could generate error messages if:
        '1) a worksheet has multiple pivot tables in it
        '2) the workbook and/or worksheets are password-protected

End Sub
