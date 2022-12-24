Attribute VB_Name = "PTCache"
Option Explicit

Sub CreateCaches()

Dim Wbk As Workbook
Dim Wks As Worksheet
Dim PT As PivotTable
Dim PC As PivotCache

Set Wbk = ActiveWorkbook
'*************************

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~only assuming your data contains strings less than 255 characters~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~if data has > 255 characters per string, the .Pivotcaches.Create returns a VBA Run Time Error 13

            'Create the cache from a normal cell reference (i.e source data is in the worksheet)
Set PTCache = Wbk.PivotCaches.Create(SourceType:=xlDatabase, sourcedata:=Range("A1").CurrentRegion)    'only available data in this ws
Set PTCache = Wbk.PivotCaches.Create(SourceType:=xlDatabase, sourcedata:=Worksheets("data").UsedRange) 'all of the used range on this ws

'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            'Create the PivotTableCache from external reference (e.g. Access DB, named Tickets, Table EYAllData)

'(1): create the connection (doesn't matter if you're going to be retrieving from Table or Query defined object in Access)
Wbk.Connections.Add2 "Tickets", _
"", Array("OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;Data Source=C:\Users\Angelina\Box Sync\DashboardDBs\Tickets.accdb"), "EYAlldata", 3

'(2): create the PivotCache
Set PTCache = Wbk.PivotCaches.Create(SourceType:=xlExternal, sourcedata:=ActiveWorkbook.Connections("Tickets"), VERSION:=6)


'***********************************************creating PivotTables from the existing PivotCache***********************************************************************

'works for both external and internal connections
Set PT = ActiveSheet.PivotTables.Add(PivotCache:=PTCache, TableDestination:=Range("A3"))

'~~~change and create a new PivotCache from a pre-defined list object based on internal connection~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
PT.ChangePivotCache ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, sourcedata:="Table1", VERSION:=xlPivotTableVersion14)

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                                                'Refresh All PivotCache(s)
                                                
Worksheets(1).PivotTables(1).PivotCache.RefreshOnFileOpen = True             'change the default setting to be refreshing the cache upon file opening

 
End Sub

Sub CountWbkCaches()

Dim PC As PivotCache
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
If ActiveWorkbook.PivotCaches.Count = 0 Then
    MsgBox "The current workbook has 0 caches"
    
    Else
    
    'counts the number of PivotCaches
    MsgBox "The current workbook has " & ActiveWorkbook.PivotCaches.Count & " caches"
    
        'shows the number of records for each cache
        For Each PC In ActiveWorkbook.PivotCaches
            MsgBox PC.RecordCount & " records"
        Next PC
    
    'Return the Date when PivotTable was last refreshed ND BY WHOM
    MsgBox Worksheets("Sheet1").PivotTables("PivotTable1").RefreshDate
    MsgBox Worksheets("Sheet1").PivotTables("PivotTable1").RefreshName

End If
End Sub

Sub CountWbkCachesAndShowWbkSize()

Dim PC As PivotCache
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
If ActiveWorkbook.PivotCaches.Count = 0 Then
    MsgBox "The current workbook has 0 caches"
    
    Else
    
    'counts the number of PivotCaches
    MsgBox "The current workbook has " & ActiveWorkbook.PivotCaches.Count & " caches" _
            & vbNewLine _
 & "The current workbook size is " & Round(FileLen(ActiveWorkbook.FullName) / 1048576, 2) & " MB"
    

End If

End Sub

Sub Allign_Source_Data()

Dim Wks As Worksheet
Dim PT As PivotTable
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
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~------
'The code above could generate error messages if:
        '1) a worksheet has multiple pivot tables in it
        '2) the workbook and/or worksheets are password-protected

End Sub

Sub RefreshPTs()

Dim Wbk As Workbook
Dim Wks As Worksheet
Dim PT As PivotTable
Dim PC As PivotCache
'*****************************
'written by Angelina Teneva
'*****************************

'refresh pivot caches
For Each PC In ActiveWorkbook.PivotCaches     'useful when you have multiple caches '+ listobjects fed through SQL queries
    PC.Refresh                                'and you only want to refresh the pivot tables but not the listobjects
Next PC

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each Wks In ActiveWorkbook.Worksheets     'refresh all pivot tables in a workbook
    Wks.Activate

    For Each PT In ActiveSheet.PivotTables
        PT.RefreshTable
        PT.SaveData = True
    Next PT
Next Wks

End Sub
