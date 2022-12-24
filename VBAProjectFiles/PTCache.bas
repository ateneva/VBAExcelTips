Attribute VB_Name = "PTCache"
Option Explicit

Sub CreateCaches()

Dim Wbk As Workbook
Dim Wks As Worksheet
Dim pt As PivotTable
Dim PC As PivotCache

Set Wbk = ActiveWorkbook
'*************************

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~only assuming your data contains strings less than 255 characters~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~if data has > 255 characters per string, the .Pivotcaches.Create returns a VBA Run Time Error 13

            'Create the cache from a normal cell reference (i.e source data is in the worksheet)
Set PTCache = Wbk.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=Range("A1").CurrentRegion)    'only available data in this ws
Set PTCache = Wbk.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=Worksheets("data").UsedRange) 'all of the used range on this ws

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            'Create the PivotTableCache from external reference (e.g. Access DB, named Tickets, Table EYAllData)

'(1): create the connection (doesn't matter if you're going to be retrieving from Table or Query defined object in Access)
Wbk.Connections.Add2 "Tickets", _
"", Array("OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;Data Source=C:\Users\Angelina\Box Sync\DashboardDBs\Tickets.accdb"), "EYAlldata", 3

'(2): create the PivotCache
Set PTCache = Wbk.PivotCaches.Create(SourceType:=xlExternal, SourceData:=ActiveWorkbook.Connections("Tickets"), Version:=6)

'***changing existing Acccess DB connection~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    With ActiveWorkbook.Connections("Tickets").OLEDBConnection
        .BackgroundQuery = True                         'set to False, to delay remaining code execution until the query has finished refreshing
        .CommandText = Array("1&1CampaignData")         'name of previous connection
        .CommandType = xlCmdTable
        .Connection = Array("OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;Data Source=C:\Users\Angelina\Box Sync\DashboardDBs\Tickets.accdb")
        .RefreshOnFileOpen = False
        .SavePassword = False
        .SourceConnectionFile = ""
        .ServerCredentialsMethod = xlCredentialsMethodIntegrated
        .AlwaysUseConnectionFile = False
    End With
    With ActiveWorkbook.Connections("Tickets")
        .name = "Tickets"                               'if you're chanigng the name, you must change it in all instances
        .Description = ""
    End With
    ActiveWorkbook.Connections("Tickets").Refresh
        
'****adding an Access DB connection by using Microsoft query~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workbooks("AddHyperLinktoaURLinPT.xlsb").Connections.Add2 "Query from MS Access Database", "" _
, Array(Array("ODBC;DSN=MS Access Database;DBQ=C:\USERS\ANGELINA\DOCUMENTS\EngagePerformance.accdb;DefaultDir=C:\USERS\ANGELINA\DOCUMENTS;DriverId=" _
  ), Array("25;FIL=MS Access;MaxBufferSize=2048;PageTimeout=5;")), _
"SELECT Global.PublisherContinent, Global.PublisherCountry, Global.PublisherID" & Chr(13) & "" & Chr(10) & "FROM `C:\USERS\ANGELINA\DOCUMENTS\EngagePerformance.accdb`.Global Global" _
        , 2
        
'***************************************refreshing an ODBC connection to MySQL database~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
 
Dim requete As String
requete = ThisWorkbook.Worksheets("main").Range("N5").Value
    With ActiveWorkbook.Connections("DWHEngage").ODBCConnection
        .BackgroundQuery = True
        .CommandText = requete
        .CommandType = xlCmdSql
        .Connection = "ODBC;DSN=DWH;OPTION=0;;PORT=3310;SERVER=mysqldwhslv.nydc1.outbrain.com;UID=ob_reader;"
        .RefreshOnFileOpen = False
        .SavePassword = False
        .SourceConnectionFile = ""
        .SourceDataFile = ""
        .ServerCredentialsMethod = xlCredentialsMethodIntegrated
        .AlwaysUseConnectionFile = False
    End With
    With ActiveWorkbook.Connections("DWHEngage")
        .name = "DWHEngage"
        .Description = ""
    End With
    ActiveWorkbook.Connections("DWHEngage").Refresh

'***********************************************creating PivotTables from the existing PivotCache***********************************************************************

'works for both external and internal connections
Set pt = ActiveSheet.PivotTables.Add(PivotCache:=PTCache, TableDestination:=Range("A3"))


'~~~change and create a new PivotCache from a pre-defined list object based on internal connection~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
pt.ChangePivotCache ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="Table1", Version:=xlPivotTableVersion14)


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                                                'Refresh All PivotCache(s)
                                                
Worksheets(1).PivotTables(1).PivotCache.RefreshOnFileOpen = True             'change the default setting to be refreshing the cache upon file opening

For Each PC In ActiveWorkbook.PivotCaches                                    'useful when you have multiple caches '+ listobjects fed through SQL queries
    PC.Refresh                                                               'and you only want to refresh the pivot tables but not the listobjects
Next PC
  
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

Sub Allign_Source_Data()

Dim Wks As Worksheet
Dim pt As PivotTable
'*************************
Application.DisplayAlerts = False

For Each Wks In ActiveWorkbook.Worksheets
    Wks.Activate

    On Error Resume Next
    For Each pt In ActiveSheet.PivotTables
    
        pt.CacheIndex = Sheets(1).PivotTables(1).CacheIndex
        pt.RefreshTable
    Next pt
Next Wks

Application.DisplayAlerts = True

End Sub
