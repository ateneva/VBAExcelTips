Attribute VB_Name = "WbkConnection"
Option Explicit

Sub Add_Wbk_Connections()

Dim Wbk As Workbook
Dim Wks As Worksheet
Dim PT As PivotTable
Dim PC As PivotCache

Set Wbk = ActiveWorkbook
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'***********OLEDB connection*******************************************************************************************

'(1): create the connection (doesn't matter if you're going to be retrieving from Table or Query defined object in Access)
Wbk.Connections.Add2 "Tickets", _
"", Array("OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin; Data Source=C:\Users\Angelina\Box Sync\DashboardDBs\Tickets.accdb"), "EYAlldata", 3


'~~~~~~~~~changing existing Acccess DB (OLEDB) connection~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
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
    
    
'****************ODBC connection***************************************************************************************


'*~~~~~~~~~~updating an ODBC connection to MySQL database~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
 
Dim requete As String
requete = ThisWorkbook.Worksheets("main").Range("N5").Value
    With ActiveWorkbook.Connections("DWHEngage").ODBCConnection
        .BackgroundQuery = False
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
    
'****adding an Access DB connection by using Microsoft query~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workbooks("AddHyperLinktoaURLinPT.xlsb").Connections.Add2 "Query from MS Access Database", "" _
, Array(Array("ODBC;DSN=MS Access Database;DBQ=C:\USERS\ANGELINA\DOCUMENTS\EngagePerformance.accdb;DefaultDir=C:\USERS\ANGELINA\DOCUMENTS;DriverId=" _
  ), Array("25;FIL=MS Access;MaxBufferSize=2048;PageTimeout=5;")), _
"SELECT Global.PublisherContinent, Global.PublisherCountry, Global.PublisherID" & Chr(13) & "" & Chr(10) & "FROM `C:\USERS\ANGELINA\DOCUMENTS\EngagePerformance.accdb`.Global Global" _
        , 2

End Sub
