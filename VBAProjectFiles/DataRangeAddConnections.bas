Attribute VB_Name = "DataRangeAddConnections"
Option Explicit

Sub AddCSVConnection()

    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;C:\Users\Angelina\Dashboards\monthlyreports\japan_brands.csv", Destination:=Range("$A$1"))
        .name = "japan_brands"
        
        'My data has headers
        .FieldNames = True
        
        .RowNumbers = False
        
        'filldown formulas
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        
        'refresh file on open
        .RefreshOnFileOpen = False
        .TextFilePromptOnRefresh = False
        .RefreshStyle = xlOverwriteCells
        
        .SavePassword = False
        .SaveData = True
        
        'adjust column width
        .AdjustColumnWidth = False
        .RefreshPeriod = 0
        .TextFilePlatform = 65001
        .TextFileStartRow = 1
        
        'determine the type of connection
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        
        'determine the parsing type
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        
        .Refresh BackgroundQuery:=False
    End With

End Sub

Sub AddTxtConnection()
Attribute AddTxtConnection.VB_ProcData.VB_Invoke_Func = " \n14"
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~add connection to .txt file~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;C:\Users\Angelina\Dashboards\monthlyreports\result_2016-08-01.txt", Destination:=Range("$A$2"))
        .name = "result_2016-08-01"
        
        'my data has headers
        .FieldNames = True
        .RowNumbers = False
        
        'filldown formulas
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        
        'refresh file on open
        .RefreshOnFileOpen = False
        .TextFilePromptOnRefresh = False
        .RefreshStyle = xlInsertDeleteCells
                
        .SavePassword = False
        .SaveData = True
        
        'adjust column width
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePlatform = 65001
        .TextFileStartRow = 1
        
        'determine the parsing type
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        
        'determine the parsing type
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With

End Sub

Sub AddAccessConnection()

'(1): create the connection (doesn't matter if you're going to be retrieving from Table or Query defined object in Access)
Wbk.Connections.Add2 "Tickets", _
"", Array("OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;Data Source=C:\Users\Angelina\Box Sync\DashboardDBs\Tickets.accdb"), "EYAlldata", 3

'****adding an Access DB connection by using Microsoft query~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workbooks("AddHyperLinktoaURLinPT.xlsb").Connections.Add2 "Query from MS Access Database", "" _
, Array(Array("ODBC;DSN=MS Access Database;DBQ=C:\USERS\ANGELINA\DOCUMENTS\EngagePerformance.accdb;DefaultDir=C:\USERS\ANGELINA\DOCUMENTS;DriverId=" _
  ), Array("25;FIL=MS Access;MaxBufferSize=2048;PageTimeout=5;")), _
"SELECT Global.PublisherContinent, Global.PublisherCountry, Global.PublisherID" & Chr(13) & "" & Chr(10) & "FROM `C:\USERS\ANGELINA\DOCUMENTS\EngagePerformance.accdb`.Global Global" _
        , 2

End Sub

Sub AddODBCConnection()

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

End Sub
