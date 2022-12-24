Attribute VB_Name = "OutFRAcctTop200URL"
Option Explicit

Sub SummarizeClicksData()

Dim Wks As Worksheet
Dim PTCache As PivotCache
Dim PT As PivotTable
Dim PF As PivotField
Dim i As Integer
Dim Title As String
Dim Field As String
'~~~~~~~~~~~~~~~~~~~~~~~~

'Create the cache from a normal cell reference
Set PTCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, sourcedata:=Worksheets("data").UsedRange)

'**************************************************************************************************************
'create 3 worksheets and Beautify them
Worksheets.Add.name = "summary"
'Worksheets.Add.name = "WidgetLevel"

For Each Wks In ActiveWorkbook.Worksheets

    If Wks.name <> "data" And Wks.name <> "interface" Then
    
        Wks.Activate
        'beautify worksheet
            ActiveWindow.Zoom = 80
            ActiveWindow.DisplayGridlines = False
        
        'create_pivottable with datafields
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    With ActiveSheet
        'Create the pivot table from the created cache and apply pivottable style
        Set PT = ActiveSheet.PivotTables.Add(PivotCache:=PTCache, TableDestination:=Range("A5"))
        
        'set PT default settings
        PT.HasAutoFormat = False 'prevent autosort columns
        PT.EnableDrilldown = False 'prevent users from reaching the raw data
        PT.ColumnGrand = False 'turn off the column totals
        PT.RowGrand = False 'turn off the column totals
        PT.DisplayErrorString = True 'shows nothing on DIV errors
        PT.TableStyle2 = "PivotStyleLight19"
        
        'add PivotFields common for all PivotTables
        For Each PF In PT.PivotFields
            
            'add Publisher and Platform as a ReportFilter
            If PF.name = "Publisher" Then PF.Orientation = xlPageField
            If PF.name = "Platform" Then PF.Orientation = xlPageField
            If PF.name = "URL" Then PF.Orientation = xlRowField
           
        Next PF
        
        On Error Resume Next 'to offset the loop from trying to create the same calculated fields twice
        'create the calculated fields for PageLevel
        PT.CalculatedFields.Add "PageCTR", "=PaidClicks/PaidPageViews", True
        PT.CalculatedFields.Add "CPC", "=GrossRevenue/PaidClicks", True
        PT.CalculatedFields.Add "GrossRPM", "=GrossRevenue/PaidPageViews*1000", True

        
        'adds all available data fields
        For i = 6 To PT.PivotFields.Count
            PT.PivotFields(i).Orientation = xlDataField
        Next i
        
        'change the orientation of the values
        PT.DataPivotField.Orientation = xlColumnField
        
        'adjust the retrieved datafields to a presentble format
        For Each PF In PT.DataFields
            PF.Function = xlSum
                 
            If PF.name Like "*RPM*" Or PF.name Like "*CTR*" Or PF.name Like "*CPC*" Then
            PF.NumberFormat = "0.00"
            Else
            PF.NumberFormat = "#,##"
            End If
                               
            Title = PF.name
            PF.name = Mid(Title, 8, Len(Title) - 7) & " " 'removes the sum of
        Next PF
        
        'adjust sorting and layouts for row fields --> it needs to be in a separate loop; if defined when the field is being added, it's not working
        For Each PF In PT.RowFields
            If PF.name = "Date" Then PF.DataRange.Sort Order1:=xlAscending, Type:=xlSortLabels, Order2:=xlDescending, Type:=xlSortValues, Orientation:=xlTopToBottom
            If PF.name = "URL" Then PF.Position = 1
            If PF.name = "URL" Then PF.PivotFilters.Add2 Type:=xlTopCount, DataField:=PT.PivotFields("PaidClicks "), Value1:=200
            If PF.name = "URL" Then PF.AutoSort xlDescending, "PaidPageViews "
        Next PF
        
 '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'adjust column width
    Cells.EntireColumn.AutoFit
    Columns("A:A").ColumnWidth = 94.29

    End With
    End If

    'protect and hide source data
    If Wks.name = "data" Then

        Wks.Protect ("inhead"), DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:=True
        Wks.Visible = xlSheetVeryHidden
    End If
'
Next Wks

End Sub




