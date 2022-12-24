Attribute VB_Name = "OutAcctReviewEngOrg"
Option Explicit

Sub SummarizeOrganicData()

Dim Wks As Worksheet
Dim PTCache As PivotCache
Dim PT As PivotTable
Dim PF As PivotField
Dim i As Integer
Dim Title As String
Dim Field As String
'~~~~~~~~~~~~~~~~~~~~~~~~

'Create the cache from a normal cell reference
Set PTCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=Worksheets("data").Range("A1").CurrentRegion)

'**************************************************************************************************************
'create 3 worksheets and Beautify them
Worksheets.Add.name = "SectionLevel"
Worksheets.Add.name = "PlatformLevel"
Worksheets.Add.name = "Geo"
Worksheets.Add.name = "Timing"

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
        
        'add (Row & Column & Page Fields) --> the same for the two PivotTables
        For Each PF In PT.PivotFields
            
            'add Publisher as a ReportFilter
            If PF.name = "Publisher" Then PF.Orientation = xlPageField
            
            'add month and date on Timing tab
            If Wks.name = "Timing" And PF.name = "DayDate" Then PF.Orientation = xlRowField
            If Wks.name = "Timing" And PF.name = "MonthDate" Then PF.Orientation = xlRowField
                              
            'add country field to the pivot table
            If Wks.name = "Geo" And PF.name = "Continent" Then PF.Orientation = xlRowField
            If Wks.name = "Geo" And PF.name = "Country" Then PF.Orientation = xlRowField
                            
            'add the section level to the pivottable
            If Wks.name = "SectionLevel" And PF.name = "SourceName" Then PF.Orientation = xlRowField
            If Wks.name = "SectionLevel" And PF.name = "Country" Then PF.Orientation = xlRowField
            
            'add the platform breakdown
            If Wks.name = "PlatformLevel" And PF.name = "Platform" Then PF.Orientation = xlRowField
           
        Next PF
        
        On Error Resume Next 'to offset the loop from trying to create the same calculated fields twice
        'create the calculated fields for PageLevel
        PT.CalculatedFields.Add "Organic Page CTR", "=OrganicCLicks/OrganicPVs", True
        PT.CalculatedFields.Add "Organic Listings per page", "=OrganicListings/TotalPVs", True
        
        'adds all available data fields
        For i = 10 To PT.PivotFields.Count
            PT.PivotFields(i).Orientation = xlDataField
        Next i
        
        'change the orientation of the values
        PT.DataPivotField.Orientation = xlColumnField
        
        'adjust the retrieved datafields to a presentble format
        For Each PF In PT.DataFields
            PF.Function = xlSum
                 
            If PF.name Like "*CTR*" Or PF.name Like "*per page*" Then
            PF.NumberFormat = "0.0%"
            Else
            PF.NumberFormat = "#,##"
            End If
                               
            Title = PF.name
            PF.name = Mid(Title, 8, Len(Title) - 7) & " " 'removes the sum of
        Next PF
        
        'adjust sorting and layouts for row fields --> it needs to be in a separate loop; if defined when the field is being added, it's not working
        For Each PF In PT.RowFields
            If PF.name = "MonthDate" Then PF.Position = 1
            If PF.name = "Country" Then PF.Position = 2
            If PF.name = "Country" Then PF.PivotFilters.Add Type:=xlValueIsGreaterThan, DataField:=PT.PivotFields("OrganicCLicks "), Value1:=20
            PF.AutoSort xlDescending, "OrganicCLicks "
        Next PF
        
           
 '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'adjust column width
    Cells.EntireColumn.AutoFit

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



