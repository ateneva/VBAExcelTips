Attribute VB_Name = "OutAcctReview_Engage_Full"
Option Explicit

Sub SummarizeData()

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
Worksheets.Add.name = "PageLevel"
Worksheets.Add.name = "WidgetLevel"

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
            If PF.name = "DayDate" Or PF.name = "MonthDate" Or PF.name = "YearDate" Then PF.Orientation = xlColumnField
            If PF.name = "Platform" Or PF.name = "Publisher" Or PF.name = "SourceName" Or PF.name = "Widget" Then PF.Orientation = xlPageField
        Next PF
        
        On Error Resume Next 'to offset the loop from trying to create the same calculated fields twice
        'create the calculated fields for PageLevel
        PT.CalculatedFields.Add "Paid Page CTR", "=PaidClicks/PaidViews", True
        PT.CalculatedFields.Add "Organic Page CTR", "=OrganicCLicks/OrganicPVs", True
        PT.CalculatedFields.Add "Average CPC", "=GrossRevenue/PaidClicks", True
        PT.CalculatedFields.Add "Page RPM CC", "=GrossRevenueCC/PaidViews*1000", True
        PT.CalculatedFields.Add "Paid Listings per page", "=PaidListings/PaidPVs", True
        PT.CalculatedFields.Add "Organic Listings per page", "=OrganicListings/OrganicPVs", True
        PT.CalculatedFields.Add "Page Adblock", "=BlockedPVs/PaidPVs", True
        PT.CalculatedFields.Add "Paid Page Viewability", "ViewedPaidPVs/TotalPVs", True
        PT.CalculatedFields.Add "Paid Page Viewable CTR", "=PaidClicks/ViewedPaidPVs", True
        PT.CalculatedFields.Add "Paid Page Viewable RPM CC", "=GrossRevenueCC/ViewedPaidPVs*1000", True
        PT.CalculatedFields.Add "BlockRate", "=BlockedPVs/TotalPVs", True
        
        'create the calculated fields for WidgetLevel
        PT.CalculatedFields.Add "Average CPC CC", "=GrossRevenueCC/PaidClicks", True
        PT.CalculatedFields.Add "Organic Request CTR", "=OrganicCLicks/TotalRequests", True
        PT.CalculatedFields.Add "Paid Request CTR", "=PaidClicks/TotalRequests", True
        PT.CalculatedFields.Add "Request RPM CC", "=GrossRevenueCC/TotalRequests*1000", True
        PT.CalculatedFields.Add "Paid Listings per request", "=PaidListings/TotalRequests", True 'also referred to as PaidCoverage
        PT.CalculatedFields.Add "Organic Listings per request", "=OrganicListings/TotalRequests", True
        PT.CalculatedFields.Add "Widget Adblock", "=BlockedPaidRequests/TotalRequests", True
        PT.CalculatedFields.Add "Viewable Requests", "ViewedRequests/TotalRequests", True
        PT.CalculatedFields.Add "Viewable Requests Paid CTR", "=PaidClicks/ViewedRequests", True
        PT.CalculatedFields.Add "Viewable RPM CC", "=GrossRevenueCC/ViewedRequests*1000", True
     
 '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        'add the DataFields per sheet
    
    If Wks.name = "PageLevel" Then
        
        For Each PF In PT.PivotFields
            Field = PF.name
            Select Case Field
        
            Case "PaidPVs", "OrganicPVs", "PaidListings", "OrganicListings", "PaidClicks", "OrganicCLicks", "GrossRevenueCC", _
            "Paid Page CTR", "Organic Page CTR", "Average CPC", "Page RPM CC", "Paid Listings per page", "Organic Listings per Page", _
            "Page Adblock", "Paid Page Viewability", "Paid Page Viewable CTR", "Paid Page Viewable RPM CC":
                PF.Orientation = xlDataField
            End Select
        Next PF
       
    Else
                
        For Each PF In PT.PivotFields
            Field = PF.name
            Select Case Field
            
            Case "TotalRequests", "PaidListings", "OrganicListings", "PaidClicks", "OrganicCLicks", "GrossRevenueCC", _
            "Average CPC CC", "Organic Request CTR", "Paid Request CTR", "Request RPM CC", "Paid Listings per Request", _
            "Paid Listings per Request", "Organic Listings per Request", "Widget Adblock", "Viewable Requests", _
            "Viewable Requests Paid CTR", "Viewable RPM CC":
                PF.Orientation = xlDataField
            End Select
        Next PF
    End If
    
        'show the added DataFields in the Row Area
        PT.DataPivotField.Orientation = xlRowField
        
        'adjust the retrieved datafields to a presentble format
            For Each PF In PT.DataFields
                PF.Function = xlSum
                
                If PF.Position <= 6 Or PF.name Like "*Revenue*" Then PF.NumberFormat = "#,##" 'format retrieveables
                If PF.name Like "*CTR*" Or PF.name Like "*Adblock*" Or PF.name Like "*Paid Page Viewability*" Or PF.name Like "*Viewable Requests*" Then PF.NumberFormat = "0.0%"
                If PF.name Like "*CPC*" Or PF.name Like "*RPM CC*" Then PF.NumberFormat = "0.00"
                If PF.name Like "*Listings per*" Then PF.NumberFormat = "0"
                               
                Title = PF.name
                PF.name = Mid(Title, 8, Len(Title) - 7) & " " 'removes the sum of
            Next PF
           
    'adjust column width
    Cells.EntireColumn.AutoFit

    End With
    End If
    
    'protect and hide source data
    If Wks.name = "data" Then
        
        Wks.Protect ("inhead"), DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:=True
        Wks.Visible = xlSheetVeryHidden
    End If
    
Next Wks

End Sub

