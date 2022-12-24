Attribute VB_Name = "PTCreateAngelina"
Option Explicit

Sub SummarizeConversionsData()

Dim Wks As Worksheet
Dim PTCache As PivotCache
Dim PT As PivotTable
Dim PF As PivotField
Dim i As Integer
Dim Title As String
Dim Field As String
'~~~~~~~~~~~~~~~~~~~~~~~~

'Create the cache from a normal cell reference
Set PTCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=Worksheets("data").Range("A1").UsedRange)

'**************************************************************************************************************
'create 3 worksheets and Beautify them
Worksheets.Add.name = "summary"
'Worksheets.Add.name = "weekly_view"

For Each Wks In ActiveWorkbook.Worksheets

    If Wks.name <> "data" And Wks.name <> "interface" And Wks.name <> "summary1" Then
    
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
        PT.RowAxisLayout xlTabularRow 'change the default layout to Tabular
        
        PT.HasAutoFormat = False 'prevent autosort columns
        PT.EnableDrilldown = False 'prevent users from reaching the raw data
        PT.ColumnGrand = False 'turn off the column totals
        PT.RowGrand = False 'turn off the column totals
        PT.DisplayErrorString = True 'shows nothing on DIV errors
        PT.ShowDrillIndicators = False 'hides drills indicators
        PT.TableStyle2 = "PivotStyleLight6"
        PT.ShowTableStyleRowHeaders = False
                
        'add (Row & Column & Page Fields) --> the same for the two PivotTables
        For Each PF In PT.PivotFields
            
            'add the fields that are common for the two views (removes torals if necessary
            If PF.name = "CampaignID" Then PF.Orientation = xlPageField
           
            If PF.name = "Campaign" Then
                PF.Orientation = xlRowField
                PF.Subtotals(1) = False
            End If
            
            If PF.name = "Publisher" Then
                PF.Orientation = xlRowField
                PF.Subtotals(1) = False
            End If
            
            If PF.name = "SectionID" Then
                PF.Orientation = xlRowField
                PF.Subtotals(1) = False
            End If
            If PF.name = "ReferringSection" Then PF.Orientation = xlRowField
                                
        Next PF
        
        'adds all available data fields
        For i = 6 To PT.PivotFields.Count
            PT.PivotFields(i).Orientation = xlDataField
        Next i
               
        'change the orientation of the values
        PT.DataPivotField.Orientation = xlColumnField
        
        'adjust the retrieved datafields to a presentble format
        For Each PF In PT.DataFields
            PF.Function = xlSum
                 
            If PF.name Like "*ConversionRate*" Then
            PF.NumberFormat = "0.0%"
            Else
            PF.NumberFormat = "0"
            End If
                               
            Title = PF.name
            PF.name = Mid(Title, 8, Len(Title) - 7) & " " 'removes the sum of
        Next PF
        
'        'adjust sorting and layouts for row fields --> it needs to be in a separate loop; if defined when the field is being added, it's not working
        For Each PF In PT.RowFields
            PF.AutoSort xlDescending, "PaidClicks "
        Next PF
        
        PT.ShowPages PageField:="CampaignID" 'makes every campignID available as a pivottable on a separate tab
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    End With
    End If

    'protect and hide source data
    If Wks.name = "data" Then
        Wks.Protect ("inhead"), DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:=True
        Wks.Visible = xlSheetVeryHidden
    End If
    
    Application.DisplayAlerts = False 'prevents a pop-out message
    If Wks.name = "summary" Then Wks.Delete 'deletes the original Wks summary
    Application.DisplayAlerts = True

Next Wks

For Each Wks In ActiveWorkbook.Worksheets

    If Wks.Visible = True Then
        Wks.Activate
            With ActiveSheet
                Cells.EntireColumn.AutoFit       'adjust column width
                Columns("B:B").ColumnWidth = 34
                Columns("D:D").ColumnWidth = 34
                ActiveWindow.Zoom = 80
                ActiveWindow.DisplayGridlines = False
            End With
    End If
Next Wks

End Sub






