Sub SummarizeCampaignData()

Dim Wks As Worksheet
Dim PTCache As PivotCache
Dim PT As PivotTable
Dim PF As PivotField
Dim i As Integer
Dim Title As String
Dim Field As String
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva, Oct 2016
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'Create the cache from a normal cell reference
Set PTCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, sourcedata:=Worksheets("data").UsedRange)

'**************************************************************************************************************
Worksheets.Add.name = "summary"

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
        PT.HasAutoFormat = False                  'prevent autosort columns
        PT.EnableDrilldown = False                'prevent users from reaching the raw data
        PT.ColumnGrand = False                    'turn off the column totals
        PT.RowGrand = False                       'turn off the column totals
        PT.DisplayErrorString = True              'shows nothing on DIV errors
        PT.TableStyle2 = "PivotStyleLight19"
        PT.RowAxisLayout xlTabularRow             'changes to tabular orientation
        PT.DisplayFieldCaptions = False           'removes filtering buttons
        PT.ShowDrillIndicators = False            'turns off drill indicators

        'add PivotFields common for all PivotTables
        For Each PF In PT.PivotFields

            If PF.name = "CampaignID" Then PF.Orientation = xlPageField

            If PF.name = "Campaign" Then
                PF.Orientation = xlRowField
                PF.Subtotals(1) = False
            End If

            If PF.name = "UserLocation" Then PF.Orientation = xlRowField
            If PF.name = "Date" Then PF.Orientation = xlRowField

        Next PF

        On Error Resume Next 'to offset the loop from trying to create the same calculated fields twice
        'create the calculated fields for PageLevel
        PT.CalculatedFields.Add "CTR", "=Clicks/Impressions", True
        PT.CalculatedFields.Add "CPC", "=Spend/Clicks", True
        PT.CalculatedFields.Add "CPM", "=Spend/Impressions*1000", True
        PT.CalculatedFields.Add "CVR", "=Conversions/Clicks", True
        PT.CalculatedFields.Add "CPA", "=Spend/Conversions", True

        'adds all available data fields
        For i = 9 To PT.PivotFields.Count
            PT.PivotFields(i).Orientation = xlDataField
        Next i

        'change the orientation of the values
        PT.DataPivotField.Orientation = xlColumnField

        'adjust the retrieved datafields to a presentble format
        For Each PF In PT.DataFields
            PF.Function = xlSum

            If PF.name Like "*CPM*" Or PF.name Like "*CPC*" Or PF.name Like "*CPA*" Then PF.NumberFormat = "[$$-en-US]0.00"
            If PF.name Like "*CTR*" Or PF.name Like "*CVR*" Then PF.NumberFormat = "0.0%"
            If PF.Position = 1 Or PF.Position <= 5 Then PF.NumberFormat = "#,##"

            Title = PF.name
            PF.name = Mid(Title, 8, Len(Title) - 7) & " " 'removes the sum of
        Next PF

      PT.PivotFields("Date").Position = 3
      PT.ShowPages PageField:="CampaignID" 'whether it should add each item in the PageField as a separate PT on a separate sheet
 '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'adjust column width
    Cells.EntireColumn.AutoFit

    End With
    End If

    'protect and hide source data
    If Wks.name = "data" Then

        Wks.Protect ("annie"), DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:=True
        Wks.Visible = xlSheetVeryHidden
    End If

Next Wks

'beautify the tab
For Each Wks In ActiveWorkbook.Worksheets

    If Wks.Visible = True Then
        Wks.Activate
            With ActiveSheet
                ActiveWindow.Zoom = 80
                ActiveWindow.DisplayGridlines = False
                Cells.EntireColumn.AutoFit       'adjust column width
                Rows("1:2").EntireRow.Hidden = True
            End With
    End If
Next Wks

End Sub
