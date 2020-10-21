
Sub ResetPTDefaultSettings()

Dim Wks As Worksheet
Dim PT As PivotTable

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva, 17/09/2016
'https://datageeking.wordpress.com/2017/08/10/reset-default-pivot-table-settings/'
'https://datageeking.wordpress.com/2017/06/27/5-default-pivot-table-settings-you-should-reset/'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each Wks In ActiveWorkbook.Worksheets

If Wks.name <> "2013" Then Wks.Activate
    For Each PT In ActiveSheet.PivotTables

        PT.TableStyle2 = "PivotStyleMedium18"                   'black and grey

        PT.RowAxisLayout xlCompactRow
        'PT.RowAxisLayout xlOutlineRow
        'PT.RowAxisLayout xlTabularRow

        PT.RefreshTable
        PT.DataPivotField.Orientation = xlRowField              'changes position of #Values

        '---------------------- --PivotTable Options --> Layout & Format tab
        PT.HasAutoFormat = False                                'turns off columns' autofit on update

        PT.DisplayErrorString = True                            'shows errors as empty cells
        PT.ErrorString = "-"                                    'displays a value for the error string
        PT.DisplayNullString = True                             'displays zeros as empty cells

        '-------------------------PivotTableOptions --> Totals & Filters tab
        PT.ColumnGrand = True                                   'shows grand totals for columns
        PT.RowGrand = True                                      'shows grand totals for rows

        PT.AllowMultipleFilters = True                          'allows multiple filters to be set on pivot fields

        '-------------------------Pivot Table Options --> Data tab
        PT.EnableDrilldown = False                              'prevents access to raw data on double click
        PT.PivotCache.MissingItemsLimit = xlMissingItemsNone    'removing non-existing items from filters

        PT.SaveData = False                                     'saves without source data; use in very special cases - e.g. manipulating your own templates
        PT.SaveData = True                                      'saves with source data


        '-----------------------------------------------------------------------------------------------------------
        PT.ShowDrillIndicators = True                           'shows +/- buttons

        PT.ShowPageMultipleItemLabel = True                     'all the page fields can be filtered for multiple items

        PF.EnableMultiplePageItems = True                       'allows more than 1 item to be selected in a page filter
        PF.EnableItemSelection = False                          'disables usage of the field dropdown in the user interface.

        PT.ShowPages PageField:="CampaignID"                    'whether it should add each item in the PageField as a separate PT on a separate sheet

        PT.ListFormulas                                         'shows the formulas of any calculated fields that have been added

        PT.ShowTableStyleRowHeaders = False                     'whether the first column should be bold

        PT.DisplayFieldCaptions = False                         'hide FieldGeaders --> useful if you want to hide that an object is PivotTable

        PT.DataLabelRange.Interior.Color = RGB(255, 0, 0)       'colours the PT header in red
        PT.RowRange.Interior.Color = RGB(255, 0, 0)             'colours the PT Row Labels in Range (applies to all row fields)
        PT.ColumnRange.Interior.Color = RGB(255, 0, 0)          'colours the PT Column Labels in Range (applies to all column fields)

    Next PT

Next Wks
End Sub
