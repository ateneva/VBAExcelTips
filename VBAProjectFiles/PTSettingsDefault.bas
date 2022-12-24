Attribute VB_Name = "PTSettingsDefault"
Option Explicit

Sub ResetPTDefaultSettings()

Dim Wks As Worksheet
Dim PT As PivotTable
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva, 17/09/2016
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each Wks In ActiveWorkbook.Worksheets

If Wks.name <> "2013" Then Wks.Activate
    For Each PT In ActiveSheet.PivotTables

        PT.RowAxisLayout xlCompactRow
        'PT.RowAxisLayout xlOutlineRow
        'PT.RowAxisLayout xlTabularRow
        
        'PivotTable Options --> Layout & Format tab
        PT.HasAutoFormat = False                                'turns off columns' autofit on update
        
        PT.DisplayErrorString = True                            'shows errors as empty cells
        PT.ErrorString = "-"                                    'displays a value for the error string
        PT.DisplayNullString = True                             'displays zeros as empty cells
        
        'PivotTableOptions --> Totals & Filters tab
        PT.ColumnGrand = True                                   'shows grand totals for columns
        PT.RowGrand = True                                      'shows grand totals for rows
        
        PT.AllowMultipleFilters = True                          'allows multiple filters to be set on pivot fields
      
        'Pivot Table Options --> Data tab
        PT.EnableDrilldown = False                              'prevents access to raw data on double click
        PT.PivotCache.MissingItemsLimit = xlMissingItemsNone    'removing non-existing items from filters
    Next PT

Next Wks
End Sub

