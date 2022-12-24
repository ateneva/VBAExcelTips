Attribute VB_Name = "PTSettings"
Option Explicit

Sub PTSettings()

Dim PT As PivotTable

'refresh all pivot tables in active workbook
For Each Wks In ActiveWorkbook.Worksheets
Wks.Activate

    For Each PT In ActiveSheet.PivotTables
        PT.RefreshTable
        
        PT.DataPivotField.Orientation = xlRowField              'changes position of #Values
        
        PT.SaveData = False                                     'saves without source data; use in very special cases - e.g. manipulating your own templates
        PT.SaveData = True                                      'saves with source data
        
        PT.EnableDrilldown = True                               'allows access to raw data on double click
        PT.EnableDrilldown = False                              'prevents access to raw data on double click
        
        PT.ShowDrillIndicators = True                           'shows +/- buttons
        
        PT.AllowMultipleFilters = True                          'allows multiple filters'
        PT.AllowMultipleFilters = False
        
        PT.ShowPageMultipleItemLabel = True                     'all the page fields can be filtered for multiple items
        
        PF.EnableMultiplePageItems = True                       'allows more than 1 item to be selected in a page filter
        PF.EnableItemSelection = False                          'disables usage of the field dropdown in the user interface.
    
        PT.ShowPages PageField:="CampaignID"                    'whether it should add each item in the PageField as a separate PT on a separate sheet
        
        PT.ListFormulas                                         'shows the formulas of any calculated fields that have been added
    
        PT.TableStyle2 = "PivotStyleLight6"                     'determine the style of the PivotTable
        PT.ShowTableStyleRowHeaders = False                     'whether the first column should be bold
        
        PT.DisplayFieldCaptions = False                         'hide FieldGeaders --> useful if you want to hide that an object is PivotTable
        
        PT.DataLabelRange.Interior.Color = RGB(255, 0, 0)       'colours the PT header in red
        PT.RowRange.Interior.Color = RGB(255, 0, 0)             'colours the PT Row Labels in Range (applies to all row fields)
        PT.ColumnRange.Interior.Color = RGB(255, 0, 0)          'colours the PT Column Labels in Range (applies to all column fields)
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~PT default settings clean up~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        PT.RowAxisLayout xlCompactRow
        'PT.RowAxisLayout xlOutlineRow
        'PT.RowAxisLayout xlTabularRow
        
        PT.HasAutoFormat = False                                'turns off columns' autofit on update
        PT.PivotCache.MissingItemsLimit = xlMissingItemsNone    'removing old items
        
        PT.DisplayErrorString = True                            'shows errors as empty
        PT.ErrorString = "-"                                    'displays a value for the error string
        
        PT.DisplayNullString = True                             'displays zeros as empty cells
        
        PT.AllowMultipleFilters = True                          'allows multiple filters to be set on pivot fields
        
        PT.ColumnGrand = True                                   'shows grand totals for columns
        PT.RowGrand = True                                      'shows grand totals for rows
                  
    Next PT

Next Wks

ActiveWorkbook.ShowPivotTableFieldList = True                   'shows the PivotTableField List
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

End Sub
