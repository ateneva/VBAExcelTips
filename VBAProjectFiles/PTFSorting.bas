Attribute VB_Name = "PTFSorting"
Option Explicit

Sub SortPivotTables()

Dim PT As PivotTable
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each PT In ActiveSheet.PivotTables

    PT.PivotFields("Company").AutoSort xlDescending, "Sum of Sales"         'sorts the pivottable by Sum of Sales DataField
    PT.PivotFields("Region").AutoSort xlAscending, "Region"                 'sorts the specified PivotField (xlRowField, xlColumnField)
    
    'sorts the PivoTable first by Labels (xlRowField, xlColumnField) and then by all the values in the DataValues area
    PT.PivotFields("URL").DataRange.Sort Order1:=xlDescending, Type:=xlSortLabels, Order2:=xlDescending, Type:=xlSortValues, Orientation:=xlTopToBottom
    
    PT.PivotFields("Region").PivotItems("North").Position = 11 'manual sorting

Next PT
End Sub

Sub SortAscOnWeek()

Dim Wks As Worksheet
Dim PT As PivotTable
Dim PF As PivotField
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each Wks In ActiveWorkbook.Worksheets
  
  For Each PT In Wks.PivotTables
    On Error Resume Next
    
        'sort on Week Row Labels
        For Each PF In PT.RowFields
        Set PF = PT.PivotFields("Week")
          If PF.Hidden = False Then PF.AutoSort xlAscending, PF.SourceName 'PF name as it appears in the datasource
        Next PF
    Next PT
Next Wks

End Sub

Sub AutoSortAllFields()

Dim Wks As Worksheet
Dim PT As PivotTable
Dim PF As PivotField
Dim CellInput As String
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
For Each Wks In ActiveWorkbook.Worksheets
  
  For Each PT In Wks.PivotTables
    On Error Resume Next

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~sorting on Labels~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        'sort on Row Labels
        For Each PF In PT.RowFields
          PF.AutoSort xlAscending, PF.SourceName 'PF name as it appears in the datasource
        Next PF

        'sorts ascending on Column Labels
        For Each PF In PT.ColumnFields
          PF.AutoSort xlAscending, PF.SourceName 'PF name as it appears in the datasource
        Next PF

'~~~~~~~~~~~~~~~sorting on DataFields (must loop through PivotFields collection)~~~~~~~~~~~~

        'sorts Descending on a Defined DataField
        For Each PF In PT.PivotFields
          PF.AutoSort xlDescending, "PaidClicks "
        Next PF
        
        'sorts Descending on a DataField defined by the user (via drop-down selection)
        For Each PF In PT.PivotFields
          PF.AutoSort xlDescending, CellInput
        Next PF

  Next PT
Next Wks
End Sub

Sub PTSortUsingCustomLists()

Dim Wks As Worksheet
Dim PT As PivotTable
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'define the custom list
Application.AddCustomList Array("Account One", "Account Two", "Account Three", "Account Four", "Account Five")

For Each Wks In ActiveWorkbook.Worksheets
  
  For Each PT In Wks.PivotTables
    On Error Resume Next
        PT.SortUsingCustomLists = False     'Autosort method is not using the custom list even though it is avaialble
        PT.PivotFields("Account").DataRange.Sort Order1:=xlAscending, Type:=xlSortLabels
        
        PT.SortUsingCustomLists = False     'Autosort method is using the available custom list despite the False setting as .OrderCustom has been given in the syntax
        PT.PivotFields("Account").DataRange.Sort Order1:=xlAscending, Type:=xlSortLabels, OrderCustom:=6
        
        PT.SortUsingCustomLists = True      'Autosort method is using the available custom list even though .OrderCustom has not been specified in the sort syntax
        PT.PivotFields("Account").DataRange.Sort Order1:=xlAscending, Type:=xlSortLabels
        
        PT.SortUsingCustomLists = True      'Autosort method is using the available custom list
        PT.PivotFields("Account").DataRange.Sort Order1:=xlAscending, Type:=xlSortLabels, OrderCustom:=6  'the 6th available CustomList
  Next PT

Next Wks
End Sub


