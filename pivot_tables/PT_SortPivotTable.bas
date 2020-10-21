
Sub AutoSortAllFields()

Dim Wks As Worksheet
Dim PT As PivotTable
Dim PF As PivotField
Dim CellInput As String
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva, Oct 2016
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

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
