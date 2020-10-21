
Sub FilterAndCopyPivotTab()

Dim Wks As Worksheet
Dim PT As PivotTable
Dim PF As PivotField
Dim PI As PivotItem
Dim AM As String
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva, October 2016
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'loop through all Pivot filter items, filter and copy the pivot table tab

For Each PT In ActiveSheet.PivotTables
    Set PF = PT.PivotFields("AccountManager")

    For Each PF In PT.PivotFields

        If PF.Orientation = xlPageField Then
            PF.ClearAllFilters

            PF.EnableMultiplePageItems = False     'allows more than 1 item to be selected in a page filter
            PF.EnableItemSelection = False         'Disables the ability to use the field dropdown in the user interface.
            PF.IncludeNewItemsInFilter             'often used when a field has been manually filtered to exclude 1 or two predefined items to make sure that all newly added items will appaear in pivot items


                For Each PI In PF.PivotItems
                    AM = PI
                    PF.CurrentPage = AM

                    '-----------------------------------------------'
                    'Option 1: Copy the whole tab
                    Worksheets("AM").Copy Before:=Worksheets(1)


                    '-----------------------------------------------'
                    'Option 2: Copy just the pivot table and paste it on a newly added tab

                    PT.PivotSelect "", xlDataAndLabel, True
                    Selection.Copy
                    Worksheets.Add.name = AM
                    Worksheets(AM).Paste
                    '------------------------------------------------'

                Next PI
        End If
    Next PF
Next PT

End Sub
