Sub ShowFieldinPT()

Dim Wks As Worksheet
Dim PT As PivotTable
Dim PF As PivotField
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva, Sept 2016
'https://datageeking.wordpress.com/2017/08/02/how-do-i-quickly-change-the-pivot-field-name-of-only-specific-fields/
'https://datageeking.wordpress.com/2017/07/17/how-to-quickly-change-a-pivot-table-summary-function-with-vba/
'https://datageeking.wordpress.com/2017/07/25/quickly-change-a-pivot-field-number-formatting-with-vba/
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each Wks In ActiveWorkbook.Worksheets
    For Each PT In Wks.PivotTables

        For Each PF In PT.PivotFields
            Set PF = PT.PivotFields("Country")
            If PF.Orientation <> xlHidden Then

                'display a field in a pivot table
                    PF.Orientation = xlPageField     'as ReportFilter
                    PF.Orientation = xlRowField      'as RowField
                    PF.Orientation = xlColumnField   'as ColumnField
                    PF.Orientation = xlDataField     'as Value Field

                If PF.Orientation = xlDataField  Then
                   'adjust the summary function
                      PF.Function = xlSum
                      PF.Function = xlCount
                      PF.Function = xlCountNums
                      PF.Function = xlAverage
                      PF.Function = xlProduct
                      PF.Function = xlMax
                      PF.Function = xlMin

                  'adjust the calculation type
                      PF.Calculation = xlNormal
                      PF.Calculation = xlNoAdditionalCalculation
                      PF.Calculation = xlPercentOfRow
                      PF.Calculation = xlPercentOfColumn
                      PF.Calculation = xlPercentOfTotal
                      PF.Calculation = xlPercentRunningTotal
                      PF.Calculation = xlRunningTotal

                      PF.Calculation = xlPercentOfParentRow
                      PF.Calculation = xlPercentOfParentColumn
                      PF.Calculation = xlPercentOfParent
                      PF.Calculation = xlIndex

                      PF.Calculation = xlRankAscending
                      PF.Calculation = xlRankDescending

                      PF.Calculation = xlPercentOf
                      PF.BaseField = "Week"
                      PF.BaseItem = "(previous)"

                      PF.Calculation = xlPercentDifferenceFrom
                      PF.BaseField = "Week"
                      PF.BaseItem = "(previous)"

                      PF.Calculation = xlDifferenceFrom
                      PF.BaseField = "Week"
                      PF.BaseItem = "(previous)"

                      'adjust the number formatting
                      PF.NumberFormat = "0.00"                    'shows only two decimals
                      PF.NumberFormat = "#,###"                   'shows numbers as thousands
                      PF.NumberFormat = "0,##"                    'shows a leading zero in numbers smaller than a thousand

                      PF.NumberFormat = "#,##0"                   'shows minus sign for negative values
                      PF.NumberFormat = "#,##0_);(#,##0)"         'shows negative numbers in brackets

                      PF.NumberFormat = "[$$-en-US]0.00"          'dollar currency symbol
                      PF.NumberFormat = "[$€-x-euro2] #,##0.00"   'euro currency symbol
                      PF.NumberFormat = "[$£-en-GB]#,##0.00"      'pound currency symbol
                      PF.NumberFormat = "0%; -0%;"                'hides zero values but keeps the negative ones
                      PF.NumberFormat = "0.0%"                    'formats as %
                End If

            End If
        Next PF
    Next PT
Next Wks
End Sub
