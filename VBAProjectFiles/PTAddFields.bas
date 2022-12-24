Attribute VB_Name = "PTAddFields"
Option Explicit

Sub UpdatePivotsinWBK()

Dim Wks As Worksheet
Dim PT As PivotTable
Dim PF As PivotField
Dim PI As PivotItem
Dim Title As String
Dim i As Integer
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
For Each Wks In ActiveWorkbook.Worksheets 'will loop through all worksheets in the workbook
Wks.Activate

With ActiveSheet
For Each PT In ActiveSheet.PivotTables 'loops through all pivottables in a worksheet

'~~~~~~~~~~~~~~~~~~~~~~~~pivot fields visibility~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
PT.PivotFields("Description").Orientation = xlRowField
PT.PivotFields("Description").Orientation = xlColumnField
PT.PivotFields("Description").Orientation = xlPageField
PT.PivotFields("Description").Orientation = xlDataField
PT.PivotFields("Description").Orientation = xlHidden
PT.DataPivotField.Orientation = xlRowField  'changes position of #Values

'Respectively you've got the following collections
For Each PF In PT.PageFields
For Each PF In PT.RowFields
For Each PF In PT.ColumnFields
For Each PF In PT.DataFields
For Each PF In PT.CalculatedFields
For Each PF In PT.HiddenFields
For Each PF In PT.VisibleFields

PT.PivotFields("Description").Position = 1
PT.PivotFields("Description").Caption = "Pillar"
PF.Caption = "Costs $" 'assumes field has been declared before that

'~~~~~~~~~~~~~~~~~~~~~~~adding data fields(PF is not declared beforehand~~~~~~~~~~~~~~~~~~~~~~~~~~
PT.AddDataField PT.PivotFields("Av. Spend per stay"), "Total Spend", xlSum
PT.AddDataField PT.PivotFields("Visitor Name"), "# visits", xlCount
PT.AddDataField PT.PivotFields("Av. Spend per stay"), "spend per visit", xlAverage

'OR
With PT.PivotFields("Variance")
.Orientation = xlDataField
.Function = xlSum
.Position = 3
.NumberFormat = "0.00%"
.Caption = "Variance-%"
End With

'OR adjusts the names to a presentable format
For Each PF In PT.DataFields
    PF.Function = xlSum
    PF.NumberFormat = "#,##"
    Title = PF.name
    PF.name = Mid(Title, 8, Len(Title) - 7) & " " 'removes the "sum of"
Next PF

'adds all available data fields
For i = 9 To PT.PivotFields.count
    PT.PivotFields(i).Orientation = xlDataField
Next i

'~~~~~~~~~~~~~~~~~~~~~~~adding calculated fields~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'the calculated field only needs to be created for one of the pivot tables sharing the same source

On Error Resume Next
'if spaces in the naming of the fields are avaialble
PT.CalculatedFields.Add "spend per day", "='Av. Spend per stay'/'# of calendar days'", True

PT.CalculatedFields.Add "CTR", "=PaidClicks/PaidPageViews", True 'if no spaces in the naming of the fields
PT.CalculatedFields.Add "Variance", " = Budget - Actual"

PT.AddDataField PT.PivotFields("spend per day"), "spend per day ", xlSum 'to make it visible in PT
PT.PivotFields("spend per day").NumberFormat = "£0"

'OR
    'add the CTR, and RPM
    PT.CalculatedFields.Add "Paid CTR", "=PaidClicks/PaidPageViews", True
    PT.CalculatedFields.Add "Organic CTR", "=OrganicClicks/OrganicPageViews", True
    PT.CalculatedFields.Add "RPM", "=GrossRevenue/OrganicPageViews", True
    
        For Each PF In PT.CalculatedFields 'if you want to manipulate/showmultiple calculated fields through a loop
            PF.Orientation = xlDataField
            PF.Function = xlSum
            PF.NumberFormat = "0.000"
        Next PF
'*****************************************************************************************************************
'adding calculated items
PT.PivotFields("Region").CalculatedItems.Add name:="E/N Am", Formula:="=Europe/North America" 'creates the field
PT.PivotFields("Region").PivotItems("E/N Am").Position = 3                                    'specifies the position in the pivot table
PT.PivotFields("Region").PivotItems("E/N Am").Caption = "Europe/N America"                    'chnages the visible name

'~~~~~~~~~~~~~~~~~~~and summary functions~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
PT.PivotFields("Activity Hours").Function = xlSum 'provided field is visible
'PT.PivotFields("Activity Hours").Function = xlCount
'PT.PivotFields("Activity Hours").Function = xlCountNums
'PT.PivotFields("Activity Hours").Function = xlAverage
'PT.PivotFields("Activity Hours").Function = xlProduct
'PT.PivotFields("Activity Hours").Function = xlMax
'PT.PivotFields("Activity Hours").Function = xlMin

'~~~~~~~~~~~~~~~~~~~~~~changing calculation types ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
PT.PivotFields("Sum of Activity Hours").Calculation = xlNormal
PT.PivotFields("Sum of Activity Hours").Calculation = xlNoAdditionalCalculation
'PT.PivotFields("Sum of Activity Hours").Calculation = xlPercentOfRow
'PT.PivotFields("Sum of Activity Hours").Calculation = xlPercentOfColumn
'PT.PivotFields("Sum of Activity Hours").Calculation = xlPercentOfTotal
'PT.PivotFields("Sum of Activity Hours").Calculation = xlPercentRunningTotal
'PT.PivotFields("Sum of Activity Hours").Calculation = xlRunningTotal

'PT.PivotFields("Sum of Activity Hours").Calculation = xlPercentOfParentRow
'PT.PivotFields("Sum of Activity Hours").Calculation = xlPercentOfParentColumn
'PT.PivotFields("Sum of Activity Hours").Calculation = xlPercentOfParent
'PT.PivotFields("Sum of Activity Hours").Calculation = xlIndex

'PT.PivotFields("Sum of Activity Hours").Calculation = xlRankAscending
'PT.PivotFields("Sum of Activity Hours").Calculation = xlRankDescending

'PT.PivotFields("Sum of Activity Hours").Calculation = xlPercentOf
'PT.PivotFields("Sum of Activity Hours").BaseField = "Week"
'PT.PivotFields("Sum of Activity Hours").BaseItem = "(previous)"

'PT.PivotFields("Sum of Activity Hours").Calculation = xlPercentDifferenceFrom
'PT.PivotFields("Sum of Activity Hours").BaseField = "Week"
'PT.PivotFields("Sum of Activity Hours").BaseItem = "(previous)"

'PT.PivotFields("Sum of Activity Hours").Calculation = xlDifferenceFrom
'PT.PivotFields("Sum of Activity Hours").BaseField = "Week"
'PT.PivotFields("Sum of Activity Hours").BaseItem = "(previous)"

'~~~~~~~~~~~~~~~~~~~changing DataFields formats~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
PT.DataBodyRange.NumberFormat = "0,0;;"                   'formats a single field
PT.DataBodyRange.NumberFormat.NumberFormat = "0.00%"
PT.DataBodyRange.NumberFormat = "#,#"                     'formats all fields in values section
PT.DataBodyRange.NumberFormat = "#,###"                   'formats all the fields currenty in the values area

'formats only numbers associated with a DataItem
PT.PivotFields("Region").PivotItems("Europe").DataRange.NumberFormat = "$#,##0"
PT.PivotFields("Region").PivotItems("North America").DataRange.NumberFormat = "$#,##0"
PT.PivotFields("Region").PivotItems("Europe/N America").DataRange.NumberFormat = "0.00%"

'.DataBodyRange -------------------------------------------> Object in the PivotTable
'.DataRange -----------------------------------------------> Object in the PivotField and PivotItems

PT.DataBodyRange.Columns(3).FormatConditions.AddDatabar   'adds databar conditional formatting to pivottable
With PT.DataBodyRange.Columns(3).FormatConditions(1)
    .BarFillType = xlDataBarFillSolid
    .MinPoint.Modify newtype:=xlConditionValueNumber, newvalue:=0
    .MaxPoint.Modify newtype:=xlConditionValueNumber, newvalue:=1
End With

Next PT
End With

Next Wks
End Sub
