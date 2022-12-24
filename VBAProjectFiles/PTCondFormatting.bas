Attribute VB_Name = "PTCondFormatting"
Option Explicit

Sub AddDataBarFieldsConditionalFormatting()

Dim Wks As Worksheet
Dim PT As PivotTable
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Set PT = ActiveSheet.PivotTables("PivotTable4")
'PT.DataBodyRange.Columns(3).FormatConditions.AddDatabar             'adds databar conditional formatting to pivottable

PT.PivotFields("Ctry Weight").DataRange.FormatConditions.AddDatabar
With PT.PivotFields("Ctry Weight").DataRange.FormatConditions(1)
    .BarFillType = xlDataBarFillGradient
    
'    .MinPoint.Modify newtype:=xlConditionValueNumber, newvalue:=0   'fixed min value
'    .MaxPoint.Modify newtype:=xlConditionValueNumber, newvalue:=1   'fixed min value
    
    .MinPoint.Modify newtype:=xlConditionValueAutomaticMin          'min value based on data
    .MaxPoint.Modify newtype:=xlConditionValueAutomaticMax          'max value based on data
    .Direction = xlContext
    
    .NegativeBarFormat.ColorType = xlDataBarColor
    .BarBorder.Type = xlDataBarBorderSolid
    .NegativeBarFormat.BorderColorType = xlDataBarColor
    .AxisPosition = xlDataBarAxisAutomatic
    
    .ScopeType = xlFieldsScope                                      'format horizontally (row-based), 3rd manual option
       
    .BarBorder.Color.Color = 2668287
    .BarBorder.Color.TintAndShade = 0

    .AxisColor.Color = 0
    .AxisColor.TintAndShade = 0
    
    .NegativeBarFormat.Color.Color = 255
    .NegativeBarFormat.Color.TintAndShade = 0
    
    .NegativeBarFormat.BorderColor.Color = 255
    .NegativeBarFormat.BorderColor.TintAndShade = 0
    
End With

End Sub

Sub Add_3ColourScale_Column_ConditionalFormatting()

Dim Wks As Worksheet
Dim PT As PivotTable
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Set PT = ActiveSheet.PivotTables("PivotTable3")

'add column-wide scale conditional formatting based on lowest and highest values
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
PT.PivotFields("% of prv month RevenueCC").DataRange.FormatConditions.AddColorScale ColorScaleType:=3
    
With PT.PivotFields("% of prv month RevenueCC").DataRange.FormatConditions(1)
    .ColorScaleCriteria(1).Type = xlConditionValueLowestValue
    .ColorScaleCriteria(1).FormatColor.Color = 7039480
    .ColorScaleCriteria(1).FormatColor.TintAndShade = 0

    .ColorScaleCriteria(2).Type = xlConditionValuePercentile
    .ColorScaleCriteria(2).Value = 50
    .ColorScaleCriteria(2).FormatColor.Color = 8711167
    .ColorScaleCriteria(2).FormatColor.TintAndShade = 0

    .ColorScaleCriteria(3).Type = xlConditionValueHighestValue
    .ColorScaleCriteria(3).FormatColor.Color = 8109667
    .ColorScaleCriteria(3).FormatColor.TintAndShade = 0
    .ScopeType = xlDataFieldScope

End With
End Sub

Sub Add_2ColourScale_Row_ConditionalFormatting()

Dim Wks As Worksheet
Dim PT As PivotTable
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Set PT = ActiveSheet.PivotTables("PivotTable4")

'add column-wide scale conditional formatting based on lowest and highest values
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
PT.PivotFields("% of Ctry Total GrossRevenue").DataRange.FormatConditions.AddColorScale ColorScaleType:=2

'conditional formatting uses BGR colour code rather than the standard RGB
'therefore the colour scale appears orange despite green and yellow colours being selected
With PT.PivotFields("% of Ctry Total GrossRevenue").DataRange.FormatConditions(1)
    .ColorScaleCriteria(1).Type = xlConditionValueLowestValue
    .ColorScaleCriteria(1).FormatColor.Color = vbYellow
    .ColorScaleCriteria(1).FormatColor.TintAndShade = 0

    .ColorScaleCriteria(2).Type = xlConditionValueHighestValue
    .ColorScaleCriteria(2).FormatColor.Color = vbGreen
    .ColorScaleCriteria(2).FormatColor.TintAndShade = 0

    .ScopeType = xlFieldsScope
End With
End Sub




