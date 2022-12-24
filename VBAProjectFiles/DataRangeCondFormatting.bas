Attribute VB_Name = "DataRangeCondFormatting"
Option Explicit

Sub Add_FormulaBased_ConditionalFormatting()

'delete previous rules
Range("H7:H10").Select
Selection.FormatConditions.Delete

'Delivery Core Utilization - red
    Range("H7").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=H7<=(M7*0.9)" 'values for EMEA region
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    
'Delivery Core Utilization - yellow
    Range("H7").Select
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=AND(H7<$M7,H7>($M7*0.9))" ' values for EMEA
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
    End With
    
 'Delivery Core Utilization - green
    Range("H7").Select
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=H7>=M7" ' values for EMEA
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
    End With

End Sub

