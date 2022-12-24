Attribute VB_Name = "DataRangeFormatting"
Option Explicit

Sub NumberFormatting()

    Range("BP4:BS" & ActiveSheet.UsedRange.Rows.count).NumberFormat = "0"  'format Labour Hours and Total Cost as number with no decimal places
    Range("BM4:BM" & ActiveSheet.UsedRange.Rows.count).NumberFormat = "@"  'format Fl Text as Text
    
    Range("M:O").NumberFormat = "0.00"          '(ITO Labour Utilization as Double)
    Range("A:A").NumberFormat = "0.00"          '(FTE & TEnure as Double)
    
    PF.NumberFormat = "0%; -0%;"                'hides zero values but keeps the negative ones
    PF.NumberFormat = "#,###"                   'shows numbers as thosands separated by a deciaml sign, if the number is smaller than a thoudsand, it will show zero
    PF.NumberFormat = "0,##"                    'shows a leading zero in numbers smaller than a thousand
    
    PF.NumberFormat = "#,##0"                   'minus sign for negative values
    PF.NumberFormat = "#,##0_);(#,##0)"         'negative numbers in brackets
    
    PF.NumberFormat = "[$$-en-US]0.00"          'dollar currency symbol
    PF.NumberFormat = "[$€-x-euro2] #,##0.00"   'euro currency symbol
    PF.NumberFormat = "[$£-en-GB]#,##0.00"      'pound currency symbol

End Sub

Sub TextFormatting()
    
    Range("AB41").WrapText = False
    
    Range(ActiveCell, ActiveCell.Offset(0, 1)).Select
    Selection.MergeCells = True
    
    ActiveCell.VerticalAlignment = xlCenter
    Range("A2:AR" & .UsedRange.Rows.count).Font.Size = 10

End Sub

Sub Coloring()

    ActiveCell.Font.Color = RGB(197, 217, 241)    'light blue font colour
    
    ActiveCell.Font.Color = RGB(255, 0, 0)        'red
    ActiveCell.Font.Color = RGB(255, 255, 255)    'white
    Range("L28").Font.Color = RGB(0, 0, 0)        'black
    
    ActiveCell.Interior.Color = RGB(255, 255, 0)  'yellow
    
    ActiveCell.Interior.Color = 65535             'green
    ActiveCell.Interior.Color = 5287936           'yellow

End Sub

Sub InsertARowColumn()

    Rows("22:22").Insert Shift:=xlDown
    Rows("16:16").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
    Columns("E:E").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Columns("E:E").EntireColumn.AutoFit
    Columns("E:E").EntireColumn.Delete

End Sub

Sub Hidng()

Columns("C:D").EntireColumn.Hidden = True
Columns("F:K").EntireColumn.Hidden = True
Columns("M:AM").EntireColumn.Hidden = True

'checks for hidden columns on HC RMR data file and if any are found, it unhides them
If Columns("C:D").EntireColumn.Hidden = True Then Columns("C:D").EntireColumn.Hidden = False
If Columns("F:K").EntireColumn.Hidden = True Then Columns("F:K").EntireColumn.Hidden = False
If Columns("M:AM").EntireColumn.Hidden = True Then Columns("M:AM").EntireColumn.Hidden = False

End Sub

Sub TextToColumns()

'text to Columns by commas
Columns("B:B").TextToColumns Destination:=Range("B1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True

'remove duplicate entries
Range("R2:R" & .UsedRange.Rows.count).RemoveDuplicates Columns:=1, Header:=xlNo 'single column
Columns("A:B").RemoveDuplicates Columns:=Array(1, 2), Header:=xlYes             'multiple columns

End Sub

Sub HideUnhideWorksheets()

Dim Wks As Worksheet
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~
For Each Wks In ThisWorkbook.Worksheets
    If Wks.Visible = xlSheetVeryHidden Or Wks.Visible = xlSheetHidden Then Wks.Visible = True
Next Wks

End Sub






