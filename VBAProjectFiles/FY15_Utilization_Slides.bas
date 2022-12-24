Attribute VB_Name = "FY15_Utilization_Slides"
Option Explicit

Sub Utilization()
Attribute Utilization.VB_ProcData.VB_Invoke_Func = " \n14"
'~~~~~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Application.Calculation = xlCalculationAutomatic

Dim PT As PivotTable
Dim PF As PivotField
Dim PI As PivotItem
Dim l As String
Dim PL As String

Dim Sh As Shape
Dim PPApp As PowerPoint.Application
Dim PPpres As PowerPoint.Presentation
Dim PPSlide As PowerPoint.Slide

'Create a PP application and make it visible
Set PPApp = New PowerPoint.Application
PPApp.Visible = msoCTrue

'Open the presentation you wish to copy to
Set PPpres = PPApp.Presentations.Open("C:\Users\hp\Desktop\Utilization.pptm")

'************************************************************
'''needed if you're using Excel 2013 to prevent PowerPoint from losing focus and returning
'"shapes (unknown member) invalid request. the specified data type is unavailable"
'- Run-time error -2147188160 (80048240):View (unknown member) error
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
PPApp.Activate
PPApp.ActiveWindow.ViewType = ppViewNormal
PPApp.ActiveWindow.Panes(2).Activate 'standard ppt view
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'********************************************************************************************************************************
Worksheets("Utilization per PL Pillar").Activate
ActiveWindow.DisplayGridlines = False

With ActiveSheet
Range("A390:N500").Clear 'clears the 2nd Pivot table in the spreadhseet
Range("C2:L2").Cut Range("C3")

PT.HasAutoFormat = False

PT.PivotFields("DeliveryFlag").PivotItems("C").Visible = False 'excluides CW workers

On Error Resume Next
PT.PivotFields("PL").Orientation = xlHidden
PT.PivotFields("Subregion ").Orientation = xlPageField

PT.PivotFields("Sub Pillar").Orientation = xlHidden
PT.PivotFields("Pillar").Orientation = xlRowField
PT.PivotFields("Pillar").Position = 1

PT.PivotFields("PL cluster").Position = 2

PT.PivotFields("MarketOffering").Orientation = xlRowField
PT.PivotFields("MarketOffering").Position = 3

ActiveSheet.PivotTables("PivotTable1").PivotFields("PL cluster").Subtotals = _
        Array(True, False, False, False, False, False, False, False, False, False, False, False)

ActiveSheet.PivotTables("PivotTable1").PivotSelect "'PL cluster'[All;Total]", xlDataAndLabel, True
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.349986266670736
        .PatternTintAndShade = 0
    End With

Columns("C:C").ColumnWidth = 22.71
Columns("D:D").ColumnWidth = 16

'**********************************************************************************************************************
'get Delivery Pillar Utilization
'**********************************************************************************************************************
Worksheets("Utilization per PL Pillar").Activate
Set PT = Worksheets("Utilization per PL Pillar").PivotTables("PivotTable1")

PT.PivotFields("Pillar").ClearAllFilters
PT.PivotFields("Pillar").PivotFilters.Add Type:=xlCaptionEquals, Value1:="Delivery"
       
For Each PT In ActiveSheet.PivotTables
    Set PF = PT.PivotFields("Subregion ")
        PF.ClearAllFilters
        
        For Each PI In PF.PivotItems
            l = PI.Value
            PF.CurrentPage = l
            
            PT.PivotSelect "", xlDataAndLabel, True
            Selection.Copy
            
            Select Case l
                Case "CEE&I": PPpres.Slides(2).Shapes.PasteSpecial ppPasteEnhancedMetafile
                Case "FRA": PPpres.Slides(11).Shapes.PasteSpecial ppPasteEnhancedMetafile
                Case "GER": PPpres.Slides(20).Shapes.PasteSpecial ppPasteEnhancedMetafile
                Case "GWE": PPpres.Slides(29).Shapes.PasteSpecial ppPasteEnhancedMetafile
                Case "IBE": PPpres.Slides(38).Shapes.PasteSpecial ppPasteEnhancedMetafile
                Case "ITA": PPpres.Slides(47).Shapes.PasteSpecial ppPasteEnhancedMetafile
                Case "MEMA": PPpres.Slides(56).Shapes.PasteSpecial ppPasteEnhancedMetafile
                Case "UKI": PPpres.Slides(65).Shapes.PasteSpecial ppPasteEnhancedMetafile
                Case "RUS": PPpres.Slides(73).Shapes.PasteSpecial ppPasteEnhancedMetafile

            End Select
        Next PI
Next PT

''****************************************************************************************************************
''get Pursuit Pillar Utilization
''****************************************************************************************************************

PT.PivotFields("Pillar").ClearAllFilters
PT.PivotFields("Pillar").PivotFilters.Add Type:=xlCaptionEquals, Value1:="Pursuit"

For Each PT In ActiveSheet.PivotTables
    Set PF = PT.PivotFields("Subregion ")
        PF.ClearAllFilters
        
        For Each PI In PF.PivotItems
            l = PI.Value
            PF.CurrentPage = l
            
            PT.PivotSelect "", xlDataAndLabel, True
            Selection.Copy
            
            Select Case l
                Case "CEE&I": PPpres.Slides(3).Shapes.PasteSpecial ppPasteEnhancedMetafile
                Case "FRA": PPpres.Slides(12).Shapes.PasteSpecial ppPasteEnhancedMetafile
                Case "GER": PPpres.Slides(21).Shapes.PasteSpecial ppPasteEnhancedMetafile
                Case "GWE": PPpres.Slides(30).Shapes.PasteSpecial ppPasteEnhancedMetafile
                Case "IBE": PPpres.Slides(39).Shapes.PasteSpecial ppPasteEnhancedMetafile
                Case "ITA": PPpres.Slides(48).Shapes.PasteSpecial ppPasteEnhancedMetafile
                Case "MEMA": PPpres.Slides(57).Shapes.PasteSpecial ppPasteEnhancedMetafile
                Case "UKI": PPpres.Slides(66).Shapes.PasteSpecial ppPasteEnhancedMetafile
                Case "RUS": PPpres.Slides(74).Shapes.PasteSpecial ppPasteEnhancedMetafile

            End Select
        Next PI
Next PT

End With

PPpres.Save
PPpres.Close

End Sub
