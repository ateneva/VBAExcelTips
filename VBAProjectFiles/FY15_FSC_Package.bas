Attribute VB_Name = "FY15_FSC_Package"
Option Explicit

Sub ExportFSCSlides()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva, 2015
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Dim PPApp As PowerPoint.Application
Dim PPpres As PowerPoint.Presentation

Dim pptx As String
pptx = ActiveWorkbook.Worksheets("calculated fields").Range("F2")

Dim Cell As Range
Dim Country As Range

Dim Cht As Chart
Dim ChtObj As ChartObject
Dim i As Integer

'Create a PP application and make it visible
Set PPApp = New PowerPoint.Application
PPApp.Visible = msoCTrue

'Open the presentation you wish to copy to
Set PPpres = PPApp.Presentations.Open(pptx)

'************************************************************
'prevent PowerPoint 2013 from losing focus and returning
'"shapes (unknown member) invalid request. the specified data type is unavailable"
'- Run-time error -2147188160 (80048240):View (unknown member) error
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
PPApp.Activate
PPApp.ActiveWindow.ViewType = ppViewNormal
PPApp.ActiveWindow.Panes(2).Activate 'standard ppt view
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each Cht In ActiveWorkbook.Charts
    i = Cht.Index
    Cht.ChartArea.Copy

    Select Case i
        Case 5: PPpres.Slides(2).Shapes.PasteSpecial ppPasteEnhancedMetafile
        Case 6: PPpres.Slides(3).Shapes.PasteSpecial ppPasteEnhancedMetafile
        Case 7: PPpres.Slides(15).Shapes.PasteSpecial ppPasteEnhancedMetafile
        Case 8: PPpres.Slides(15).Shapes.PasteSpecial ppPasteEnhancedMetafile
    End Select
    
Next Cht

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'copy HC charts from Dashboard tab (This is .ChartObjects collection and Object must always be activated first)
Worksheets("Dashboard").Activate
Set Country = ActiveSheet.Range("C8")

    For Each Cell In ActiveSheet.Range("U17:AD17")
        Country = Cell.Value
        
        Set ChtObj = ActiveSheet.ChartObjects("HC sub all MRUs")
        ChtObj.Activate
        ActiveChart.ChartArea.Copy

        Select Case Country
            Case "CEE&I": PPpres.Slides(4).Shapes.PasteSpecial ppPasteEnhancedMetafile
            Case "FRA": PPpres.Slides(5).Shapes.PasteSpecial ppPasteEnhancedMetafile
            Case "GER": PPpres.Slides(6).Shapes.PasteSpecial ppPasteEnhancedMetafile
            Case "GWE": PPpres.Slides(7).Shapes.PasteSpecial ppPasteEnhancedMetafile
            Case "IBE": PPpres.Slides(8).Shapes.PasteSpecial ppPasteEnhancedMetafile
            Case "ITA": PPpres.Slides(9).Shapes.PasteSpecial ppPasteEnhancedMetafile
            Case "MEMA": PPpres.Slides(10).Shapes.PasteSpecial ppPasteEnhancedMetafile
            Case "RUS": PPpres.Slides(11).Shapes.PasteSpecial ppPasteEnhancedMetafile
            Case "UKI": PPpres.Slides(12).Shapes.PasteSpecial ppPasteEnhancedMetafile
            Case "HQ": PPpres.Slides(13).Shapes.PasteSpecial ppPasteEnhancedMetafile
        End Select
    Next Cell

''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''copy the financial charts from Dashboard tab (This is .ChartObjects collection and Object must always be activated first)
'
'Worksheets("Dashboard").Activate
'With ActiveSheet
'
''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~CEE&I~~~~~~~~~~~~~~~~~~~~~
'Range("BZ8").Value = "CEE&I"
'.ChartObjects("sub A$ all MRUs").Activate 'actual dollars
'ActiveChart.ChartArea.Copy
'PPpres.Slides(16).Shapes.PasteSpecial ppPasteEnhancedMetafile
'
'Range("CT6").Value = "CEE&I"
'.ChartObjects("sub C$ all MRUs").Activate 'constant dollars
'ActiveChart.ChartArea.Copy
'PPpres.Slides(16).Shapes.PasteSpecial ppPasteEnhancedMetafile
'
''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~FRA~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Range("BZ8").Value = "FRA"
'.ChartObjects("sub A$ all MRUs").Activate 'actual dollars
'ActiveChart.ChartArea.Copy
'PPpres.Slides(17).Shapes.PasteSpecial ppPasteEnhancedMetafile
'
'Range("CT6").Value = "FRA"
'.ChartObjects("sub C$ all MRUs").Activate 'constant dollars
'ActiveChart.ChartArea.Copy
'PPpres.Slides(17).Shapes.PasteSpecial ppPasteEnhancedMetafile
'
''~~~~~~~~~~~~~~~~~~~~~~~~~~~~GER~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Range("BZ8").Value = "GER"
'.ChartObjects("sub A$ all MRUs").Activate 'actual dollars
'ActiveChart.ChartArea.Copy
'PPpres.Slides(18).Shapes.PasteSpecial ppPasteEnhancedMetafile
'
'Range("CT6").Value = "GER"
'.ChartObjects("sub C$ all MRUs").Activate 'constant dollars
'ActiveChart.ChartArea.Copy
'PPpres.Slides(18).Shapes.PasteSpecial ppPasteEnhancedMetafile
'
''~~~~~~~~~~~~~~~~~~~~~~~~~~~~GWE~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Range("BZ8").Value = "GWE"
'.ChartObjects("sub A$ all MRUs").Activate 'actual dollars
'ActiveChart.ChartArea.Copy
'PPpres.Slides(19).Shapes.PasteSpecial ppPasteEnhancedMetafile
'
'Range("CT6").Value = "GWE"
'.ChartObjects("sub C$ all MRUs").Activate 'constant dollars
'ActiveChart.ChartArea.Copy
'PPpres.Slides(19).Shapes.PasteSpecial ppPasteEnhancedMetafile
'
''~~~~~~~~~~~~~~~~~~~~~~~~~~~~IBE~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Range("BZ8").Value = "IBE"
'.ChartObjects("sub A$ all MRUs").Activate 'actual dollars
'ActiveChart.ChartArea.Copy
'PPpres.Slides(20).Shapes.PasteSpecial ppPasteEnhancedMetafile
'
'Range("CT6").Value = "IBE"
'.ChartObjects("sub C$ all MRUs").Activate 'constant dollars
'ActiveChart.ChartArea.Copy
'PPpres.Slides(20).Shapes.PasteSpecial ppPasteEnhancedMetafile
'
''~~~~~~~~~~~~~~~~~~~~~~~~~~~~ITA~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Range("BZ8").Value = "ITA"
'.ChartObjects("sub A$ all MRUs").Activate 'actual dollars
'ActiveChart.ChartArea.Copy
'PPpres.Slides(21).Shapes.PasteSpecial ppPasteEnhancedMetafile
'
'Range("CT6").Value = "ITA"
'.ChartObjects("sub C$ all MRUs").Activate 'constant dollars
'ActiveChart.ChartArea.Copy
'PPpres.Slides(21).Shapes.PasteSpecial ppPasteEnhancedMetafile
'
''~~~~~~~~~~~~~~~~~~~~~~~~~~~~MEMA~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Range("BZ8").Value = "MEMA"
'.ChartObjects("sub A$ all MRUs").Activate 'actual dollars
'ActiveChart.ChartArea.Copy
'PPpres.Slides(22).Shapes.PasteSpecial ppPasteEnhancedMetafile
'
'Range("CT6").Value = "MEMA"
'.ChartObjects("sub C$ all MRUs").Activate 'constant dollars
'ActiveChart.ChartArea.Copy
'PPpres.Slides(22).Shapes.PasteSpecial ppPasteEnhancedMetafile
'
''~~~~~~~~~~~~~~~~~~~~~~~~~~~~RUS~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Range("BZ8").Value = "MEMA"
'.ChartObjects("sub A$ all MRUs").Activate 'actual dollars
'ActiveChart.ChartArea.Copy
'PPpres.Slides(23).Shapes.PasteSpecial ppPasteEnhancedMetafile
'
'Range("CT6").Value = "MEMA"
'.ChartObjects("sub C$ all MRUs").Activate 'constant dollars
'ActiveChart.ChartArea.Copy
'PPpres.Slides(23).Shapes.PasteSpecial ppPasteEnhancedMetafile
'
''~~~~~~~~~~~~~~~~~~~~~~~~~~~~UKI~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Range("BZ8").Value = "UKI"
'.ChartObjects("sub A$ all MRUs").Activate 'actual dollars
'ActiveChart.ChartArea.Copy
'PPpres.Slides(24).Shapes.PasteSpecial ppPasteEnhancedMetafile
'
'Range("CT6").Value = "UKI"
'.ChartObjects("sub C$ all MRUs").Activate 'constant dollars
'ActiveChart.ChartArea.Copy
'PPpres.Slides(24).Shapes.PasteSpecial ppPasteEnhancedMetafile
'
''~~~~~~~~~~~~~~~~~~~~~~~~~~~~HQ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Range("BZ8").Value = "HQ"
'.ChartObjects("sub A$ all MRUs").Activate 'actual dollars
'ActiveChart.ChartArea.Copy
'PPpres.Slides(25).Shapes.PasteSpecial ppPasteEnhancedMetafile
'
'Range("CT6").Value = "HQ"
'.ChartObjects("sub C$ all MRUs").Activate 'constant dollars
'ActiveChart.ChartArea.Copy
'PPpres.Slides(25).Shapes.PasteSpecial ppPasteEnhancedMetafile
'
'End With

PPpres.Save
End Sub

