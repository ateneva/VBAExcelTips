Attribute VB_Name = "FY15_Int_ExportPPT"
Option Explicit

Sub Export_PPT_Internal()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Angelina Teneva, Aug 2014
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Dim Sh As Shape
Dim PT As PivotTable
Dim PL As String

Dim PPApp As PowerPoint.Application
Set PPApp = GetObject(, "Powerpoint.Application") 'use if you are planning on having your ppt open

Dim PPpres As PowerPoint.Presentation
Set PPpres = PPApp.ActivePresentation

Dim PPS As Integer
Dim Wks As Worksheet

'prevent PowerPoint 2013 from losing focus
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
PPApp.Activate
PPApp.ActiveWindow.ViewType = ppViewNormal
PPApp.ActiveWindow.Panes(2).Activate

'*******************************************************************************************
Worksheets("Int Imp % cat").Activate 'export internal imports
With ActiveSheet

'put date stamp
Range("A1:O3").Copy 'date stamp on Slides
For PPS = 14 To 19
    PPpres.Slides(PPS).Shapes.PasteSpecial ppPasteEnhancedMetafile
Next PPS

For Each PT In ActiveSheet.PivotTables
    PL = PT.name
    PT.PivotSelect "", xlDataAndLabel, True
    Selection.Copy
    
    Select Case PL
        Case "DCC": PPpres.Slides(15).Shapes.PasteSpecial ppPasteEnhancedMetafile
        Case "DCC/IC": PPpres.Slides(16).Shapes.PasteSpecial ppPasteEnhancedMetafile
        Case "CFS": PPpres.Slides(17).Shapes.PasteSpecial ppPasteEnhancedMetafile
        Case "TC": PPpres.Slides(14).Shapes.PasteSpecial ppPasteEnhancedMetafile
        Case "SIS": PPpres.Slides(18).Shapes.PasteSpecial ppPasteEnhancedMetafile
        Case "NMC": PPpres.Slides(19).Shapes.PasteSpecial ppPasteEnhancedMetafile
    End Select

Next PT

End With
PPpres.Save

'****************************************************************************************************
'export internal exports
If ActiveWorkbook.Worksheets("Internal Export %").Visible = False Then Worksheets("Internal Export -new-").Visible = True
ActiveWorkbook.Worksheets("Internal Export %").Activate

With ActiveSheet

Range("A1:N3").Copy 'put a date stamp
For PPS = 23 To 28
    PPpres.Slides(PPS).Shapes.PasteSpecial ppPasteEnhancedMetafile
Next PPS

For Each PT In ActiveSheet.PivotTables
    PL = PT.name
    PT.PivotSelect "", xlDataAndLabel, True
    Selection.Copy
    
    Select Case PL
        Case "DCC": PPpres.Slides(24).Shapes.PasteSpecial ppPasteEnhancedMetafile
        Case "DCC/IC": PPpres.Slides(25).Shapes.PasteSpecial ppPasteEnhancedMetafile
        Case "CFS": PPpres.Slides(26).Shapes.PasteSpecial ppPasteEnhancedMetafile
        Case "TC": PPpres.Slides(23).Shapes.PasteSpecial ppPasteEnhancedMetafile
        Case "SIS": PPpres.Slides(27).Shapes.PasteSpecial ppPasteEnhancedMetafile
        Case "NMC": PPpres.Slides(28).Shapes.PasteSpecial ppPasteEnhancedMetafile
    End Select
Next PT

End With
PPpres.Save

End Sub
