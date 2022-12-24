Attribute VB_Name = "FY15_ImportExportPPT"
Option Explicit

Sub ExcelToPowerPoint_Open()
Attribute ExcelToPowerPoint_Open.VB_ProcData.VB_Invoke_Func = "A\n14"
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Angelina Teneva, Aug 2014
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Dim Sh As Shape
Dim PPApp As PowerPoint.Application
Dim PPpres As PowerPoint.Presentation
Dim PPS As Integer
Dim Wks As Worksheet

Dim PT As PivotTable
Dim PF As PivotField
Dim PF2 As PivotField
Dim PL As String

''Create a PP application and make it visible
Set PPApp = New PowerPoint.Application
PPApp.Visible = True

'Open the presentation you wish to copy to
Set PPpres = PPApp.Presentations.Open("C:\Users\Angelina\Documents\Import-Export Balance.pptm")

'prevent PowerPoint 2013 from losing focus and returning
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
PPApp.Activate
PPApp.ActiveWindow.ViewType = ppViewNormal
PPApp.ActiveWindow.Panes(2).Activate


'****************************************************************************************************************************
If ActiveWorkbook.Worksheets.Count = 9 Then Application.Run "PERSONAL.XLSB!Export_PPT_Internal" 'stored in FY14_Int_ExportPPT"

If ActiveWorkbook.Worksheets.Count = 8 Then 'check if it is import file
    Worksheets("Project Import (RD&CoE)").Activate
    
    With ActiveSheet
    Range("A1:N4").Copy
    
    For PPS = 2 To 12 Step 2
        PPpres.Slides(PPS).Shapes.PasteSpecial ppPasteEnhancedMetafile
    Next PPS
    
    For Each PT In ActiveSheet.PivotTables
        PL = PT.name
        PT.PivotSelect "", xlDataAndLabel, True
        Selection.Copy
    
        Select Case PL
            Case "TC": PPpres.Slides(2).Shapes.PasteSpecial ppPasteMetafilePicture
            Case "SIS": PPpres.Slides(12).Shapes.PasteSpecial ppPasteEnhancedMetafile
            Case "CFS": PPpres.Slides(8).Shapes.PasteSpecial ppPasteEnhancedMetafile
            Case "DCC/IC": PPpres.Slides(6).Shapes.PasteSpecial ppPasteEnhancedMetafile
            Case "DCC": PPpres.Slides(4).Shapes.PasteSpecial ppPasteEnhancedMetafile
            Case "NMC": PPpres.Slides(10).Shapes.PasteSpecial ppPasteEnhancedMetafile
        End Select
    Next PT
    
    End With
End If

'**********************************************************************************************************************************
If ActiveWorkbook.Worksheets.Count >= 10 Then
    Worksheets("Export Pivot % breakdown").Activate 'check if it is Export file
    
    With ActiveSheet
    Range("A1:L4").Copy
    
    For PPS = 1 To 11 Step 2
        PPpres.Slides(PPS).Shapes.PasteSpecial ppPasteEnhancedMetafile
    Next PPS
    
    For Each PT In ActiveSheet.PivotTables
        PL = PT.name
        PT.PivotSelect "", xlDataAndLabel, True
        Selection.Copy
    
        Select Case PL
            Case "TC": PPpres.Slides(1).Shapes.PasteSpecial ppPasteEnhancedMetafile
            Case "SIS": PPpres.Slides(11).Shapes.PasteSpecial ppPasteEnhancedMetafile
            Case "CFS": PPpres.Slides(7).Shapes.PasteSpecial ppPasteEnhancedMetafile
            Case "DCC/IC": PPpres.Slides(5).Shapes.PasteSpecial ppPasteEnhancedMetafile
            Case "DCP": PPpres.Slides(3).Shapes.PasteSpecial ppPasteEnhancedMetafile
            Case "NMC": PPpres.Slides(9).Shapes.PasteSpecial ppPasteEnhancedMetafile
        End Select
    
    Next PT
    
    End With
End If

Application.CutCopyMode = False
PPpres.Save
PPpres.Close

End Sub

