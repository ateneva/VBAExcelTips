Attribute VB_Name = "WksBeautify"
Option Explicit

Sub RemoveGridlines()
Dim Wks As Worksheet
'--------------------------------
'written by Angelina Teneva, 2014
'--------------------------------

For Each Wks In ActiveWorkbook.Worksheets

If Wks.Visible = True Then Wks.Activate
    
    With ActiveWindow
        .DisplayGridlines = False
        .Zoom = 80
    End With

Next Wks

End Sub

Sub Group_Rows() 'Angelina Teneva, November 2012

Dim Wks As Worksheet

For Each Wks In ThisWorkbook.Worksheets
If Wks.Visible = True Then Wks.Activate

With ActiveSheet

    Rows("13:35").Select
    Selection.Rows.Group
    ActiveSheet.Outline.ShowLevels RowLevels:=1
    
    .Outline.ShowLevels RowLevels:=2 'expands outlines

    .Outline.ShowLevels RowLevels:=1 'collapses outlines
    .Outline.ShowLevels RowLevels:=2 'expands outlines

    .Outline.ShowLevels RowLevels:=2 'hide 3P Labour & HP Labour
    .Outline.ShowLevels ColumnLevels:=1 'hide old data to make conditonal fromatting rules visible

    .Outline.ShowLevels RowLevels:=0, ColumnLevels:=2 'can be applied to both rows and columns at the same time
    
End With
Next Wks
