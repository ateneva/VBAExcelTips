Attribute VB_Name = "Module2"
Option Explicit

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Range("A6").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Rows("18406:19853").Select
    Range("BD19853").Activate
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Range("A5387").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.EntireRow.Delete
    ActiveWorkbook.Save
End Sub
