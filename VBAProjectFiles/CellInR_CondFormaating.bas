Attribute VB_Name = "CellInR_CondFormaating"
Option Explicit

Sub PasteCondFormatting()
Attribute PasteCondFormatting.VB_ProcData.VB_Invoke_Func = " \n14"

Dim Cell As Range
'--------------------------------------------

ActiveSheet.Range("AH3:AS3").Copy

For Each Cell In ActiveSheet.Range("AG3:AG86")
    Cell.Offset(0, 1).PasteSpecial Paste:=xlPasteFormats

Next Cell
End Sub
