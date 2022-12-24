Attribute VB_Name = "CellInR_ZapFormulasinRange"
Option Explicit

Sub ZapValuesUserInput()
Attribute ZapValuesUserInput.VB_ProcData.VB_Invoke_Func = "Z\n14"
Dim MyRange As Range
Dim Cell As Range
Dim prv As Variant

With ActiveSheet

On Error GoTo handler
Set MyRange = Application.InputBox(Prompt:="Please Select a Range", _
Title:="Choose Range to convert to values", Type:=8)

For Each Cell In MyRange.SpecialCells(xlCellTypeVisible)
    If Not IsEmpty(Cell) = True Then
    
        prv = Cell.Value
        If Cell.HasFormula = True Then Cell.Value = prv
    End If

Next Cell

handler: MsgBox ("Operation Cancelled or Completed")
End With
End Sub

Sub ZapValuesNoUserInput()

Dim Cell As Range

'paste special as values
For Each Cell In ActiveSheet.Range("AB2:AB" & ActiveSheet.UsedRange.Rows.Count)
    On Error Resume Next
    If Cell.Value <> "0" And Len(Cell) > 1 Then
    
        Cell.Activate
        ActiveCell.Copy
        ActiveCell.PasteSpecial xlPasteValuesAndNumberFormats
    End If
Next Cell

Application.CutCopyMode = False
End Sub
