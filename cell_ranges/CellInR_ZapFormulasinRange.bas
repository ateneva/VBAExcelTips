
Sub values()
Attribute values.VB_ProcData.VB_Invoke_Func = "Z\n14"
Dim MyRange As Range
Dim Cell As Range
Dim prv As Variant

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva
'https://datageeking.wordpress.com/2018/05/15/quickly-zap-formulas-based-on-a-cell-value/'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

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
