Attribute VB_Name = "EventsWbkSave"
Option Explicit

Private Sub Workbook_AfterSave(ByVal Success As Boolean)

Dim Ans As Integer
Ans = MsgBox("Would you like to update today's date & QTD conditional formatting now?", vbYesNo)

'get an answer whether you want sth updated

Select Case Ans

Case vbYes:
Call Updated
Worksheets("EMEA").Activate

Call Edit_QTD_Conditional_formatting  'chnages the formula reference for QTD columns (stored in H2_conditionalformatting_QTD module)
Call ImportExports                    'edits conditional formatting for MTD, YTD import/exports (stored on H2_MTD_Export_Import_Formats module)
Call YTD_edit_perQ_formatting         'adds conditional formatting to the new Q as we enter & edits QTD colors for column H (depending on the Q
Call edit_HTD_imports_exports         'edit colour coding for HTD imports/exports

Case vbNo:
MsgBox ("Do not forget to update those before exporting to slides")

End Select
End Sub
