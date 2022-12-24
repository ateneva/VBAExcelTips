Attribute VB_Name = "MsgBoxGetAnAnswer"
Option Explicit

Private Sub Workbook_AfterSave(ByVal Success As Boolean)

Dim Ans As Integer
Ans = MsgBox("Would you like to update today's date & QTD conditional formatting now?", vbYesNo)

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

Private Sub Workbook_BeforeClose(Cancel As Boolean)

Dim Ans As Integer, Ans2 As Integer, Ans3 As Integer
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Did you check for commas reminder

Ans = MsgBox("Did you check for commas?", vbYesNo)

Select Case Ans
    Case vbYes: ThisWorkbook.Close
    Case vbNo: MsgBox ("Please, check for commas")
End Select

End Sub

Sub GetAnswer4()

Dim Msg As String, Title As String
Dim Config As Integer, Ans As Integer
'get an answer + explanation

    Msg = "Do you want to process the monthly report?"
    Msg = Msg & vbNewLine & vbNewLine
    Msg = Msg & "Processing the monthly report will"
    Msg = Msg & "take approximately 15 minutes. It"
    Msg = Msg & "will generate a 30-page report for"
    Msg = Msg & "all sales offices for the current"
    Msg = Msg & "month."
    
    Title = "XYZ Marketing Company"
    Config = vbYesNo + vbQuestion
    
    Ans = MsgBox(Msg, Config, Title)
    If Ans = vbYes Then RunReport
    
End Sub
