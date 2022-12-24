Attribute VB_Name = "ForEachWbkInWbksOpenSaveClose"
Option Explicit

Sub Check_WbkIsOpen()
Dim Wbk As Workbook
Dim WbkName As String
WbkName = InputBox("Please enter file name")
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
For Each Wbk In Application.Workbooks
    If Wbk.name Like "*" & WbkName & "*" Then
    MsgBox WbkName & " is open"
   
    Else
    MsgBox WbkName & " is not open. Please open before proceeding"
    End If
    
Next Wbk
End Sub

Sub SaveAllWorkbooks()
Dim Wbk As Workbook
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each Wbk In Application.Workbooks
    If Wbk.path <> "" Then Wbk.Save
Next Wbk

End Sub

Sub CloseAllWorksbooks()
Dim Wbk As Workbook
'~~~~~~~~~~~~~~~~~~

For Each Wbk In Application.Workbooks
    If Wbk.name <> ThisWorkbook.name Then
    Wbk.Close savechanges:=True
    End If
Next Wbk

ThisWorkbook.Close savechanges:=True
End Sub
