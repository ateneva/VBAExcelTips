Attribute VB_Name = "InputBoxMethod"
Option Explicit

Sub SelectRange()

Dim MyRange As Range
'~~~~~~~~~~~~~~~~~~~~~'this is a metod in the application.object
With ActiveSheet
    Set MyRange = Application.InputBox(Prompt:="Please Select a Range", Title:="Choose Range to convert to values", Type:=8)

        'Code Meaning
        '0 ---> A Formula
        '1 ---> A Number
        '2 ---> A string (text)
        '4 ---> A logical value (True or False)
        '8 ---> A cell reference, as a range object
        '16 --> An error value, such as #N/A
        '64 --> An array of values

End Sub

Sub GetUserRange()
Dim UserRange As Range
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Prompt = "Select a range for the random numbers."
Title = "Select a range"

On Error Resume Next
Set UserRange = Application.InputBox(Prompt:=Prompt, Title:=Title, Default:=ActiveCell.Address, Type:=8) 'Range selection

On Error GoTo 0
If UserRange Is Nothing Then
    MsgBox "Canceled."
    Else
    UserRange.Formula = "=RAND()"
End If

End Sub

Sub CopyChartOrRange()

Dim Ans As Integer
Dim inputdata As String
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Ans = MsgBox("Would you like to copy chart", vbYesNo)

Select Case Ans
Case vbYes
    ActiveSheet.ChartObjects("Chart 1").Copy
    ActiveWorkbook.Worksheets.Add.name = "1" 'pastes in a newly added sheets

Case vbNo
On Error GoTo notification

    inputdata = Application.InputBox("Enter the range you want to copy", Type:=8)
    ActiveWorkbook.Worksheets.Add.name = "2" 'pastes in a newly added sheet
End Select

notification:
MsgBox ("Range not selected")

Application.CutCopyMode = False
End Sub
