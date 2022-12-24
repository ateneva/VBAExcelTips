Attribute VB_Name = "InputBoxFunction"
Option Explicit

Sub EnterValues()

Dim num As Integer
Dim float As Double
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

num = InputBox("Please enter a number")
    ActiveCell = num

float = InputBox("Please enter a decimal")
    ActiveCell = float

'unlike Python's raw input method, VBA allows you to declare beforehand the type of input you want - e.g. - integer, double, string

End Sub

Sub GettingUserDefinedString()

Dim Ans As String
Dim Ans2 As Integer
'~~~~~~~~~~~~~~~~~~~~~~~
Worksheets("Export Costs Analysis").Activate

Ans2 = MsgBox("Would you like to refilter pivot tables", vbYesNo)

Select Case Ans2
Case vbYes
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'apply filter for the latest month to all pivot tables as given by user --> might also be a keyword
    Ans = InputBox("Please enter latest fiscal period in the format Period nn yyyy")

'this is a inputbox function that always returns a string

Case vbNo
    MsgBox ("Remember to re-filter before closing")
End Select

End Sub


