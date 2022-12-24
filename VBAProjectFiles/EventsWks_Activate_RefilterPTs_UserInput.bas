Attribute VB_Name = "EventsWks_Activate"
Option Explicit

Private Sub Worksheet_Activate()

Dim Ans As String
Dim Ans2 As Integer
'~~~~~~~~~~~~~~~~~~~~~~~
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
