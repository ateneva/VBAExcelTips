Attribute VB_Name = "MsgBoxOptions"
Option Explicit

Sub Prompts()
Attribute Prompts.VB_ProcData.VB_Invoke_Func = " \n14"

'You can use the MsgBox function in two ways:

'----> To simply show a message to the user.
'In this case, you donft care'about the result returned by the function.

'----> To get a response from the user. In this case, you do care about the
'result returned by the function. The result depends on the button that
'the user clicks. See Sub RenminderCommas()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'If you use the MsgBox function by itself, don’t include parentheses around the arguments.
'The following example simply displays a message and does not return a result.

MsgBox "Click OK to begin printing."

MsgBox ("Successfully Completed the Task."), vbInformation, "Example of vbInformation"
MsgBox "Report generation complete", vbInformation + vbOKOnly 'Chase the Tail Report --> Wendy's file

MsgBox ("Are you fresher?"), vbQuestion, "Example of vbQuestion"

MsgBox ("Input Data is not valid!"), vbExclamation, "Example of vbExclamation"
MsgBox ("Please enter valid Number!"), vbCritical, "Example of vbCritical"

''~~~~~~~~~~~~~~~~~~~~~~ok/cancel~~~~~~~~~~yes/no~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MsgBox ("Thanks for visiting Analysistabs!"), vbOKOnly, "Example of vbOKOnly"
MsgBox ("Did you paste special the values for the current Q?"), vbQuestion + vbOKCancel
MsgBox ("You are VBA Expert, is it True?"), vbOKCancel, "Example of vbOKCancel"

MsgBox ("Would you like to re-filter"), vbYesNo
MsgBox ("File already exists. Do you want to replace?"), vbYesNoCancel, "Example of vbYesNoCancel"

'Constant----------------------------------> Value-------------------> Description
'vbOKOnly ---------------------------------> 0 ----------------------> Display OK button only.
'vbOKCancel -------------------------------> 1 ----------------------> Display OK and Cancel buttons.
'vbAbortRetryIgnore -----------------------> 2 ----------------------> Display Abort, Retry, and Ignore buttons.
'vbYesNoCancel ----------------------------> 3 ----------------------> Display Yes, No, and Cancel buttons.
'vbYesNo ----------------------------------> 4 ----------------------> Display Yes and No buttons.
'vbRetryCancel ----------------------------> 5 ----------------------> Display Retry and Cancel buttons.
'vbCritical -------------------------------> 16 ---------------------> Display Critical Message icon.
'vbQuestion -------------------------------> 32 ---------------------> Display Warning Query icon.
'vbExclamation ----------------------------> 48 ---------------------> Display Warning Message icon.
'vbInformation ----------------------------> 64 ---------------------> Display Information Message icon.

'***********************************************************************************************************************
                                    'setting default buttons
'***********************************************************************************************************************

'vbDefaultButton1 -------------------------> 0 ----> First button is default.
'vbDefaultButton2 -------------------------> 256 --> Second button is default.
'vbDefaultButton3 -------------------------> 512 --> Third button is default.
'vbDefaultButton4 -------------------------> 768 --> Fourth button is default.
'vbSystemModal ----------------------------> 4096 --> All applications are suspended until the user responds (might not work under all conditions).
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'***********************************************************************************************************************
                                    'Constant Value Button Clicked
'***********************************************************************************************************************
                                    'vbOK ------------------------------> 1 OK
                                    'vbCancel --------------------------> 2 Cancel
                                    'vbAbort ---------------------------> 3 Abort
                                    'vbRetry ---------------------------> 4 Retry
                                    'vbIgnore --------------------------> 5 Ignore
                                    'vbYes -----------------------------> 6 Yes
                                    'vbNo ------------------------------> 7 No

End Sub



