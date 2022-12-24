Attribute VB_Name = "AppExcelToOutlook"
Option Explicit

Sub SendMail()

Dim MyOutlook As Outlook.Application
Dim MyMail As Outlook.MailItem
Dim objOutlook As Object

Dim strbody As String
Dim rangebody As Range
Dim greeting As Range
Set greeting = ActiveWorkbook.Worksheets("Sheet2").Range("E1:E2")
Dim Cell As Range
Dim name As String
'******************************************

ActiveWorkbook.Worksheets("Sheet2").Activate
For Each Cell In ActiveSheet.Range("A5:A" & ActiveSheet.UsedRange.Rows.Count)

    Cell.Activate
    'name = ActiveCell.Value
    'Range("E1").Value = "Dear " & name & Chr(44)
    Set rangebody = Range(Cells(ActiveCell.row, "C"), Cells(ActiveCell.row, "J"))
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Set MyOutlook = New Outlook.Application
    Set MyMail = MyOutlook.CreateItem(olMailItem)
                 
    MyMail.To = Cells(ActiveCell.row, "B")
    MyMail.Subject = "Daily tasks_" & Format(Date, "dd-mmm-yy")
    MyMail.Display 'places the cursor in part of your e-mail (if after body, places the cursor in your body
    
    rangebody.Copy
    SendKeys "^({v})", True 'equals Ctrl+V command

Next Cell
End Sub

Sub ChritmasWishes()

Dim MyOutlook As Outlook.Application
Dim MyMail As Outlook.MailItem
Dim objOutlook As Object
Dim myAttachments As Outlook.Attachments

Dim strbody As Range
Dim Cell As Range
Dim rngTo As Range

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'loop through all recipient             'Christmas wishes.xlsm file
For Each Cell In ThisWorkbook.Worksheets("data").Range("D2:D19")
        Set rngTo = Cell.Offset(0, 3)
        Cell.Copy

        Worksheets("design").Activate
        With ActiveSheet
                Range("K11").PasteSpecial xlPasteValues
                ActiveSheet.Range("G4:Q34").Copy
        End With
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Set MyOutlook = New Outlook.Application
        Set MyMail = MyOutlook.CreateItem(olMailItem) '' This creates the e-mail
        
        MyMail.To = rngTo
        MyMail.Subject = "Merry Christmas"
        MyMail.Display 'must be placed after To & Subject otherwise it pastes te copied content in the To line
        
        SendKeys "^({a})", True 'selects signature
        SendKeys "({DEL})", True 'removes signature
        SendKeys "^({v})", True '= Ctrl+V; pastes range from Excel
        
        MyMail.Send 'if you want to send automatically

Next Cell

End Sub

Sub LineManagers()

Dim MyOutlook As Outlook.Application
Dim MyMail As Outlook.MailItem
Dim objOutlook As Object
Dim myAttachments As Outlook.Attachments

Dim strbody As String

Dim saves As String
Dim attach As String
'****************************************************************************************
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'prepare file
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Worksheets("Default").Activate
Range("A14:B" & ActiveSheet.UsedRange.Rows.Count).Copy
 
Workbooks.Add
Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

'saves to avoid running error if the new book is not Book 1
ActiveWorkbook.saveas FileName:="C:\Users\TENEVAA\Documents\TS EMEA\I am Responsible For\Import-Export Balance\Direct Managers List.xlsx" _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
   
'adjust column width in new file
Columns("B:I").EntireColumn.AutoFit
Range("A1:I1").Font.Bold = True

ActiveWorkbook.Save
ActiveWorkbook.Close

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'send as an attachment in e-mail
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
saves = "C:\Users\Angelina\DirectManagersList.xlsx"
attach = "DirectManagersList.xlsx"

Set MyOutlook = New Outlook.Application
Set MyMail = MyOutlook.CreateItem(olMailItem)                   'This creates the e-mail
Set myAttachments = MyMail.Attachments

MyMail.Body = "Hi, please find attached latest direct managers and employeee CL extract from RMR"

myAttachments.Add saves, olByValue, 1, attach                   'adding RMR extract

MyMail.Display

MyMail.To = "angelina@abv.bg"
MyMail.Subject = "Direct Manager Lists" & Format(Date, "dd-mmm-yyyy")
MyMail.Send

End Sub

Sub EmpName()
''stoted in TS Headcount Template - EmpNamemod module

Dim empnamepath As String
empnamepath = ThisWorkbook.Worksheets("MACROS").Range("A27")

Workbooks.Open FileName:=empnamepath, ReadOnly:=False, UpdateLinks:=False
ActiveWorkbook.Worksheets("HRBI").Activate

With ActiveSheet
    Range("A1:K" & .UsedRange.Rows.Count).Clear
    Range("A1").Activate
End With
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ThisWorkbook.Worksheets("data").Activate
With ActiveSheet
    Range("B5:C" & .UsedRange.Rows.Count).Copy
End With
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workbooks("Emp Names.xlsx").Activate
ActiveWorkbook.Worksheets("HRBI").Activate

With ActiveSheet
.Paste
Columns("B:B").EntireColumn.AutoFit

Application.CutCopyMode = False
Columns("B:B").Select
    Selection.TextToColumns Destination:=Range("B1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True

Range("D1").Formula = "=TRIM(CONCATENATE(C1,"" "",B1))"
Range("D1:D" & .UsedRange.Rows.Count).FillDown
Range("D1:D" & .UsedRange.Rows.Count).Copy
Range("D1").PasteSpecial xlPasteValues
Columns("D:D").EntireColumn.AutoFit

End With
ActiveWorkbook.saveas FileName:=empnamepath
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workbooks("Emp Names.xlsx").RefreshAll 'refreshes the pivot table with new/left employees for Dani
Workbooks("Emp Names.xlsx").Save
Workbooks("Emp Names.xlsx").Close

End Sub

Sub EmpNameAttach()
''stoted in TS Headcount Template - EmpNamemod module

Dim MyOutlook As Outlook.Application
Dim MyMail As Outlook.MailItem
Dim objOutlook As Object
Dim myAttachments As Outlook.Attachments

Dim strbody As String
Dim strbody1 As String

Dim attach As String
attach = ThisWorkbook.Worksheets("MACROS").Range("A27").Value
'****************************************************************************************

Set MyOutlook = New Outlook.Application
Set MyMail = MyOutlook.CreateItem(olMailItem) '' This creates the e-mail
Set myAttachments = MyMail.Attachments

strbody = "Please find attached the TS Consulting employee names in the needed format"

strbody1 = "Dear all," & vbNewLine & vbNewLine & _
              "Please find attached the monthly headcount pursuit view" & vbNewLine & _
              "Should you have any questions, please do not hesitate to come back to me"


MyMail.Body = strbody 'this gives it the body

myAttachments.Add attach, olByValue, 1
MyMail.Display

MyMail.To = "danail.yordanov@hp.com; teodor.dilov@hp.com; petar.boskov@hp.com"
MyMail.CC = "Federspiel, Hugo; Mircheva, Iva"
MyMail.Subject = "Emp Names_ " & ThisWorkbook.Worksheets("MACROS").Range("E1").Value
MyMail.Send

End Sub
Sub StartEmail()
Dim Addr As String
Dim Result As Long
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'opens the default mail client and sends an e-mail to the recipient
Addr = "mailto: bgates@ microsoft.com"
Result = ShellExecute(0&, vbNullString, Addr, vbNullString, vbNullString, vbNormalFocus)
If Result < 32 Then MsgBox "Error"

End Sub

Sub SendEmail()
'Uses early binding
'Requires a reference to the Outlook Object Library
Dim OutlookApp As Outlook.Application
Dim MItem As Outlook.MailItem

Dim Cell As Range
Dim Subj As String

Dim EmailAddr As String
Dim Recipient As String
Dim Bonus As String
Dim Msg As String
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Create Outlook object
Set OutlookApp = New Outlook.Application

'Loop through the rows
For Each Cell In Columns("B").Cells.SpecialCells(xlCellTypeConstants)

If Cell.Value Like "*@*" Then

    'Get the data
    Subj = "Your Annual Bonus"
    Recipient = Cell.Offset(0, -1).Value
    EmailAddr = Cell.Value
    Bonus = Format(Cell.Offset(0, 1).Value, "$0,000.")
    
    'Compose Message
    Msg = "Dear " & Recipient & vbCrLf & vbCrLf
    Msg = Msg & "I am pleased to inform you that your annual bonus is "
    Msg = Msg & Bonus & vbCrLf & vbCrLf
    Msg = Msg & "William Rose" & vbCrLf
    Msg = Msg & "President"
    
    'Create Mail Item and send it
    Set MItem = OutlookApp.CreateItem(olMailItem)
    With MItem
    .To = EmailAddr
    .Subject = Subj
    .Body = Msg
    .Send
    End With

End If

Next Cell
End Sub

Sub SendAsPDF()
' Uses early binding
' Requires a reference to the Outlook Object Library
Dim OutlookApp As Outlook.Application
Dim MItem As Object
Dim Recipient As String, Subj As String
Dim Msg As String, fname As String
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

' Message details
    Recipient = "myboss@xrediyh.com"
    Subj = "Sales figures"
    Msg = "Hey boss, here’s the PDF file you wanted."
    Msg = Msg & vbNewLine & vbNewLine & " - Frank"
    fname = Application.DefaultFilePath & " \ " & ActiveWorkbook.name & ".pdf"

' Create the attachment
ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=fname

' Create Outlook object
Set OutlookApp = New Outlook.Application

' Create Mail Item and send it
Set MItem = OutlookApp.CreateItem(olMailItem)
    With MItem
    .To = Recipient
    .Subject = Subj
    .Body = Msg
    .Attachments.Add fname
    .Save 'to Drafts folder
    '.Send
    End With
    
Set OutlookApp = Nothing
Kill fname ' Delete the file

End Sub
