Attribute VB_Name = "AppOutlookToExcel"
Option Explicit

Sub SaveAttachments_RMR()
    
    Dim myOlapp As Outlook.Application
    Dim myNameSpace As Outlook.Namespace
    Dim myFolder As Outlook.MAPIFolder
    Dim myItem As Outlook.MailItem
    Dim myAttachment As Outlook.Attachment
    Dim FileName As String
     
    Set myOlapp = CreateObject("Outlook.Application")
    Set myNameSpace = myOlapp.GetNamespace("MAPI")
    Set myFolder = myNameSpace.GetDefaultFolder(olFolderInbox)
    Set myFolder = myFolder.Folders("RMR")
     
   '*****************************************
For Each myItem In myFolder.items
        
If myItem.Attachments.Count <> 0 Then
    For Each myAttachment In myItem.Attachments
             
        FileName = "C:\Users\Angelina\InputData\" & myAttachment.FileName
        myAttachment.SaveAsFile FileName
    Next
End If
  
Next

Application.Run "OpenFileOrFolderOrWebsite"
      
End Sub

Sub SaveAttachments_BW()
    
    Dim myOlapp As Outlook.Application
    Dim myNameSpace As Outlook.Namespace
    Dim myFolder As Outlook.MAPIFolder
    Dim myItem As Outlook.MailItem
    Dim myAttachment As Outlook.Attachment
    Dim i As Long
    Dim FileName As String
    Dim Subject As String
     
    Set myOlapp = CreateObject("Outlook.Application")
    Set myNameSpace = myOlapp.GetNamespace("MAPI")
    Set myFolder = myNameSpace.GetDefaultFolder(olFolderInbox)
    Set myFolder = myFolder.Folders("BW")
     
   '*****************************************
For Each myItem In myFolder.items
        
If myItem.Attachments.Count <> 0 Then
    For Each myAttachment In myItem.Attachments
    
    Subject = myItem.Subject
    myAttachment.SaveAsFile "C:\Users\Angelina\Checks\" & Subject & ".zip"
                
    i = i + 1
    Next
End If
  
Next
End Sub

Sub ReportEmail()

Dim shReport As Worksheet
Dim OutApp As Outlook.Application
Dim sess As Outlook.Namespace
Dim Fld As Outlook.MAPIFolder

Set OutApp = CreateObject("Outlook.application")
Set sess = OutApp.Session

'here the folder were you read mails items
'this is your receipt default folder
Set Fld = sess.GetDefaultFolder(olFolderInbox)

'you can also use this:
'Set Fld = sess.Folders("yourfolder")

Set shReport = ThisWorkbook.Worksheets(1)

LookInFolder Fld

Set OutApp = Nothing
End Sub

Sub LookInFolder(Fld As MAPIFolder)
Dim itm As MailItem
Dim subFld As MAPIFolder
Dim Cell As Range
Dim i As Integer

For Each itm In Fld.items
    Set Cell = shReport.[A65535].End(xlUp).Offset(1, 0)
    
    Cell.Value = itm.SenderName
    Cell.Offset(0, 1) = itm.Subject
    Cell.Offset(0, 2) = itm.ReceivedTime
Next itm

For Each subFld In Fld.Folders
    LookInFolder subFld
Next subFld

End Sub
