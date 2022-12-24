Attribute VB_Name = "OutlookGetAttachments"
Option Explicit

Sub SaveAttachments()
    
    Dim myOlapp As Outlook.Application
    Dim myNameSpace As Outlook.Namespace
    Dim myFolder As Outlook.MAPIFolder
    Dim myItem As Outlook.MailItem
    Dim myAttachment As Outlook.Attachment
    Dim i As Long
    Dim FileName As String
    
    Set myOlapp = CreateObject("Outlook.Application")
    Set myNameSpace = myOlapp.GetNamespace("MAPI")
    Set myFolder = myNameSpace.GetDefaultFolder(olFolderInbox)
    Set myFolder = myFolder.Folders("RMR")
        
   '*****************************************
   
'upon receipt of the e-mail, a rule moves it into RMR folder and marks it as high importance and flags it as task
For Each myItem In myFolder.items
        
    If myItem.Attachments.Count <> 0 Then
        For Each myAttachment In myItem.Attachments
        FileName = "C:\Users\Angelina\InputData\" & myAttachment.FileName
        myAttachment.SaveAsFile FileName
        i = i + 1
        Next
    End If
    Next
  
'after saving the attachments, the code proceeds to clear the flags and reset the messages to normal importance upon which another rule activates,
'which moves them to Import-Export pst folder

For Each myItem In myFolder.items
    If myItem.IsMarkedAsTask Then myItem.ClearTaskFlag
    If myItem.Importance = olImportanceHigh Then myItem.Importance = olImportanceNormal
    myItem.Save

Next
   
End Sub
