Attribute VB_Name = "OutlookGetFiles"
Option Explicit

Sub GetImportFileName()
Dim Filt As String
Dim FilterIndex As Integer
Dim Title As String
Dim FileName As Variant
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Set up list of file filters
Filt = "Text Files (*.txt),*.txt," & "Lotus Files (*.prn),*.prn," & "Comma Separated Files (*.csv),*.csv," & _
"ASCII Files (*.asc),*.asc," & "All Files (*.*),*.*"

'Display *.* by default
FilterIndex = 5

'Set the dialog box caption
Title = "Select a File to Import"

'Get the file name
FileName = Application.GetOpenFilename(FileFilter:=Filt, FilterIndex:=FilterIndex, Title:=Title)

'Exit if dialog box canceled
If FileName = False Then
    MsgBox "No file was selected."
    Exit Sub
End If

'Display full path and name of the file
MsgBox "You selected " & FileName
End Sub

Sub GetAFolder()

With Application.FileDialog(msoFileDialogFolderPicker)
.InitialFileName = Application.DefaultFilePath & " \ "
.Title = "Select a location for the backup"
.Show

If .SelectedItems.Count = 0 Then
MsgBox "Canceled"
Else
MsgBox .SelectedItems(1)
End If

End With

End Sub
