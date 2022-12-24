Attribute VB_Name = "AddVBEModulefromFolder"
Option Explicit

Public Sub ImportModules()
    Dim wkbTarget As Excel.Workbook
    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.file
    Dim szTargetWorkbook As String
    Dim szImportPath As String
    Dim szFileName As String
    Dim cmpComponents As VBIDE.VBComponents
    
    'needed to import in your personal book
    Dim vbProj As VBIDE.VBProject
    Set vbProj = Application.VBE.VBProjects("Angelina")

'    If ActiveWorkbook.name = ThisWorkbook.name Then
'
'        MsgBox "Select another destination workbook" & "Not possible to import in this workbook "
'        Exit Sub
'    End If

    'Get the path to the folder with modules
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Import Folder not exist"
        Exit Sub
    End If

'    ''' NOTE:
     '''This workbook must be open in Excel.
'    szTargetWorkbook = ActiveWorkbook.name
'    Set wkbTarget = Application.Workbooks(szTargetWorkbook)
'
'    If wkbTarget.VBProject.protection = 1 Then
'    MsgBox "The VBA in this workbook is protected," & _
'        "not possible to Import the code"
'    Exit Sub
'    End If

    ''' NOTE: Path where the code modules are located.
    szImportPath = FolderWithVBAProjectFiles  'this is the UDF function specified below; change path there
        
    Set objFSO = New Scripting.FileSystemObject
    If objFSO.GetFolder(szImportPath).Files.Count = 0 Then
       MsgBox "There are no files to import"
       Exit Sub
    End If

'    'Delete all modules/Userforms from the ActiveWorkbook
'    Call DeleteVBAModulesAndUserForms

    'Set cmpComponents = wkbTarget.VBProject.VBComponents 'to add to workbook
    Set cmpComponents = Application.VBE.VBProjects("Angelina").VBComponents
    
    ''' Import all the code modules in the specified path
    ''' to the ActiveWorkbook.
    For Each objFile In objFSO.GetFolder(szImportPath).Files
    
        If (objFSO.GetExtensionName(objFile.name) = "cls") Or _
            (objFSO.GetExtensionName(objFile.name) = "frm") Or _
            (objFSO.GetExtensionName(objFile.name) = "bas") Then
            cmpComponents.Import objFile.path
        End If
        
    Next objFile
       
    MsgBox "Import is ready"
End Sub

Function FolderWithVBAProjectFiles() As String
    Dim WshShell As Object
    Dim FSO As Object
    Dim SpecialPath As String

    Set WshShell = CreateObject("WScript.Shell")
    Set FSO = CreateObject("scripting.filesystemobject")
        
    SpecialPath = WshShell.SpecialFolders("MyDocuments")
    'SpecialPath = WshShell.SpecialFolders("C:\Users\hp\Documents\") will not work!!!
    
    '.SpecialFolders is a collection of folders - you can only access the following folders
    '    • AllUsersDesktop, • Desktop • Favorites, * MyDocuments
    'import from DropBox Sync Desktop App, Google Drive Sync App or any other folder outside the ones above will not work

    If Right(SpecialPath, 1) <> "\" Then
        SpecialPath = SpecialPath & "\"
    End If
    
    If FSO.FolderExists(SpecialPath & "VBAProjectFiles") = False Then
        On Error Resume Next
        MkDir SpecialPath & "VBAProjectFiles"
        On Error GoTo 0
    End If
    
    If FSO.FolderExists(SpecialPath & "VBAProjectFiles") = True Then
        FolderWithVBAProjectFiles = SpecialPath & "VBAProjectFiles"
    Else
        FolderWithVBAProjectFiles = "Error"
    End If
    
End Function

Function DeleteVBAModulesAndUserForms()
        Dim vbProj As VBIDE.VBProject
        Dim vbComp As VBIDE.VBComponent
        
        Set vbProj = ActiveWorkbook.VBProject
        
        For Each vbComp In vbProj.VBComponents
            If vbComp.Type = vbext_ct_Document Then
                'Thisworkbook or worksheet module
                'We do nothing
            Else
                vbProj.VBComponents.Remove vbComp
            End If
        Next vbComp
End Function

