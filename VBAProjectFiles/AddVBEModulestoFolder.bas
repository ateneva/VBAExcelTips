Attribute VB_Name = "AddVBEModulestoFolder"
Option Explicit

Public Sub ExportModules()
    Dim bExport As Boolean
    Dim wkbSource As Excel.Workbook
    Dim szSourceWorkbook As String
    Dim Source As VBIDE.VBComponents
    Dim szExportPath As String
    Dim szFileName As String
    Dim cmpComponent As VBIDE.VBComponent
    
    ''code written by Ron De Bruin

    ''' The code modules will be exported in a folder named.
    ''' VBAProjectFiles in the Documents folder.
    ''' The code below create this folder if it not exist
    ''' or delete all files in the folder if it exist.
    
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Export Folder not exist"
        Exit Sub
    End If
    
    On Error Resume Next
        Kill FolderWithVBAProjectFiles & "\*.*"
    On Error GoTo 0

'    ''' NOTE: This workbook must be open in Excel if you're importing to a Workbook
    'szSourceWorkbook = ActiveWorkbook.name
    'Set wkbSource = Application.Workbooks(szSourceWorkbook) export from Wbk
    Set Source = Application.VBE.VBProjects("Angelina").VBComponents
    
'    If wkbSource.VBProject.protection = 1 Then
'    MsgBox "The VBA in this workbook is protected," & _
'        "not possible to export the code"
'    Exit Sub
'    End If
    
    szExportPath = FolderWithVBAProjectFiles & "\"
    
    For Each cmpComponent In Source ''.VBProject.VBComponents (add if you're exporting from a Wbk)
        
        bExport = True
        szFileName = cmpComponent.name

        ''' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule: szFileName = szFileName & ".cls"
            Case vbext_ct_MSForm: szFileName = szFileName & ".frm"
            Case vbext_ct_StdModule: szFileName = szFileName & ".bas"
            Case vbext_ct_Document
                ''' This is a worksheet or workbook object.
                ''' Don't try to export.
                bExport = False
        End Select
        
        If bExport Then
            ''' Export the component to a text file.
            cmpComponent.Export szExportPath & szFileName
            
        ''' remove it from the project if you want
        '''wkbSource.VBProject.VBComponents.Remove cmpComponent
                       
        If cmpComponent.name <> "Add*" Then Source.Remove cmpComponent
        
        End If
   
    Next cmpComponent

    MsgBox "Export is ready"
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

