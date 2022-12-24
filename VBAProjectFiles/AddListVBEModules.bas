Attribute VB_Name = "AddListVBEModules"
Option Explicit

Sub ShowProjectComponents()
'written by John WalkenBach
'the will put in excel all the names of all the modules found in a project

Dim VBP As VBIDE.VBProject
Dim VBC As VBComponent
Dim row As Long
Set VBP = Application.VBE.VBProjects("Angelina")
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~returns in Excel the avaialble VBE Components in a Workbook --> will help me build MyLibrary
'Write headers
Cells.ClearContents
Range("A1:C1") = Array("Name", "Type", "Code Lines")
Range("A1:C1").Font.Bold = True
row = 1

'Loop through the VB Components
For Each VBC In VBP.VBComponents
    row = row + 1
    
    'name
    Cells(row, 1) = VBC.name
    
    'Type
    Select Case VBC.Type
        Case vbext_ct_StdModule: Cells(row, 2) = "Module"
        Case vbext_ct_ClassModule: Cells(row, 2) = "Class Module"
        Case vbext_ct_MSForm: Cells(row, 2) = "UserForm"
        Case vbext_ct_Document: Cells(row, 2) = "Document Module"
    End Select

'Lines of code
Cells(row, 3) = VBC.CodeModule.CountOfLines
Next VBC

End Sub
