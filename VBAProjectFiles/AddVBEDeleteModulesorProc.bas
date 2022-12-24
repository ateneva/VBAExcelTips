Attribute VB_Name = "AddVBEDeleteModulesorProc"
Option Explicit


Sub DeleteModule()
    Dim vbProj As VBIDE.VBProject
    Dim VBP As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    
    Set VBP = Application.VBE.VBProject("Angelina")
    Set vbProj = ActiveWorkbook.VBProject
    Set vbComp = vbProj.VBComponents("Module1") 'or Set VBComp = VBP.VBComponents("App")
    
    VBProject.VBComponents("OldName").name = "NewName" 'renames a module
    vbProj.VBComponents.Remove vbComp 'deletes a module
    VBP.VBComponents.Remove vbComp
    
End Sub

Sub DeleteProcedureFromModule()
    Dim vbProj As VBIDE.VBProject
    Dim VBP As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    Dim CodeMod As VBIDE.CodeModule
    Dim StartLine As Long
    Dim NumLines As Long
    Dim ProcName As String
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Set VBP = Application.VBE.VBProject("Angelina")
    Set vbProj = ActiveWorkbook.VBProject
    Set vbComp = vbProj.VBComponents("Module1") 'or Set VBComp = VBP.VBComponents("App")
    Set CodeMod = vbComp.CodeModule

    ProcName = "DeleteThisProc"
    With CodeMod
        StartLine = .ProcStartLine(ProcName, vbext_pk_Proc)
        NumLines = .ProcCountLines(ProcName, vbext_pk_Proc)
        .DeleteLines StartLine:=StartLine, Count:=NumLines
End With
End Sub
