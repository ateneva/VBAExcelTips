Attribute VB_Name = "AddVBEModule"
Option Explicit

Sub AddModuleToProject()
'
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
    ''Set VBProj = ActiveWorkbook.VBProject
    Set vbComp = vbProj.VBComponents.Add(vbext_ct_StdModule)
    vbComp.name = "NewModule"
End Sub
