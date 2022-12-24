# VBAExcelTips
Excel is you playtool? 
Tired of munging all your data manually and taking hours to update a well-established and productionized report?
Here you can find some time-savers from my time as a Reporting Analyst


# The code below illustrate various ways to reference Extensibility objects.
'IDE Object Model
'VBE --> VBProject --> VBComponent --> CodeModule
'VBE --> VBProject --> VBComponent --> Designer
'VBE --> VBProject --> VBComponent --> Property
'VBE --> VBProject --> Reference
'VBE --> Window


Dim VBAEditor As VBIDE.VBE
Dim vbProj As VBIDE.VBProject
Dim VBP As VBProject
Dim vbComp As VBIDE.VBComponent
Dim CodeMod As VBIDE.CodeModule

Set VBAEditor = Application.VBE
Set VBP = Application.VBE.VBProjects("Angelina")
Set VBP = VBAEditor.VBProjects("Angelina")
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Set vbProj = VBAEditor.ActiveVBProject
Set vbProj = Application.Workbooks("Book1.xls").VBProject
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Set vbComp = ActiveWorkbook.VBProject.VBComponents("Module1")
Set vbComp = vbProj.VBComponents("Module1")

Set vbComp = ThisWorkbook.VBProject.VBComponents(1)
Set vbComp = ThisWorkbook.VBProject.VBComponents("Module1")
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Set CodeMod = ActiveWorkbook.VBProject.VBComponents("Module1").CodeModule
Set CodeMod = vbComp.CodeModule