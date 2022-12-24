Attribute VB_Name = "AddVBEModuleReferences"
Option Explicit

Sub MyCodes()

Dim VBAEditor As VBIDE.VBE
Set VBAEditor = Application.VBE
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Dim VBP As VBProject
Set VBP = VBAEditor.VBProjects("Angelina") 'accesses your Personal Macro file
'or
'Set VBP = Application.VBE.VBProjects("Angelina")

Set VBP = VBAEditor.ActiveVBProject
MsgBox (VBP.HelpContextID)
    
End Sub

Sub ListAllProceduresMsgBox()
'written by John WalkenBach
'the code will list in a MsgBox all procedures found in a module

Dim VBP As VBIDE.VBProject
Dim VBC As VBComponent
Dim CM As CodeModule
Dim StartLine As Long
Dim Msg As String
Dim ProcName As String
'~~~~~~~the code lists all VBA Procedures in a module
'Uses the the PERSONAL macro file
Set VBP = Application.VBE.VBProjects("Angelina")

'Loop through the VB components
For Each VBC In VBP.VBComponents
    Set CM = VBC.CodeModule
    Msg = Msg & vbNewLine
    StartLine = CM.CountOfDeclarationLines + 1
    
    Do Until StartLine >= CM.CountOfLines
        Msg = Msg & VBC.name & ": " & CM.ProcOfLine(StartLine, vbext_pk_Proc) & vbNewLine
        StartLine = StartLine + CM.ProcCountLines(CM.ProcOfLine(StartLine, vbext_pk_Proc), vbext_pk_Proc)
    Loop
Next VBC

ActiveCell.Value = Msg
End Sub
Sub ListReferences()
Dim Ref As Reference
'~~~~~~~~~~~~~~~~~~~~~displays references for active Project
Msg = ""
For Each Ref In ActiveWorkbook.VBProject.References
    Msg = Msg & Ref.name & vbNewLine
    Msg = Msg & Ref.Description & vbNewLine
    Msg = Msg & Ref.FullPath & vbNewLine & vbNewLine
Next Ref

MsgBox Msg
End Sub



