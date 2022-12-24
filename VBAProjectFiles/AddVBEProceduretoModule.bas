Attribute VB_Name = "AddVBEProceduretoModule"
Option Explicit

Sub AddProcedureToModule()
    Dim vbProj As VBIDE.VBProject
    Dim VBP As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    Dim CodeMod As VBIDE.CodeModule
    Dim LineNum As Long
    Const DQUOTE = """" ' one " character

    Set vbProj = ActiveWorkbook.VBProject 'if adding to Wbk
    Set VBP = Application.VBE.VBProject("Angelina") 'if adding to Personal
    
    Set vbComp = vbProj.VBComponents("Module1")
    Set vbComp = VBP.VBComponents("App") 'if adding to Personal
    Set CodeMod = vbComp.CodeModule
    
    With CodeMod
        LineNum = .CountOfLines + 1
        .InsertLines LineNum, "Public Sub SayHello()"
        
        LineNum = LineNum + 1
        .InsertLines LineNum, "    MsgBox " & DQUOTE & "Hello World" & DQUOTE
        
        LineNum = LineNum + 1
        .InsertLines LineNum, "End Sub"
    End With
End Sub

Sub GetProcedureText()
'the procedure inserts the text of another procedure, specified as string
'the text does not have to be hard-coded here but can simply be a reference to another procedure in another module

Dim CodePan As VBIDE.CodeModule
Dim S As String
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Set CodePan = ThisWorkbook.VBProject.VBComponents("Module2").CodeModule 'if stored in a workbook
Set CodePan = Application.VBE.VBProjects("Angelina").VBComponents("App").CodeModule 'if stored in Personal

S = "Sub ABC()" & vbNewLine & " MsgBox ""Hello World"",vbOkOnly" & vbNewLine & "End Sub" & vbNewLine

With CodePan
    .InsertLines .CountOfLines + 1, S
End With

End Sub
