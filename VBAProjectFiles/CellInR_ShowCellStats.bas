Attribute VB_Name = "CellInR_ShowCellStats"
Option Explicit
Function IsFormula(cell_ref As Range)

    IsFormula = cell_ref.HasFormula
    
End Function

Function IsBold(Cell) As Boolean 'Returns TRUE if cell is bold, even if from conditional formatting

IsBold = Cell.Range("A1").DisplayFormat.Font.Bold
'will return an error if some characters are bold and the others not
End Function
Function IsItalic(Cell) As Boolean 'Returns TRUE if cell is italic

IsItalic = Cell.Range("A1").Font.Italic

End Function

Function IsLike(text As String, pattern As String) As Boolean
'Returns true if the first argument is like the second

IsLike = text Like pattern
End Function

Sub CopyComment()

Dim Cell As Range
Dim cmt As Comment
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
With ActiveSheet

    For Each Cell In Range(Cells(2, 6), Cells(2, 14))
        Cell.Comment.Delete
        Cell.AddComment
        Cell.Comment.text text:="Formula developed by Angelina Teneva"
    Next Cell
    
    '~~~~~~~~~~~~shows comments in a MsgBox~~~~~~~~~~~~~
    For Each cmt In ActiveSheet.Comments
        MsgBox cmt.text
    Next cmt
    '~~~~~~~~~~~~prints the comments in the Immediate window
    For Each cmt In ActiveSheet.Comments
        Debug.Print cmt.text
    Next cmt

'changes the colour of comments
.Comments(1).Shape.Fill.ForeColor.RGB = RGB(0, 255, 0)

End With
End Sub

Function CondFormula(myCell As Range, Optional cond As Long = 1) As String
  'Bernie Deitrick programming 2000-02-18, modified D.McR 2001-08-07
  
  Application.Volatile
  CondFormula = ""
  On Error Resume Next
  CondFormula = myCell.FormatConditions(cond).Formula1

End Function

Sub Show_Formulas_in_comments()

  Dim Cell As Range
  On Error Resume Next
  Selection.ClearComments
  On Error GoTo 0
  
  For Each Cell In Intersect(Selection, ActiveSheet.UsedRange)
     If Cell.Formula <> "" Then
        Cell.AddComment
        Cell.Comment.Visible = False
        On Error Resume Next  'fails on invalid formula
        Cell.Comment.text text:=Cell.Address(0, 0) & _
           "  value:    " & Cell.Value & Chr(10) & _
           "  format:   " & Cell.NumberFormat & Chr(10) & _
           "  Formula:  " & Cell.Formula
        On Error GoTo 0
     End If
  Next Cell

End Sub

