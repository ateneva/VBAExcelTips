Attribute VB_Name = "CellInR_ShowCellStats"
Option Explicit

Function IsFormula(cell_ref As Range)

    IsFormula = cell_ref.HasFormula
    
End Function

Function IsBold(Cell) As Boolean
'Returns TRUE if cell is bold, even if from conditional formatting

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

Function CondFormula(myCell As Range, Optional cond As Long = 1) As String
  'Bernie Deitrick programming 2000-02-18, modified D.McR 2001-08-07
  
  Application.Volatile
  CondFormula = ""
  
  On Error Resume Next
  CondFormula = myCell.FormatConditions(cond).Formula1

End Function

Function ConditionalColor(ByVal Cell As Range)

Dim colors As String
Dim i As Long
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'will return the color codes for all the conditional formats a cell contains.
'It will return nothing if there are no conditions,
'and if there is a condition but no color is set for it, then it tells you "none".

'*******************************************************************************
'it will not work for colour scales --> the solution in that case is to paste to word,
'which will convert to normal RGB scales--> macro below
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
For i = 1 To Range(Cell.Address).FormatConditions.Count
    If Range(Cell.Address).FormatConditions(i).Interior.Color <> 0 Then
        colors = colors & "Condition " & i & ": " & _
        Range(Cell.Address).FormatConditions(i).Interior.Color & vbLf
    Else
        colors = colors & "Condition " & i & ": None" & vbLf
    End If
Next

If Len(colors) <> 0 Then
    colors = Left(colors, Len(colors) - 1)
End If

ConditionalColor = colors

End Function

Function GetRGB(ByVal Cell As Range) As String

Dim r As String, G As String
Dim B As String, hexColor As String
hexCode = Hex(Cell.Interior.Color)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    'Note the order excel uses for hex is BGR.
    B = Val("&H" & Mid(hexCode, 1, 2))
    G = Val("&H" & Mid(hexCode, 3, 2))
    r = Val("&H" & Mid(hexCode, 5, 2))
    
    GetRGB = r & ":" & G & ":" & B
    
End Function

Sub ReturnRGB_CF()

    Dim appWord As word.Application
    Dim currentExcel As Worksheet
    Dim StartCell As Range, Cell As Range
    Dim strFormulaRef As String
    
    Dim MyRange As Range
    Dim dest As Range
    
    Set MyRange = Application.InputBox(Prompt:="Please Select a Range", Title:="Choose Range", Type:=8)

    MyRange.Copy
    Set StartCell = Cells(Selection.row, Selection.Column)
    Set appWord = New word.Application

    With appWord
        .Visible = True
        .Documents.Add.Content.Paste
    End With
    With appWord.ActiveWindow.Selection
        .WholeStory
        .Copy
    End With
    
    Set dest = Application.InputBox(Prompt:="Please Select a Range", Title:="Choose a cell to paste", Type:=8)

    ActiveSheet.Paste StartCell.Offset(0, 1)
    strFormulaRef = Replace(StartCell.Offset(0, 1).Address, "$", "")
    With StartCell
        .Offset(0, 2).Formula = "=GetRGB(" & strFormulaRef & ")"
        .Offset(0, 3).Formula = "=GetDEC(" & strFormulaRef & ")"
        .Offset(0, 4).Formula = "=GetHEX(" & strFormulaRef & ")"
    End With
    Range(StartCell.Offset(0, 2), StartCell.Offset(Selection.Rows.Count - 1, 4)).FillDown
    For Each Cell In Range(StartCell.Offset(0, 2), StartCell.Offset(Selection.Rows.Count - 1, 4))
        Cell.Formula = Cell.Formula
    Next
    
    With appWord
        .ActiveWindow.Close wdDoNotSaveChanges
    End With

End Sub

