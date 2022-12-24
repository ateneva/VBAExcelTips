Attribute VB_Name = "CellInR_CleanUpString"
Option Explicit

Sub ReplaceACharInString()

Dim Cell As Range
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

With ActiveSheet
For Each Cell In Range("A2:A" & ActiveSheet.UsedRange.Rows.Count).SpecialCells(xlCellTypeVisible)

Cell.Activate
 'blank character was in the middle
 'VBA built-in would not work
 'however the WS function wouild

ActiveCell.Value = Trim(ActiveCell)
ActiveCell.Value = Application.WorksheetFunction.Trim(ActiveCell)

If InStr(Cell, "n") = 1 Then Cell.Replace "n", ""

Next Cell

End With
End Sub

Sub EncloseinQuotes()

Dim Cell As Range
Dim IP As String
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

With ActiveSheet
For Each Cell In Range("A2:A" & ActiveSheet.UsedRange.Rows.Count)
    IP = Cell.Value
    Cell.Value = "''" & IP & "',"
Next Cell

End With
End Sub

Sub EncloseHeadersinQuotes()

Dim Cell As Range
Dim IP As String
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

With ActiveSheet
For Each Cell In Range("O1:O201")
    IP = Cell.Value
    Cell.Value = "''" & IP & "'"
Next Cell

End With
End Sub

Sub EncloseinPercentageSigns()

Dim Cell As Range
Dim IP As String
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

With ActiveSheet
For Each Cell In Range("Z2:Z105")
    IP = Cell.Value
    Cell.Value = "''%" & IP & "%' "
Next Cell

End With
End Sub

Sub TestCharactersCase()

Dim mStr As String
Dim i As Integer
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
mStr = "This Is A Test"

For i = 1 To Len(mStr)

    Select Case Asc(Mid(mStr, i, 1))
        Case 65 To 90: MsgBox "Upper"
        Case 97 To 122: MsgBox "Lower"
        Case Else: MsgBox "Not an alpha character"
    End Select
Next i
End Sub

Sub InsertSpacesBetweenCharactersCase()

Dim mStr As String
Dim i As Integer
Dim Cell As Range
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
For Each Cell In ActiveSheet.Range("G3:G43")
    mStr = Cell.Value
    
    For i = 1 To Len(mStr)
    
        Select Case Asc(Mid(mStr, i, 1))
            Case 65 To 90: Cell.Value = Left(mStr, i - 1) & Chr(32)
        End Select
    Next i
    
Next Cell
End Sub

Sub FindFirstUpperCharacter()

Dim mStr As String
Dim FindUpper As String
Dim FindLower As String
Dim i As Integer
Dim Cell As Range
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'assumes the string has only two Upper characters that need separating by blank space
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
For Each Cell In ActiveSheet.Range("G3:G43")
    mStr = Cell.Value
    
    For i = 2 To Len(mStr)
        If Mid(mStr, i, 1) Like "[A-Z]" Then
            FindUpper = i
            Cell.Value = Left(mStr, FindUpper - 1) & Chr(32) & Right(mStr, Len(mStr) - FindUpper + 1)
            Exit For
        End If
    Next i
Next Cell
End Sub

Sub FindFirstUpperCharacterAfterABlankSpace()

Dim mStr As String
Dim FindUpper As String
Dim FindLower As String
Dim i As Integer
Dim Cell As Range
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'assumes the string has only two Upper characters that need separating by blank space
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
For Each Cell In ActiveSheet.Range("G3:G43")
    mStr = Cell.Value
    Cell.Offset(0, 1).Value = InStr(mStr, Chr(32))     'identifies the position of the first blank space
    Cell.Offset(0, 2).Value = InStr(mStr, Chr(32)) + 1 'adds 1
    
    For i = InStr(mStr, Chr(32)) + 2 To Len(mStr)      'loops between the second character after the blank space and the remaining part of the string
        If Mid(mStr, i, 1) Like "[A-Z]" Then
            FindUpper = i
            Cell.Value = Left(mStr, FindUpper - 1) & Chr(32) & Right(mStr, Len(mStr) - FindUpper + 1)
            Exit For
        End If
    Next i
Next Cell
End Sub


