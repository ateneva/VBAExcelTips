
Sub InsertBlankSpacesAfterFirstBlankSpaceUpperCharactersInName()

Dim Wks As Worksheet
Dim PT As PivotTable
Dim PF As PivotField
Dim mStr As String
Dim i As Integer
Dim FindUpper As Integer
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva, Feb 2017; assumes the characters has more than two upper characters
'https://datageeking.wordpress.com/2017/08/06/insert-blank-space-between-upper-characters-in-a-pivot-field-title/
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each Wks In ActiveWorkbook.Worksheets
    For Each PT In Wks.PivotTables

        On Error Resume Next
        For Each PF In PT.DataFields

            MsgBox (PF.Caption & Chr(32) & PF.name & Chr(32) & PF.SourceName)

            'PF.Caption                  'The label text for the pivot field. Read-only String.
            'PF.name                     'Returns or sets the name of the object. Read/write String.
            'PF.SourceName               'Returns the specified object�s name as it appears in the original source data.
                                         'This might be different from the current item name if it has been renamed. Read-only String.

              If PF.Position < 34 Then
                     mStr = PF.Caption

                For i = InStr(mStr, Chr(32)) + 2 To Len(mStr)
                    If Mid(mStr, i, 1) Like "[A-Z]" Then
                        FindUpper = i
                          PF.Caption = Left(mStr, FindUpper - 1) & Chr(32) & Right(mStr, Len(mStr) - FindUpper + 1)
                          Exit For
                    End If
                Next i
              End If
        Next PF

    Next PT

Next Wks
End Sub
