Attribute VB_Name = "PTNames"
Option Explicit

Sub NamePTs()

Dim i As Integer
For i = 1 To ActiveWorkbook.Worksheets.Count

Worksheets(i).Activate
With ActiveSheet
   For i = 1 To 3
   Select Case i
      Case 1: .PivotTables(i).name = "First"
      Case 2: .PivotTables(i).name = "Second"
      Case 3: .PivotTables(i).name = "Third"
   End Select
   Next i
End With

Next i
End Sub

Sub ShowPFName()

Dim Wks As Worksheet
Dim PT As PivotTable
Dim PF As PivotField
Dim PFname As String
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each Wks In ActiveWorkbook.Worksheets

    For Each PT In Wks.PivotTables
        
        For Each PF In PT.DataFields     'only makes sense for DataFields as their names typically get changed
        
            MsgBox (PF.Caption & Chr(32) & PF.name & Chr(32) & PF.SourceName)
        
'            PF.Caption                  'The label text for the pivot field. Read-only String.
'            PF.name                     'Returns or sets the name of the object. Read/write String.
'            PF.SourceName               'Returns the specified object’s name as it appears in the original source data.
'                                        'This might be different from the current item name if it has been renamed. Read-only String.
        Next PF
    Next PT
        
Next Wks
End Sub

Sub AddDefaultName()

Dim Wks As Worksheet
Dim PT As PivotTable
Dim PF As PivotField
Dim Title As String
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva, September 2016
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each Wks In ActiveWorkbook.Worksheets
    For Each PT In Wks.PivotTables
        
        On Error Resume Next
        For Each PF In PT.DataFields
                Title = PF.SourceName & " "
                PF.Caption = Title

        Next PF
    
    Next PT
        
Next Wks
End Sub

Sub InsertBlankSpacesBetweenUpperCharactersInName()

Dim Wks As Worksheet
Dim PT As PivotTable
Dim PF As PivotField
Dim mStr As String
Dim i As Integer
Dim FindUpper As Integer
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva, Feb 2017; assumes the characters has only two upper characters
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each Wks In ActiveWorkbook.Worksheets
    For Each PT In Wks.PivotTables
        
        On Error Resume Next
        For Each PF In PT.DataFields
               mStr = PF.Caption
                
                    For i = 2 To Len(mStr)
                        If Mid(mStr, i, 1) Like "[A-Z]" Then
                            FindUpper = i
                            PF.Caption = Left(mStr, FindUpper - 1) & Chr(32) & _
                                            Right(mStr, Len(mStr) - FindUpper + 1)
                            Exit For
                        End If
                    Next i
        Next PF
    
    Next PT
        
Next Wks
End Sub

Sub InsertBlankSpacesAfterFirstBlankSpaceUpperCharactersInName()

Dim Wks As Worksheet
Dim PT As PivotTable
Dim PF As PivotField
Dim mStr As String
Dim i As Integer
Dim FindUpper As Integer
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva, Feb 2017; assumes the characters has more than two upper characters
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each Wks In ActiveWorkbook.Worksheets
    For Each PT In Wks.PivotTables
        
        On Error Resume Next
        For Each PF In PT.DataFields
            If PF.Position < 34 Then
                   mStr = PF.Caption
                    
                        For i = InStr(mStr, Chr(32)) + 2 To Len(mStr)  'loops between the second character after the blank space and the remaining part of the string
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

Sub ChangeDefaultPFName()

Dim Wks As Worksheet
Dim PT As PivotTable
Dim PF As PivotField
Dim Title As String
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each Wks In ActiveWorkbook.Worksheets
    For Each PT In Wks.PivotTables
            
        On Error Resume Next
        For Each PF In PT.DataFields
            
            Title = PF.name
            
            PF.name = Mid(Title, 8, Len(Title) - 7) & " "   'removes the "sum of", "max of", "min of"
            PF.name = Mid(Title, 10, Len(Title) - 9) & " "  'removes the "count of"
            PF.name = Mid(Title, 12, Len(Title) - 11) & " " 'removes the "average of", "product of"
        
        Next PF
    Next PT
Next Wks
End Sub

Sub ChangePFCaptionOfCertainFields()

Dim Wks As Worksheet
Dim PT As PivotTable
Dim PF As PivotField
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva, Oct 2016
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each Wks In ActiveWorkbook.Worksheets

    For Each PT In Wks.PivotTables
    
        'replace part of the name of a data field
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        For Each PF In PT.DataFields
            
            'inserts blank between a currency symbol and the text
            If Left(PF.Caption, 1) = Chr(128) Then PF.name = Chr(128) & Chr(32) & Right(PF.name, Len(PF.name) - 1)
            
            'replace pound with euro
            If PF.Caption Like "*£*" Then PF.name = Chr(128) & Chr(32) & Right(PF.name, Len(PF.name) - 1)
                                   
        Next PF
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        'completely change the name of a data field
        For Each PF In PT.DataFields
            
            If PF.Caption Like "*USD*" Then PF.name = "AUD"
        
        Next PF
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        'change the name of a column, row or page field
        
        For Each PF In PT.PivotFields
            If PF.Orientation <> xlHidden And PF.Orientation <> xlDataField Then
                If PF.Caption Like "*Country*" Then PF.name = "User Country"
    
            End If
        Next PF
    
    Next PT
    
Next Wks
End Sub




