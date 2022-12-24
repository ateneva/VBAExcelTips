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
Dim PFName As String
'~~~~~~~~~~~~~~~~~~~~~~~~~~

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

Sub ChangeDefaultPFName()

Dim Wks As Worksheet
Dim PT As PivotTable
Dim PF As PivotField
'~~~~~~~~~~~~~~~~~~~~~~~~~~

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
