Sub AddDefaultName()

Dim Wks As Worksheet
Dim PT As PivotTable
Dim PF As PivotField
Dim Title As String

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva, September 2016
'https://datageeking.wordpress.com/2017/07/29/quickly-change-a-pivot-value-field-name-with-vba/
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each Wks In ActiveWorkbook.Worksheets
    For Each PT In Wks.PivotTables

        On Error Resume Next
        For Each PF In PT.DataFields

  '------------------Option 1 ------------------------------------------------------
          Title = PF.name
          'comment out the line(s) that you do not need
              PF.name = Mid(Title, 8, Len(Title) - 7) & " "   'removes the "sum of", "max of", "min of"
              PF.name = Mid(Title, 10, Len(Title) - 9) & " "  'removes the "count of"
              PF.name = Mid(Title, 12, Len(Title) - 11) & " " 'removes the "average of", "product of"

  '------------------Option 2------------------------------------------------------
          Title = PF.SourceName & " "
          PF.Caption = Title

        Next PF

    Next PT

Next Wks
End Sub
