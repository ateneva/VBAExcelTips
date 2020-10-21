Sub AddDataField()

Dim Wks As Worksheet
Dim PC As PivotCache
Dim PT As PivotTable
Dim PF As PivotField
Dim i As Integer
Dim FieldPosition As Integer
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva, Oct 2016
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each Wks In ActiveWorkbook.Worksheets
    Wks.Activate
    i = ActiveSheet.Index       'not having the Worksheet Index condition causes the field to be added multiple times to the same sheet
    If i >= 5 And i <= 12 Then

        With ActiveSheet

        For Each PT In ActiveSheet.PivotTables
            FieldPosition = PivotFields("Paid Coverage").Position - 1

            For Each PF In PT.PivotFields
                If PF.Name = "Revenue" Then                         'add a field
                    PF.Orientation = xlDataField
                    PF.Function = xlSum
                    PF.Calculation = xlRunningTotal
                    PF.Position = FieldPosition
                    PF.NumberFormat = "#,###"
                    PF.Caption = "Revenue YTD"
                End If

                If PF.Name Like "kw*" Then PF.Orientation = xlHidden    'remove a field from view
            Next PF
        Next PT
        End With
    End If
Next Wks
End Sub
