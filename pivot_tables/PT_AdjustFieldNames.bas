
Sub ChangePFCaptionOfCertainFields()

Dim Wks As Worksheet
Dim PT As PivotTable
Dim PF As PivotField
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva, Oct 2016
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each Wks In ActiveWorkbook.Worksheets

    For Each PT In Wks.PivotTables

        'replace part of the name of a data field
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        For Each PF In PT.DataFields

            'inserts blank between a currency symbol and the text
            If Left(PF.Caption, 1) = Chr(128) Then PF.name = Chr(128) & Chr(32) & Right(PF.name, Len(PF.name) - 1)

            'replace pound with euro
            If PF.Caption Like "*�*" Then PF.name = Chr(128) & Chr(32) & Right(PF.name, Len(PF.name) - 1)

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
