Sub KeepData()

Dim Wks As Worksheet

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva
'toggling sheet and workbook protection on/off with a password
'https://datageeking.wordpress.com/2017/09/01/how-to-quickly-protectunprotect-worksheets-in-your-spreadsheet/
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

If ActiveWorkbook.ProtectStructure = True Then
ActiveWorkbook.Unprotect ("annie")

    For Each Wks In ActiveWorkbook.Worksheets
        If Wks.Visible = False Then Wks.Visible = True

        Wks.Activate
        If ActiveSheet.ProtectContents = True Then ActiveSheet.Unprotect ("annie")

    Next Wks

Else

    ActiveWorkbook.Protect ("annie"), Structure:=True

    For Each Wks In ActiveWorkbook.Worksheets
    If Wks.Visible = True Then Wks.Activate

        ActiveSheet.Protect ("annie"), DrawingObjects:=True, Contents:=True, _
        Scenarios:=True, AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:=True

    Next Wks
End If
End Sub
