
Sub TweakAllFiles()

Dim Cell As Range
Dim path As String
Dim file As String
Dim fullfilepath As String

'-------------------------------------------------------------------------------
'written by Angelina Teneva, March 2017
'https://datageeking.wordpress.com/2018/02/24/tweak-all-files/
'-------------------------------------------------------------------------------


Application.DisplayAlerts = False
For Each Cell In ThisWorkbook.Worksheets("UCAS").Range("A2:A95")

    file = Cell.Value
    path = Cell.Offset(0, 6).Value
    fullfilepath = path & file

    If file <> "EOC_data_resource_2014-DR2_001_01.csv" Then

        Workbooks.Open FileName:=fullfilepath, ReadOnly:=False, UpdateLinks:=False

        With ActiveWorkbook
            With ActiveSheet
                Rows("1:5").EntireRow.Delete
                Columns("B:C").Replace "'", ""
            End With
            .Save
            .Close
        End With
    End If

Next Cell
Application.DisplayAlerts = True
End Sub
