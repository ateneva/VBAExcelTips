Attribute VB_Name = "CellInR_ColourCells"
Option Explicit

Sub ColourWordsInString()

Dim Cell As Range
Dim prv As String
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each Cell In ActiveSheet.Range("B2:B" & ActiveSheet.UsedRange.Rows.count)

    prv = Cell.Value
    If InStr(prv, "fetish") <> 0 Then Cell.Activate

        With ActiveCell.Characters(Start:=InStr(prv, "fetish"), Length:=6).Font
            .Color = -16776961
        End With
    
Next Cell
End Sub

Sub ColorNegativeValuesInCurrentRange()
Dim WorkRange As Range
Dim Cell As Range
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
If TypeName(Selection) <> "Range" Then Exit Sub
Set WorkRange = Application.Intersect(Selection, ActiveSheet.UsedRange)

For Each Cell In WorkRange
    If Cell.Value < 0 Then
        Cell.Font.Color = RGB(255, 0, 0) 'Makes negative cells red
        Else
        Cell.Font.Color = xlNone
    End If
Next Cell
End Sub

Sub ColourCellsonAbsoluteValues()

Dim Cell As Range
For Each Cell In Selection

    Cell.Activate
    If Abs(Cell.Value) > 1.96 Then Cell.Interior.Color = RGB(0, 255, 204)
    If Cell.Value = "Grand Total" Then Cell.Font.Color = RGB(255, 0, 0)
Next Cell
End Sub

Sub ColourCellsin_DifferentColumn()

Dim Cell As Range
'~~~~~~~~~~~~~~~~~~~~~~~~~~

 For Each Cell In Range("F5:F" & ActiveSheet.UsedRange.Rows.count)
    Cell.Activate
        Range(ActiveCell.Offset(0, -5), ActiveCell).Font.Color = RGB(255, 255, 255)
        If ActiveCell.Value > 10 Then Range(ActiveCell.Offset(0, -5), ActiveCell).Font.Color = RGB(255, 0, 0)
    Next Cell
End If
End Sub

Sub VBAConditonalColorCoding_OffsetColumns()

Dim Cell As Range
Dim Area As Range
Set Area = Worksheets("ConsultantList").Range(Cells(4, 12), Cells(ActiveSheet.UsedRange.Rows.count, 12))

For Each Cell In Area

    Cell.NumberFormat = "mmm"
    If Cell.text = "Apr" Then Cell.Offset(0, -11).Select
    
        Range(ActiveCell, Cells((ActiveCell.row), ActiveSheet.UsedRange.Columns.count)).Select
                                        'Range(ActiveCell, Cells((ActiveCell.Row), 23)).Select - if we know which the end column is
        Selection.Interior.Color = 49407

Next Cell
End Sub

Sub VBAColorCoding_to_OneColumn()

Dim Wks As Worksheet
Dim Col As Integer
Dim Q As Range
Dim Cell As Range

Dim m As String
m = ActiveSheet.Range("G2")
'**************************************

For Each Wks In ThisWorkbook.Worksheets

    If Wks.name = "EMEA" Or Wks.name = "CEE" Or Wks.name = "FRA" Or Wks.name = "GER" Or _
    Wks.name = "GWE" Or Wks.name = "IBE" Or Wks.name = "ITA" Or Wks.name = "MEMA" Or Wks.name = "UKI" Then Wks.Activate

With ActiveSheet

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'edit cond formatting for Import/Export values for QTD (column H)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'**************************************************
If Range("G2").text = "November" Or Range("G2").text = "December" Or Range("G2").text = "January" Then
Col = 13
Set Q = Range("H9")

If Q.Value <= Q.Offset(0, 5).Value * 0.9 Then 'RED Total Utilization
For Each Cell In Range("H19:H23", "H26:H28")

    'Import Local
    If Cell.row = 19 And Cell.Value >= Cell.Offset(0, 5).Value Then Cell.Interior.Color = 65535 'green
    If Cell.row = 19 And Cell.Value < Cell.Offset(0, 5).Value Then Cell.Interior.Color = 255 'red
    
    'Import other country
    If Cell.row = 20 And Cell.Value >= Cell.Offset(0, 5).Value Then Cell.Interior.Color = 65535
    If Cell.row = 20 And Cell.Value < Cell.Offset(0, 5).Value Then Cell.Interior.Color = 5287936 'yellow
    
    'Import IET
    If Cell.row = 21 And Cell.Value >= Cell.Offset(0, 5).Value Then Cell.Interior.Color = 65535
    If Cell.row = 21 And Cell.Value < Cell.Offset(0, 5).Value Then Cell.Interior.Color = 5287936
    
    'Import Other BU
    If Cell.row = 22 And Cell.Value >= Cell.Offset(0, 5).Value Then Cell.Interior.Color = 255
    If Cell.row = 22 And Cell.Value < Cell.Offset(0, 5).Value Then Cell.Interior.Color = 65535
    
    'Import 3P local
    If Cell.row = 23 And Cell.Value >= Cell.Offset(0, 5).Value Then Cell.Interior.Color = 255
    If Cell.row = 23 And Cell.Value < Cell.Offset(0, 5).Value Then Cell.Interior.Color = 5287936
    
    'Export TC local
    If Cell.row = 26 And Cell.Value >= Cell.Offset(0, 5).Value Then Cell.Interior.Color = 65535
    If Cell.row = 26 And Cell.Value < Cell.Offset(0, 5).Value Then Cell.Interior.Color = 255
    
    'Export to other country
    If Cell.row = 27 And Cell.Value >= Cell.Offset(0, 5).Value Then Cell.Interior.Color = 5287936
    If Cell.row = 27 And Cell.Value < Cell.Offset(0, 5).Value Then Cell.Interior.Color = 255
    
    'Export Other BU
    If Cell.row = 28 And Cell.Value >= Cell.Offset(0, 5).Value Then Cell.Interior.Color = 5287936
    If Cell.row = 28 And Cell.Value < Cell.Offset(0, 5).Value Then Cell.Interior.Color = 255

Next Cell
End If
End With
Next Wks

End Sub

Sub VBAColorCoding_to_MultipleColumns()

Dim Cell As Range
Dim Col As Integer
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
If Target.Address = "$EC$1" Then

For Each Cell In ActiveSheet.Range("DQ8:EG11")
    If Cell.Value = "" Then Cell.Value = 0
Next Cell
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'UPDATE Import/Export Conditional Formats on EMEA MTD and EMEA YTD tab ---> i.e. setting is Actual/Target column

For Col = 121 To 139
    Select Case Col
    Case 121, 123, 125, 127, 129, 131, 133, 135, 137, 139
    '*******************************************************
    If Cells(10, Col).Value <= Cells(10, Col + 1).Value * 0.9 Then 'RED Utilization
    
    For Each Cell In Range(Cells(20, Col), Cells(29, Col))
    
        'Import Local
        If Cell.row = 20 And Cell.Value >= Cell.Offset(0, 1).Value Then Cell.Interior.Color = 65535
        If Cell.row = 20 And Cell.Value < Cell.Offset(0, 1).Value Then Cell.Interior.Color = 255
        
        'Import other country
        If Cell.row = 21 And Cell.Value >= Cell.Offset(0, 1).Value Then Cell.Interior.Color = 65535
        If Cell.row = 21 And Cell.Value < Cell.Offset(0, 1).Value Then Cell.Interior.Color = 5287936
        
        'Import IET
        If Cell.row = 22 And Cell.Value >= Cell.Offset(0, 1).Value Then Cell.Interior.Color = 65535
        If Cell.row = 22 And Cell.Value < Cell.Offset(0, 1).Value Then Cell.Interior.Color = 5287936
        
        'Import Other BU
        If Cell.row = 23 And Cell.Value >= Cell.Offset(0, 1).Value Then Cell.Interior.Color = 255
        If Cell.row = 23 And Cell.Value < Cell.Offset(0, 1).Value Then Cell.Interior.Color = 65535
        
        'Import 3P local
        If Cell.row = 24 And Cell.Value >= Cell.Offset(0, 1).Value Then Cell.Interior.Color = 255
        If Cell.row = 24 And Cell.Value < Cell.Offset(0, 1).Value Then Cell.Interior.Color = 5287936
        
        'Export TC local
        If Cell.row = 27 And Cell.Value >= Cell.Offset(0, 1).Value Then Cell.Interior.Color = 65535
        If Cell.row = 27 And Cell.Value < Cell.Offset(0, 1).Value Then Cell.Interior.Color = 255
        
        'Export to other country
        If Cell.row = 28 And Cell.Value >= Cell.Offset(0, 1).Value Then Cell.Interior.Color = 5287936
        If Cell.row = 28 And Cell.Value < Cell.Offset(0, 1).Value Then Cell.Interior.Color = 255
        
        'Export Other BU
        If Cell.row = 29 And Cell.Value >= Cell.Offset(0, 1).Value Then Cell.Interior.Color = 5287936
        If Cell.row = 29 And Cell.Value < Cell.Offset(0, 1).Value Then Cell.Interior.Color = 255
    
    Next Cell
    End Select
Next Col

End If
End Sub

