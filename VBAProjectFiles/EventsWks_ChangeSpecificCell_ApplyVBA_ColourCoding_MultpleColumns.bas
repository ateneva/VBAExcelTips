Attribute VB_Name = "EventsWks_ChangeSpecificCell"
Option Explicit

Private Sub Worksheet_VBAColorCoding_to_MultipleColumns()

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
