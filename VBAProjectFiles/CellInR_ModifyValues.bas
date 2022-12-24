Attribute VB_Name = "CellInR_ModifyValues"
Option Explicit

Sub ReplaceConstantValues()
Dim Cell As Range

For Each Cell In ActiveSheet.Range("E2:H" & ActiveSheet.UsedRange.Rows.Count).SpecialCells(xlCellTypeVisible)

    Cell.Replace "CEE", "CEE&I"
    Cell.Replace "France", "FRA"
    Cell.Replace "Germany", "GER"
    Cell.Replace "Iberia", "IBE"
    Cell.Replace "Italy", "ITA"
    Cell.Replace "EMEA", "EMEA HQ"

Next Cell

End Sub

Sub DecideConstantValues()
Dim Cell As Range

''***********************using multiple IF statements**********************************************************************

For Each Cell In Range("EN2:EN" & ActiveSheet.UsedRange.Rows.Count)

'~~~~~~~~~~~~~~~~~~~~~~~~hardcoding constants~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    If Cell.Value = "Regular" Then Cell.Value = "R"
    If Cell.Value = "Temporary" Then Cell.Value = "T"
    If Cell.Value = "LCL" Then Cell.Value = "3P"
    
    If Cell.Value = "Russian Federation" Then Cell.Value = "Russia"
    If Cell.Value = "Netherland" Then Cell.Value = "Netherlands"
    If Cell.Value = "United Arab Emirates" Then Cell.Value = "UAE"
    If Cell.Value = "UNITED ARAB EMIRATES" Then Cell.Value = "UAE"
    If Cell.Value = "United Kingdom" Then Cell.Value = "Great Britain"
    
    '~~~~~~~~~~~~~~~~~~using more than one condition~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    If Cell.Value = "NWC" And Cell.Offset(0, -2).Value = "G400" Then Cell.Value = "DCC"
    If Cell.Value = "SIS" And Cell.Offset(0, -2).Value = "G400" Then Cell.Value = "DCC"
    If Cell.Value = "DCC" And Cell.Offset(0, -2).Value = "1Z00" Then Cell.Value = "NWC"
    
    If Cell.Value = "United Kingdom" Then Cell.Offset(0, 1).Value = "Great Britain"
    If Cell.Value = "Great Britain" Then Cell.Offset(0, 1).Value = "Great Britain"
    If Cell.Value = "Netherland" Then Cell.Value = "Netherlands"
    If Cell.Value = "Russian Federation" Then Cell.Value = "Russia"
    If Cell.Value = "Czechia" Then Cell.Value = "Czech Republic"
    
    '~~~~~~~~~~~~~~~~use Like operator and wildcards to avoid searching for exacr match
    If Unit.Value Like "ES*" Then Unit.Value = "ES"
    If Unit.Value Like "Enterprise*" Then Unit.Value = "ES"
    
    '~~~~~~~~~~~~~~~~use the same cell value as a basis for the new value~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    If Right(ActiveCell, 2) = "MM" Then Cell.Value = Left(ActiveCell, Len(ActiveCell) - 2) & "000 000"
    If Right(ActiveCell, 3) = "MMM" Then Cell.Value = Left(ActiveCell, Len(ActiveCell) - 3) & "000 000 000"
    
    Cell.Formula = Left(Cell, Len(Cell) - 6)
    Cell.Formula = Left(Cell, 4)
    Cell.Formula = WorksheetFunction.Trim(Cell)
    Cell.Value = WorksheetFunction.Trim(Cell)

Next Cell

'*******************using SELECT CASE structure -->unlike choose can be ~used with both integers and strings******************

'edit QTD formulae
For Each Cell In Range("F2:F" & .UsedRange.Rows.Count).SpecialCells(xlCellTypeVisible)

If IsNumeric(Cell) Then
    Select Case Q
    Case 1: Cell.Formula = "=SUM(RC[20]:RC[22])"
    Case 2: Cell.Formula = "=SUM(RC[23]:RC[25])"
    Case 3: Cell.Formula = "=SUM(RC[26]:RC[28])"
    Case 4: Cell.Formula = "=SUM(RC[29]:RC[31])"
    End Select
End If
Next Cell

'edit HTD formulae
For Each Cell In Range("G2:G" & .UsedRange.Rows.Count).SpecialCells(xlCellTypeVisible)

If IsNumeric(Cell) Then
    Select Case Q
    Case 1 To 2: Cell.Formula = "=SUM(RC[19]:RC[24])"
    Case 3 To 4: Cell.Formula = "=SUM(RC[25]:RC[30])"
    End Select
End If
Next Cell

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~source Template for FSC Dashaboard~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Range("BK2:BK" & .UsedRange.Rows.Count).ClearContents 'clear previous entries

For Each Cell In Range("BL2:BL" & .UsedRange.Rows.Count)
MRU = Cell.Value

    Select Case MRU
        Case "B039", "B038", "C515", "C513", "D178", "D179", "B016", "C753", "C754", "4D3A", "62UW", "B027"
        Cell.Offset(0, -1).Value = "EMSP"
        
        Case "D011", "E066", "E064"
        Cell.Offset(0, -1).Value = "EMCT"
    End Select
Next Cell

End Sub







