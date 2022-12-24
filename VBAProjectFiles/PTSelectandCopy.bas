Attribute VB_Name = "PTSelectandCopy"
Option Explicit

Sub SelectAndCopyPT()

Dim PT As PivotTable
Dim PF As PivotField
Dim PI As PivotItem

Set PT = Worksheets("Visits bckgrnd").PivotTables("PivotTable1")
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Worksheets("Visits bckgrnd").Activate
With ActiveSheet

'.DataBodyRange = method in the PivotTable collection
'.DataRange = method in the PivotItem collection and PivotField collections

'*************************************************************partial ranges*********************************************************************************************
PT.RowRange.Copy Worksheets("new").Range("A5")                                                  'selects all the row fields in the pivot table (may be 1 or more than 1)
PT.ColumnRange.Copy Worksheets("new").Range("J2")                                               'selects all the column field in the pivot table (may be 1 or more than 1)
PT.DataBodyRange.Copy Worksheets("new").Range("J5")                                             'select the data for all the pivotfields in the values section

                                                     'select/copy pivotitems
PT.PivotFields("Year").PivotItems("2006").DataRange.Copy Worksheets("new").Range("P5")          'selects the data for this particular item only
PT.PivotFields("Quarter").PivotItems("Quarter 2").DataRange.Copy Worksheets("new").Range("X5") ''selects the data for this particular item only
PT.PivotFields("Purpose").PivotItems("Business").DataRange.Copy Worksheets("new").Range("P30") ''selects the data for this particular item only
PT.PivotFields("Purpose").DataRange.Copy Worksheets("new").Range("P30")                         'selects data for the pivotfield

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~OR a neater appraoch + allows selecting more than 1 item~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PivotField is called Subregion, PivotItem is UKI
PT.PivotSelect "Subregion[UKI]", xlLabelOnly, True                                              'selects only UKI pivotitem and copies the label only
PT.PivotSelect "Subregion[UKI]", xlDataAndLabel, True                                           'select the data and label for the whole pivotitem
PT.PivotSelect "Subregion[UKI]", xlDataOnly, True                                               'selects the data items of the pivotitem only
PT.PivotSelect "Subregion[UKI]", xlFirstRow, True                                               'only selects the first row

PT.PivotSelect "Country[Romania,Croatia]", xlDataAndLabel, True                                 'selects more than 1 pivotitems from a given field
Selection.Group                                                                                 'groups the selected pivotitems

PT.PivotSelect "Subregion", xlDataAndLabel, True 'assuming there are more than 1 pivotfields in the PivotTable and you only want to copy 1

'***********************************************************************whole*******************************************************************************************
PT.PivotSelect "", xlDataAndLabel, True                                                         'select the whole pivottable
Selection.Copy Worksheets("visits").Range("A5")                                                 'and copies it on an exisitng sheet in cell A5

PT.PivotSelect "", xlDataAndLabel                                                               'selects the whole pivot table and copies in another workbook
Selection.Copy TargetWbk.Worksheets("TC Residuals").Range("B147")                               '(must be opened before that)

PT.PivotSelect "", xlDataOnly, True
Selection.Copy
Worksheets("visits").Paste                                                                      'pastes in the cell which is active

PT.PivotSelect "", xlDataAndLabel, True
Selection.Copy
Worksheets.Add.name = "BC"                                                                      'adds a new worksheet with the specified name
Worksheets("BC").Paste                                                                          'and copies on it in cell A1 always

With ActiveSheet                                                                                'copy and paste as values so as not to increase file size too much
    .PivotTables(1).PivotSelect "", xlDataAndLabel
    Selection.Copy
    Output.Range("S13").PasteSpecial xlPasteValuesAndNumberFormats 'Output is a declared workbook
End With

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~finding the last row of a pivot table~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Worksheets ("Visits bckgrnd")

With PT
'copies the last row of the pivot table /assuming there are no empty spaces in the field/
.RowRange.End(xlDown).Select
Range(ActiveCell, Cells((ActiveCell.row), ActiveSheet.UsedRange.Columns.Count)).Copy Worksheets("Summary").Range("B1").End(xlDown).Offset(1, 0)

'copies the 5th visible row in the pivottable
.RowRange.Rows("5:5").Select
Range(ActiveCell, Cells((ActiveCell.row), ActiveSheet.UsedRange.Columns.Count)).Copy Worksheets("Summary").Range("B1").End(xlDown).Offset(1, 0)

'copies the whole column range
.ColumnRange.End(xlDown).Copy
.RowRange.Copy
.DataBodyRange.Copy

'.DataBodyRange = method in the PivotTable collection
'.DataRange = method in the PivotItem collection and PivotField collections

PF.DataRange.End(xlDown).Select 'selects the last row for a pivotfield
PI.DataRange.End(xlDown).Select 'select the last row for a pivoItems (assumes the grand totals for columns is turned on and that the pivot item has been declared

'PT.PivotFields("Supplier").ShowDetail = False
'PT.PivotSelect "Subregion[UKI]", xlLabelOnly, True
''Selection.Offset(0, 1).End(xlDown).Select  'selects the Grand Totals column and goes to the last cell in the same column
''Selection.End(xlDown).Select 'repetition necessary because the subtotals are placed on top (i.e 1st item is blank)

End With

End Sub
