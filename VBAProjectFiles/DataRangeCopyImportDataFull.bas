Attribute VB_Name = "DataRangeCopyImportDataFull"
Option Explicit

Sub Update_Workbook()

Application.Run "GetExport"
Application.Run "GetSubregion"
Application.Run "GetPolaris"
Application.Run "GetHCdata"

End Sub

Sub GetExport()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'code written by Angelina Teneva, August 2012
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Dim OutputBk As Workbook
Set OutputBk = ThisWorkbook

Dim exportpath As String
Dim exportfscpath As String

Dim Cell As Range

exportpath = ThisWorkbook.Worksheets("Control File Locations").Range("A4").Value
exportfscpath = ThisWorkbook.Worksheets("Control File Locations").Range("A7").Value

'****************************************************************************

'get export hours data
Workbooks.Open FileName:=exportpath, ReadOnly:=True, UpdateLinks:=False
With ActiveWorkbook
    Worksheets("Export Data").Activate
    
    With ActiveSheet
        Range(Cells(7, 1), Cells(ActiveSheet.UsedRange.Rows.Count, 71)).Copy OutputBk.Worksheets("ExportBW").Range("A2")
    End With
        ActiveWorkbook.Close savechanges:=False
End With

'get export FSC data
Workbooks.Open FileName:=exportfscpath, ReadOnly:=True, UpdateLinks:=False
ActiveWorkbook.Worksheets("Re-arrange").Activate
With ActiveSheet
    Range(Cells(5, 1), Cells(ActiveSheet.UsedRange.Rows.Count, 71)).Copy
End With

OutputBk.Worksheets("ExportBW").Activate
With ActiveSheet
    Range("A2").Rows.End(xlDown).Offset(1, 0).Select
.Paste
End With

'get CATIS data
Windows("Export FSC Template.xlsm").Activate
ActiveWorkbook.Worksheets("CATIS data").Activate
With ActiveSheet
    Range(Cells(5, 1), Cells(ActiveSheet.UsedRange.Rows.Count, 71)).Copy
End With

OutputBk.Worksheets("ExportBW").Activate
With ActiveSheet
    Range("A2").Rows.End(xlDown).Offset(1, 0).PasteSpecial xlPasteValuesAndNumberFormats
End With

Windows("Export FSC Template.xlsm").Activate
Application.CutCopyMode = False
ActiveWorkbook.Close savechanges:=False

OutputBk.Activate
OutputBk.Save

'get rid of C6 values
For Each Cell In ActiveSheet.Range("E2:E" & ActiveSheet.UsedRange.Rows.Count)
    If Cell.Text = "C6" Then Cell.Value = Mid(Cell.Offset(0, 18), 3, 2)
Next Cell

OutputBk.Save

End Sub

Sub GetSubregion()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'code written by Angelina Teneva, August 2012
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Dim OutputBk As Workbook
Set OutputBk = ThisWorkbook

Dim subregionpath As String
Dim rmrpath As String

subregionpath = ThisWorkbook.Worksheets("Control File Locations").Range("A10").Value
rmrpath = ThisWorkbook.Worksheets("Control File Locations").Range("A13").Value

'*****************************************************************************
'get Subregion Data
Workbooks.Open FileName:=subregionpath, ReadOnly:=True, UpdateLinks:=False
With ActiveWorkbook
    Worksheets("existing CLs").Cells.Copy
End With

OutputBk.Worksheets("Country_CL_Sub_Reg").Activate
With ActiveSheet
    Range("A1").Select
    Selection.PasteSpecial xlPasteValues
End With

'get RMR Data
Workbooks.Open FileName:=rmrpath, ReadOnly:=True, UpdateLinks:=False

With ActiveWorkbook
Worksheets("Default").Activate

With ActiveSheet
    Range("A15:A" & .UsedRange.Rows.Count).Copy OutputBk.Worksheets("RMRincludingMRUdescr").Range("A2")
    Range("AA15:AA" & .UsedRange.Rows.Count).Copy OutputBk.Worksheets("RMRincludingMRUdescr").Range("B2")
    Range("AB15:AC" & .UsedRange.Rows.Count).Copy OutputBk.Worksheets("RMRincludingMRUdescr").Range("C2")
    Range("E15:E" & .UsedRange.Rows.Count).Copy OutputBk.Worksheets("RMRincludingMRUdescr").Range("F2")
    Range("AP15:AP" & .UsedRange.Rows.Count).Copy OutputBk.Worksheets("RMRincludingMRUdescr").Range("G2")
    Range("AF15:AF" & .UsedRange.Rows.Count).Copy OutputBk.Worksheets("RMRincludingMRUdescr").Range("H2")
    Range("C15:D" & .UsedRange.Rows.Count).Copy OutputBk.Worksheets("RMRincludingMRUdescr").Range("I2")
    Range("Z15:Z" & .UsedRange.Rows.Count).Copy OutputBk.Worksheets("RMRincludingMRUdescr").Range("K2")
    Range("S15:S" & .UsedRange.Rows.Count).Copy OutputBk.Worksheets("RMRincludingMRUdescr").Range("L2")
End With
End With

ActiveWorkbook.Close savechanges:=False
Windows("CL_Subregion_Country_mapping.xlsx").Activate
ActiveWorkbook.Close savechanges:=False

OutputBk.Worksheets("RMRincludingMRUdescr").Activate
OutputBk.Save

End Sub

Sub GetPolaris()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'code written by Angelina Teneva, August 2012
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Dim OutputBk As Workbook
Set OutputBk = ThisWorkbook

Dim polarispath As String
Dim mrupath As String

polarispath = ThisWorkbook.Worksheets("Control File Locations").Range("A16").Value
mrupath = ThisWorkbook.Worksheets("Control File Locations").Range("A19").Value

'*****************************************************************************
'clear previous data
Worksheets("Polaris").Activate
With ActiveSheet
    Range("A59:C" & .UsedRange.Rows.Count).Clear
End With

'get POLARIS data
Workbooks.Open FileName:=polarispath, ReadOnly:=True, UpdateLinks:=False
With ActiveWorkbook
Worksheets("Sheet1").Activate
    With ActiveSheet
        Range("B2:B" & .UsedRange.Rows.Count).Copy OutputBk.Worksheets("Polaris").Range("A59")
        Range("F2:F" & .UsedRange.Rows.Count).Copy OutputBk.Worksheets("Polaris").Range("C59")
        Range("I2:J" & .UsedRange.Rows.Count).Copy OutputBk.Worksheets("Polaris").Range("D59")
    End With
End With
ActiveWorkbook.Close savechanges:=False

OutputBk.Worksheets("Polaris").Activate
With ActiveSheet
    Range("B58").Formula = "=LEFT(A58,10)"
    Range("B58:B" & .UsedRange.Rows.Count).FillDown
    Rows("58:58").EntireRow.Delete

    Range("F58:F" & .UsedRange.Rows.Count).FillDown 'proft center description trimmed
End With

'clear previous MRU codes data
Worksheets("MRU code list").Activate
With ActiveSheet
    Columns("A:B").Clear
End With

Workbooks.Open FileName:=mrupath, ReadOnly:=True, UpdateLinks:=False

With ActiveWorkbook
    Worksheets("MRU Hierarchy").Activate
        With ActiveSheet
            Columns("D:D").Copy OutputBk.Worksheets("MRU code list").Range("A1")
            Columns("I:I").Copy OutputBk.Worksheets("MRU code list").Range("B1")
        End With
End With
ActiveWorkbook.Close savechanges:=False

OutputBk.Worksheets("MRU code list").Activate
OutputBk.Save

End Sub

Sub GetHCdata()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'code written by Angelina Teneva, August 2012
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Dim OutputBk As Workbook
Set OutputBk = ThisWorkbook

Dim hcdata As String
Dim hcrmrdata As String

hcdata = ThisWorkbook.Worksheets("Control File Locations").Range("A22").Value
hcrmrdata = ThisWorkbook.Worksheets("Control File Locations").Range("A25").Value

'*****************************************************************************
Workbooks.Open FileName:=hcdata, ReadOnly:=True, UpdateLinks:=False
With ActiveWorkbook

Worksheets("data").Activate
With ActiveSheet
    Range(Cells(5, 2), Cells(ActiveSheet.UsedRange.Rows.Count, 32)).Copy OutputBk.Worksheets("HC Report").Range("B4")
End With

End With
ActiveWorkbook.Close savechanges:=False

'************************************************************************************
'get HC RMR data
Workbooks.Open FileName:=hcrmrdata, ReadOnly:=True, UpdateLinks:=False
With ActiveWorkbook

'checks for hidden columns on HC RMR data file and if any are found, it unhides them
If Columns("C:D").EntireColumn.Hidden = True Then Columns("C:D").EntireColumn.Hidden = False
If Columns("F:K").EntireColumn.Hidden = True Then Columns("F:K").EntireColumn.Hidden = False
If Columns("M:AM").EntireColumn.Hidden = True Then Columns("M:AM").EntireColumn.Hidden = False

'filters for HP employees only
Rows(18).AutoFilter Field:=3, Criteria1:="HP"
Range(Cells(19, 1), Cells(ActiveSheet.UsedRange.Rows.Count, 42)).SpecialCells(xlCellTypeVisible).Copy OutputBk.Worksheets("TSC HP").Range("A2")
Rows(18).AutoFilter

'reset fitlers and copy new data
Rows(18).AutoFilter Field:=3, Criteria1:="nonHP"
Range(Cells(19, 1), Cells(ActiveSheet.UsedRange.Rows.Count, 42)).SpecialCells(xlCellTypeVisible).Copy OutputBk.Worksheets("TS Contractors All").Range("A3")

End With
ActiveWorkbook.Close savechanges:=False

'****************************************************************************************
'get employees ids and put on check tab
OutputBk.Worksheets("ExportBW").Activate
With ActiveSheet
    Range("BC2:BC" & .UsedRange.Rows.Count).Copy Worksheets("check").Range("A6")
End With

OutputBk.Worksheets("check").Activate
With ActiveSheet
    Range("A6:A" & .UsedRange.Rows.Count).RemoveDuplicates Columns:=1, Header:=xlNo
End With

OutputBk.Save

End Sub
