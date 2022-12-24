Attribute VB_Name = "DataRangeFilter"
Option Explicit

Sub FilterDataRange()
Attribute FilterDataRange.VB_ProcData.VB_Invoke_Func = " \n14"

With ActiveSheet

Range("$B$1:$B9000").AutoFilter Field:=1 'turns filters on/off
ActiveSheet.ShowAllData 'clears filters

ThisWorkbook.Worksheets("Flatfile").Rows(1).AutoFilter 'Chase teh tail Wendy's file' turns filters on
ThisWorkbook.Worksheets("Flatfile").Rows(1).AutoFilter Field:=FLD_VERSION_CLOSED, Criteria1:="=" 'text was declared as a Global Constant
                                                'Global Const FLD_VERSION_CLOSED = 26

'number of filtering field = start counting from the column where the first filter is applied
'~~~~~~~~~~~~~~~~~~~~~~~~~~~regular filters~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Range("EY:EY").AutoFilter Field:=155, Criteria1:=">1" 'filtering columns
Rows(1).AutoFilter Field:=34, Criteria1:="FIXED" 'filtering rows
Rows(1).AutoFilter Field:=28, Criteria1:="<>Autocomplete"
Rows(4).AutoFilter Field:=3, Criteria1:="<>" 'filter to exclude blank

Range("R2").AutoFilter Field:=18, Criteria1:="<>renewed"

Rows(5).AutoFilter Field:=77, Criteria1:="D011" 'formula populated column looking at latest headcount data tor retrieve FSC people
Rows(5).AutoFilter Field:=64, Criteria1:="Non Bill (Person Zero Rate)" 'to avoud double counting of D011 employees already appearing in export report

''eliminate duplicates 'not needed always update to the latest FY15 file
'Range("EZ:EZ").AutoFilter Field:=156, Criteria1:=">1"
'Range("EY:EY").AutoFilter Field:=155, Criteria1:="1900-01"  'dmre reliable in removing duplicates than for filtering for last month in report date
'Range("A2:A10000").SpecialCells(xlCellTypeVisible).EntireRow.Delete 'set this way on purpose --> othwerwise deletess the header row

'precise string whole column
Range("AA:AA").AutoFilter Field:=27, Criteria1:="Approved"
Range("Y:Y").AutoFilter Field:=25, Criteria1:="CATIS II"
Range("D:D").AutoFilter Field:=4, Criteria1:="=Core", Operator:=xlOr, Criteria2:="bid"

'precise string (used range only)
Range("BO6:BO" & ActiveSheet.UsedRange.Rows.Count).AutoFilter Field:=67, Criteria1:="Result", Operator:=xlOr, Criteria2:="Overall Result"

Range("$C$1:$CS$5000").AutoFilter Field:=1, Criteria1:="*HC*" 'contains string; single criteria
Range("$B$1:$B9000").AutoFilter Field:=1, Criteria1:="=*ucp.*", Operator:=xlOr, Criteria2:="=*forum*" ''contains string; double criteria
Rows(1).AutoFilter Field:=2, Criteria1:="NS*", Operator:=xlOr, Criteria2:="=SC*" 'begins with string

Range("$A$:$EZ$").AutoFilter Field:=1, Criteria1:="<06/01/2015" 'less than a particular date
Range("$A$:$EZ$").AutoFilter Field:=1, Criteria1:=xlFilterLastMonth, Operator:=xlFilterDynamic  'last month (dynamic) syntax

Rows(1).AutoFilter Field:=27, Criteria1:=xlFilterLastQuarter, Operator:=xlFilterDynamic 'lastquarter

'Yotta manipulation sheet
Range("A:A").AutoFilter Field:=1, Criteria1:=RGB(255, 0, 0), Operator:=xlFilterFontColor 'filter by color
        
'***************************************fitering a variable range of data for more than 2 values**********************************************************
Range("E2:E" & ActiveSheet.UsedRange.Rows.Count).AutoFilter Field:=5, Criteria1:=Array("Approved", "Filled", "To Be Approved"), Operator:=xlFilterValues

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Select Case today

    Case 12 'calendar month December; reported month November = FY period 01
    Range("D:D").AutoFilter Field:=4, Criteria1:=Array("FY2015-01", "FY2015-02", _
    "FY2015-03", "FY2015-04", "FY2015-05", "FY2015-06", "FY2015-07", "FY2015-08", "FY2015-09", "FY2015-10", "FY2015-11", "FY2015-12"), Operator:=xlFilterValues
    
    Case 1 'calendar month January; reported month December = FY period 02
    Range("D:D").AutoFilter Field:=4, Criteria1:=Array("FY2015-02", "FY2015-03", _
    "FY2015-04", "FY2015-05", "FY2015-06", "FY2015-07", "FY2015-08", "FY2015-09", "FY2015-10", "FY2015-11", "FY2015-12"), Operator:=xlFilterValues
    
    Case 2 ''calendar month February; reported month January = FY period 03
    Range("D:D").AutoFilter Field:=4, Criteria1:=Array("FY2015-03", "FY2015-04", _
    "FY2015-05", "FY2015-06", "FY2015-07", "FY2015-08", "FY2015-09", "FY2015-10", "FY2015-11", "FY2015-12"), Operator:=xlFilterValues
    
    Case 3 'calendar month March; reported month February = FY period 04
    Range("D:D").AutoFilter Field:=4, Criteria1:=Array("FY2015-04", "FY2015-05", _
    "FY2015-06", "FY2015-07", "FY2015-08", "FY2015-09", "FY2015-10", "FY2015-11", "FY2015-12"), Operator:=xlFilterValues
    
    Case 4 'calendar month April; reported month March = FY period 05
    Range("D:D").AutoFilter Field:=4, Criteria1:=Array("FY2015-05", "FY2015-06", _
    "FY2015-07", "FY2015-08", "FY2015-09", "FY2015-10", "FY2015-11", "FY2015-12"), Operator:=xlFilterValues
    
    Case 5 'calendar month May; reported month April = FY period 06
    Range("D:D").AutoFilter Field:=4, Criteria1:=Array("FY2015-06", "FY2015-07", _
    "FY2015-08", "FY2015-09", "FY2015-10", "FY2015-11", "FY2015-12"), Operator:=xlFilterValues
    
    Case 6 'calendar month June; reported month May = FY period 07
    Range("D:D").AutoFilter Field:=4, Criteria1:=Array("FY2015-07", "FY2015-08", _
    "FY2015-09", "FY2015-10", "FY2015-11", "FY2015-12"), Operator:=xlFilterValues
    
    Case 7 ''calendar month July; reported month June = FY period 08
    Range("D:D").AutoFilter Field:=4, Criteria1:=Array("FY2015-08", "FY2015-09", _
    "FY2015-10", "FY2015-11", "FY2015-12"), Operator:=xlFilterValues
    
    Case 8 'calendar month August; reported month July = FY period 09
    Range("D:D").AutoFilter Field:=4, Criteria1:=Array("FY2015-09", "FY2015-10", _
    "FY2015-11", "FY2015-12"), Operator:=xlFilterValues
    
    Case 9 'calendar month September; reported month August = FY period 10
    Range("D:D").AutoFilter Field:=4, Criteria1:=Array("FY2015-10", "FY2015-11", _
    "FY2015-12"), Operator:=xlFilterValues
    
    Case 10 'calendar month October; reported month September = FY period 11
    Range("D:D").AutoFilter Field:=4, Criteria1:=Array("FY2015-11", "FY2015-12"), Operator:=xlFilterValues
    
    Case 11 ''calendar month November; reported month October = FY period 12
    Range("D:D").AutoFilter Field:=4, Criteria1:="FY2015-" & Format(Application.WorksheetFunction.EoMonth(Date, 1), "mm"), Operator:=xlFilterValues

End Select
 
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~filtering tables~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ActiveSheet.ListObjects("Table3").Range.AutoFilter Field:=64, Criteria1:="3P"
ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=3, Criteria1:=Array("D011", "E064", "E066"), Operator:=xlFilterValues

   
End With
End Sub
