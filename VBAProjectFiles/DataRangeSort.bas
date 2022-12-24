Attribute VB_Name = "DataRangeSort"
Option Explicit

Sub SortData()
Attribute SortData.VB_ProcData.VB_Invoke_Func = " \n14"

'sorting multiple columns by using the regular Soet Ascending/Descending button
ActiveSheet.Range("A:FB").Sort Key1:=Range("A1"), Order1:=xlDescending, HEADER:=xlYes 'sort with newest dates on top so tha tthe Uniques? formula can work properly
ActiveSheet.Range("A:AZ").Sort Key1:=Range("D1"), Order1:=xlAscending, HEADER:=xlYes

'sort using the DataSort (multiple options) button
SDMInput.AutoFilter.Sort.SortFields.Add Key:=Range("X1"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal

ActiveSheet.Range("A1").Sort Key1:=Range("D6"), Order1:=xlAscending, _
        HEADER:=xlYes, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom

''''~~~~~~~~~~~~~~~~~~~~~~sorting on more than 1 criteria
With ActiveSheet
'~~~~~data needs to be sorted so that C6 PRJ Resp Cost Centers can be correctly attributed to a PRJ Country~~~~~~~~~~~~~~~~~~~~
.Sort.SortFields.Add(Range("A2:A" & .UsedRange.Rows.Count), xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color = RGB(204, 255, 204)
.Sort.SortFields.Add Key:=Range("D2:D" & .UsedRange.Rows.Count), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

.Sort.SetRange Columns("A:F")
.Sort.HEADER = xlYes
.Sort.MatchCase = False
.Sort.Orientation = xlTopToBottom
.Sort.SortMethod = xlPinYin
.Sort.Apply
'~~~~~~~~~~~~~~~~~~~~~~sorting is applied in the Export file--> copying must be the last thing to do~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
End With

End Sub

