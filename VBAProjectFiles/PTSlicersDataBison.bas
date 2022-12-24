Attribute VB_Name = "PTSlicersDataBison"
Sub create_slicer()
Dim i As SlicerCaches
Dim j As Slicers
Dim k As Slicer

Set i = ActiveWorkbook.SlicerCaches
Set j = i.Add(ActiveSheet.PivotTables(1), "Region", "My_Region").Slicers
Set k = j.Add(ActiveSheet, , "My_Region", "Region", 0, 0, 200, 200)

MsgBox "Created Slicer"

End Sub

Sub turn_off_on_slicer_fiekd()
Dim i As SlicerCaches
Dim j As Slicers
Dim k As Slicer

Set i = ActiveWorkbook.SlicerCaches
Set j = i.Add(ActiveSheet.PivotTables(1), "Region", "Region").Slicers
Set k = j.Add(ActiveSheet, , "Region", "Region", 0, 0, 200, 200)
i("Region").SlicerItems("West").Selected = False

'You can also use
k.SlicerCache.SlicerItems("West").Selected = False

'Or
k.SlicerCache.SlicerItems(1).Selected = False
MsgBox "Turned off WEST"
i("Region").SlicerItems("West").Selected = True

'Or
k.SlicerCache.SlicerItems("West").Selected = True
MsgBox "Turned on WEST"

End Sub

Sub change_name_caption_slicer()
Dim i As SlicerCaches
Dim j As Slicers
Dim k As Slicer

Set i = ActiveWorkbook.SlicerCaches
Set j = i.Add(ActiveSheet.PivotTables(1), "Region", "Region").Slicers
Set k = j.Add(ActiveSheet, , "Region", "Region", 0, 0, 200, 200)

k.name = "Slicer_Name"         'Or use j(1).Name = "Slicer_Name"
k.Caption = "My Caption"    'Or use j(1).Caption = "My Caption"
MsgBox "Changed slicer name and caption"

End Sub

Sub delete_slicer()
Dim i As SlicerCaches
Dim j As Slicers
Dim k As Slicer

Set i = ActiveWorkbook.SlicerCaches
Set j = i.Add(ActiveSheet.PivotTables(1), "Region", "My_Region").Slicers
Set k = j.Add(ActiveSheet, , "My_Region", "Region", 0, 0, 200, 200)

MsgBox "Created Slicer"
k.Delete
'OR
'j("My_Region").Delete
MsgBox "Deleted Slicer"
End Sub

Sub change_slicer_look_feel()

Dim i As SlicerCaches
Dim j As Slicers
Dim k As Slicer
Set i = ActiveWorkbook.SlicerCaches
Set j = i.Add(ActiveSheet.PivotTables(1), "Region", "My_Region").Slicers
Set k = j.Add(ActiveSheet, , "My_Region", "Region", 0, 0, 200, 200) ' Specify the dimensions here
MsgBox "Created Slicer"

k.Top = 200
k.Left = 200
'OR
'j("My_Region").Top = 200
'j("My_Region").Left = 200

MsgBox "Moved Slicer"
k.Shape.ScaleWidth 0.4, msoFalse, msoScaleFromTopLeft
k.Shape.ScaleHeight 0.6, msoFalse, msoScaleFromTopLeft

MsgBox "Changed Slicers Row Height and Width"
k.RowHeight = 8.4
k.ColumnWidth = 358.4

'OR use the 'j("My_Region") syntax as shown above
MsgBox "Changed Slicer Field Row Height and Width"
k.Style = "SlicerStyleLight3"
MsgBox "Changed Slicer Color"
End Sub

Sub change_slicer_settings()

Dim i As SlicerCaches
Dim j As Slicers
Dim k As Slicer

Set i = ActiveWorkbook.SlicerCaches
Set j = i.Add(ActiveSheet.PivotTables(1), "Region", "Region").Slicers
Set k = j.Add(ActiveSheet, , "Region", "Region", 0, 0, 200, 200)

k.Caption = "Amazing Slicer"

k.DisplayHeader = True
k.SlicerCache.CrossFilterType = xlSlicerNoCrossFilter ' OR xlSlicerCrossFilterShowItemsWithNoData / xlSlicerCrossFilterShowItemsWithDataAtTop
k.SlicerCache.SortItems = xlSlicerSortDescending 'OR xlSlicerSortAscending
k.SlicerCache.SortUsingCustomLists = True
k.SlicerCache.ShowAllItems = False

End Sub

Sub create_duplicate_slicer()

Set i = ActiveWorkbook.SlicerCaches
Set j = i.Add(ActiveSheet.PivotTables(1), "Region", "My_Region1").Slicers
Set k = j.Add(ActiveSheet, , "My_Region1", "Region", 0, 0, 200, 200)
ActiveSheet.Shapes.Range("My_Region1").Duplicate.Select

End Sub


