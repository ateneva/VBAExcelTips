
Sub Merg()
Dim Cell As Range

Application.DisplayAlerts = False
With ActiveSheet

    For Each Cell In Range("C1:C12")
      Cell.Activate
          Range(ActiveCell, ActiveCell.Offset(0, 1)).Select
          Selection.MergeCells = True
          'Selection.MergeCells = True
          Selection.VerticalAlignment = xlCenter
    Next Cell

End With
Application.DisplayAlerts = True

End Sub

'------------------------------------------------------------------------------'
Sub UnmergeCells()

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva
'https://datageeking.wordpress.com/2017/05/31/quickly-un-merge-cells-with-vba/
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Dim Cell As Range
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~
For Each Cell In ActiveSheet.Range("A2:AS4100")
    If Cell.MergeArea.Address <> Cell.Address Then Cell.UnMerge

Next Cell

End Sub
