Attribute VB_Name = "DataRangeClear"
Option Explicit

Sub ClearData()

Worksheets("ExportBW").Activate
With ActiveSheet
    Range("A2:CL" & .UsedRange.Rows.Count).Clear 'data + formats
End With

Worksheets("Country_Cl_Sub_Reg").Activate
With ActiveSheet
    Cells.Select
    Selection.ClearContents 'data only
    
    Cells.Clear             'clears everything
    Cells.ClearContents     'clears contents
    Cells.ClearFormats      'clears formats
    Cells.ClearHyperlinks   'clears formats
    Cells.ClearComments     'clears comments
End With

'~~~~~~~~~~~~~~~~~~~~~~deleting avaialable names~~~~~~~~~~~~~~~~~~~~~~~~~~~~
With ActiveWorkbook 'deleting available names

    .Names("BankHolidays").Delete
    .Names("Contract_Type").Delete
    .Names("CurrentWorkRequestStatus").Delete
    .Names("Days").Delete
    .Names("Discretionary_IT_Plans").Delete
    .Names("MoveRequest").Delete
    .Names("PlanRef").Delete
    .Names("RequestType").Delete
    .Names("text_closed").Delete
    .Names("text_launched").Delete
    
End With

End Sub

Sub DeleteHiddenNames()

Dim n As name
Dim Count As Integer
'~~~~~~~~~~~~~~~~~~~~~~~
For Each n In ActiveWorkbook.Names
    If Not n.Visible Then
                n.Delete
                Count = Count + 1
    End If
Next n

MsgBox Count & "hidden names were deleted."
End Sub

