Attribute VB_Name = "HandyStuff"

Sub Start_Test_Time()
Attribute Start_Test_Time.VB_ProcData.VB_Invoke_Func = "T\n14"

ActiveCell.Value = "=I11-TIME(HOUR(NOW()),MINUTE(NOW()),SECOND(NOW()))"
ActiveCell.Copy
ActiveCell.PasteSpecial xlPasteValuesAndNumberFormats
Application.CutCopyMode = False

End Sub

Sub GetRid()
Attribute GetRid.VB_ProcData.VB_Invoke_Func = "C\n14"

Application.CutCopyMode = False

End Sub

Sub PasteSp()
Attribute PasteSp.VB_ProcData.VB_Invoke_Func = "V\n14"

Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
       
End Sub

Sub TransposeasValues()

    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True

End Sub

Sub UnhideAll()

Dim Wks As Worksheet

For Each Wks In ActiveWorkbook.Worksheets
If Wks.Visible = False Then Wks.Visible = xlSheetVisible
If Wks.Visible = xlSheetVeryHidden Then Wks.Visible = xlSheetVisible
Next Wks

End Sub

Sub d()
Attribute d.VB_ProcData.VB_Invoke_Func = "D\n14"
Selection.Value = Date

End Sub

Sub Cal()

Calendar.Show False

End Sub

Sub DeleteRow()

Dim Cell As Range
Dim Sheet As Worksheet
Set Sheet = ActiveWorkbook.Worksheets("Sheet2")

Worksheets("India Debts").Activate

Range("A9").Formula = "=IFERROR(VLOOKUP(B9,Sheet2!A:A,1,FALSE),0)"
Range("A9:A" & ActiveSheet.UsedRange.Rows.Count).FillDown
Range("A10:A" & ActiveSheet.UsedRange.Rows.Count).Copy
Range("A10").PasteSpecial xlPasteValues
Application.CutCopyMode = False

For Each Cell In ActiveWorkbook.Worksheets("India Debts").Range("B9:B28")
If Cell.Value = Cell.Offset(0, -1).Value Then Cell.EntireRow.Delete
Next Cell

End Sub

Sub CutPaste()

Dim i As Long
Dim Master As Worksheet
Dim Released As Worksheet

Set Master = ActiveWorkbook.Worksheets("Master Tracker")
Set Released = ActiveWorkbook.Worksheets("Released")

Master.Activate
With ActiveSheet

For i = 10 To Cells(Rows.Count, "K").End(xlUp).row
If Cells(i, "K").Value = "YES" Then Cells(i, "K").EntireRow.Cut Released.Range("A1").Rows.End(xlDown).Offset(1, 0)
Next i

End With
End Sub





