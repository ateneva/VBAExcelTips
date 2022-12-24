Attribute VB_Name = "MsgBoxShow"
Option Explicit

Sub ShowExcelVersion()

MsgBox Application.VERSION
MsgBox "Welcome to Microsoft Excel version " & Application.VERSION & " running on " & Application.OperatingSystem & "!"

'Excel 2007 (12.0)
'Excel 2010 (14.0)
'Excel 2013 (15.0)
'Excel 2016 (16.0)

End Sub


Sub Show_Colour_Code()

    MsgBox (Selection.Interior.Color) & " " & " is the color code for selected cell"

End Sub

Sub Count_Rules()

    MsgBox (Selection.FormatConditions.Count) & " " & "conditional formatting formulas"

End Sub

Sub Show_Selected_Object()

    MsgBox (TypeName(Selection))
    
    MsgBox ActiveSheet.name & " " & _
    TypeName(ActiveSheet) 'can be chartsheet or worksheet
    'Sheet encompasses both chartsheet and worksheet collections

End Sub

Sub Show_ErrorDescr()
MsgBox Err.Number & ": " & Error(Err.Description)

End Sub

Sub ShowDates()

    MsgBox Month(Date)
    MsgBox (Format(Date, "m"))
    
    MsgBox DateAdd("m", -2, Date)
    MsgBox DateAdd("m", -1, Date)
    
    MsgBox Format(Application.WorksheetFunction.EoMonth(Date, 1), "mm")
    MsgBox Format(Application.WorksheetFunction.EoMonth(Date, 2), "mm")
    MsgBox Format(Application.WorksheetFunction.EoMonth(Date, 3), "mm")
    
    MsgBox ("Period " & Format(Application.WorksheetFunction.EoMonth(Date, 1), "mm") & Chr(32) & Worksheets("USC_act_IC").Range("H1"))
    MsgBox "DATE(" & Format(Worksheets("Delivery Headcounts").Range("C9").Value, "yyyy,m,d") & ")"
    MsgBox "DATE(" & Format(Worksheets("Delivery Headcounts").Range("B9").Value, "yyyy,m,d") & ")"
    
    MsgBox (UCase(Format(DateAdd("m", -1, Date), "MMM")) & Chr(32) & Year(Date))
    MsgBox (UCase(Format(DateAdd("m", -2, Date), "MMM")) & Chr(32) & Year(Date))
    MsgBox (UCase(Format(DateAdd("m", -3, Date), "MMM")) & Chr(32) & Year(Date))
    
    MsgBox Format(Application.WorksheetFunction.EoMonth(Date, -2) + 1, "dd/mm/yyyy")
    MsgBox Format(Application.WorksheetFunction.EoMonth(Date, -1), "m")

End Sub


