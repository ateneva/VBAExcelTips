Attribute VB_Name = "WksAddLabels"
Option Explicit

Sub Add_New_Labels()

Dim Wks As Worksheet
Dim m As String
Range = Format(Date, "mmmm")

For Each Wks In ThisWorkbook.Worksheets

If Wks.name = "EMEA" Or Wks.name = "CEE" Or Wks.name = "FRA" Or Wks.name = "GER" Or _
Wks.name = "GWE" Or Wks.name = "IBE" Or Wks.name = "ITA" Or Wks.name = "MEMA" Or Wks.name = "UKI" Then Wks.Activate

With ActiveSheet
    Range("E18").Value = "New Starters"
    Range("E19").Value = "Left Employees"
    
    'Range("E24").Value = "TC int-cross-country import in %"
    'Range("E25").Value = "TC int-cross-country export in %"
    
    'Range("E27").Value = "3P Labour (K$)"
    'Range("E28").Value = "HP Labour (K$)"

Select Case m
    Case "November"
    Range("L28, N28, P28, R28, T28, V28, X28").Select
    Selection.Font.Color = RGB(255, 255, 255)
    
    Case "January": Range("L28").Font.Color = RGB(0, 0, 0)
    Case "April": Range("N28,T28").Font.Color = RGB(0, 0, 0)
    Case "July": Range("P28").Font.Color = RGB(0, 0, 0)
    Case "October": Range("R28,V28").Font.Color = RGB(0, 0, 0)
End Select

End With
Next Wks

End Sub

