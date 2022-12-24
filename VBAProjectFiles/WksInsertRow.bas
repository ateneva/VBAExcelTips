Attribute VB_Name = "WksInsertRow"
Option Explicit

Sub InsertRow()

Dim Wks As Worksheet
Dim i As Integer
Dim SL As String
'***********************

For i = 5 To ThisWorkbook.Worksheets.Count
Worksheets(i).Activate

With ActiveSheet
Rows("16:16").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
Range("E16").Value = "Delivery Headcount Chargeable"
Range("G17:Y17").Copy Range("G16")

SL = ActiveSheet.Range("I2").Value

Select Case SL
    Case "DELIVERY SCORECARD (DCC/CFS)": Range("A16").Value = "Delivery Chargeable DCC-CFS"
    Case "DELIVERY SCORECARD (DCC/IC&Cloud)": Range("A16").Value = "Delivery Chargeable DCC-IC&Cloud"
    Case "DELIVERY SCORECARD (NMC = 1Z+5V)": Range("A16").Value = "Delivery Chargeable NWS (1Z)"
    Case "DELIVERY SCORECARD (SIS)": Range("A16").Value = "Delivery Chargeable SC (6C)"
    Case "DELIVERY SCORECARD (TC excl. EDU)": Range("A16").Value = "Delivery Chargeable TC excl. EDU"

End Select
End With
Next i
'~~~~~~~~~~~~~~~~MTD & YTD~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For i = 3 To 4
Worksheets(i).Activate

With ActiveSheet
    Rows("17:17").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
    Range("E17").Value = "Delivery Headcount Chargeable"
    Range("DQ18:EH18").Copy Range("DQ17")

SL = ActiveSheet.Range("DR2").Value

Select Case SL
    Case "DELIVERY SCORECARD (DCC/CFS)": Range("A17").Value = "Delivery Chargeable DCC-CFS"
    Case "DELIVERY SCORECARD (DCC/IC&Cloud)": Range("A17").Value = "Delivery Chargeable DCC-IC&Cloud"
    Case "DELIVERY SCORECARD (NMC = 1Z+5V)": Range("A17").Value = "Delivery Chargeable NWS (1Z)"
    Case "DELIVERY SCORECARD (SIS)": Range("A17").Value = "Delivery Chargeable SC (6C)"
    Case "DELIVERY SCORECARD (TC excl. EDU)": Range("A16").Value = "Delivery Chargeable TC excl. EDU"
End Select
End With
Next i

End Sub
