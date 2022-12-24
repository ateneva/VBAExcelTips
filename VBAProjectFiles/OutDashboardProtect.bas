Attribute VB_Name = "OutDashboardProtect"
Option Explicit

Sub OutbrainDataProtection()

Dim Wks As Worksheet
Dim update As String
'~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva
'toggling sheet and workbook protection on/off with a password

update = MsgBox("Would you like to update the dashboard", vbYesNo)
Select Case update

Case vbYes
For Each Wks In ActiveWorkbook.Worksheets

On Error Resume Next
    If Wks.Visible = False Then Wks.Visible = True
    If Wks.Visible = xlSheetVeryHidden Then Wks.Visible = True
    
    Wks.Activate
    If ActiveSheet.ProtectContents = True Then ActiveSheet.Unprotect ("inhead")
       
Next Wks

Case vbNo
For Each Wks In ActiveWorkbook.Worksheets
    'drawing objects = False allows the editing of Slicers
    If Wks.name = "Slicers" Or Wks.name = "Targets" Or Wks.name = "QTD" Or Wks.name = "data" Then
    
        Wks.Protect ("inhead"), DrawingObjects:=False, Contents:=True, _
            Scenarios:=True, AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:=True
        
        Wks.Visible = xlSheetVeryHidden
    End If
    
    'drawing objects = False allows the editing of Slicers
    If Wks.Visible = True Then Wks.Protect ("inhead"), DrawingObjects:=False, Contents:=True, _
        Scenarios:=True, AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:=True

Next Wks
End Select

If ThisWorkbook.Worksheets("data").Visible = True Then Application.Run "UpdateDashboard"

End Sub

