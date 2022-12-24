Attribute VB_Name = "WksToggleProtection"
Option Explicit

Sub KeepData()

Dim Wks As Worksheet
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva, 2013
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

If ActiveWorkbook.ProtectStructure = True Then
ActiveWorkbook.Unprotect ("annie")

    For Each Wks In ActiveWorkbook.Worksheets
        If Wks.Visible = False Then Wks.Visible = True
        
        Wks.Activate
        If ActiveSheet.ProtectContents = True Then ActiveSheet.Unprotect ("annie")
    
    Next Wks

Else

    ActiveWorkbook.Protect ("annie"), Structure:=True

    For Each Wks In ActiveWorkbook.Worksheets
    If Wks.Visible = True Then Wks.Activate
    
        ActiveSheet.Protect ("annie"), DrawingObjects:=True, Contents:=True, _
        Scenarios:=True, AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:=True
    
    Next Wks
End If
End Sub

