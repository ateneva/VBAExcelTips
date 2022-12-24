Attribute VB_Name = "EventsWbkOpenUnprotect"
Option Explicit

Private Sub Workbook_Open()

Dim Wks As Worksheet
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'removes protection and unhides very hidden sheets

If Application.UserName = "Angelina" Then
    'Application.Run "OutbrainDataProtection"
    
    For Each Wks In ThisWorkbook.Worksheets
    
    On Error Resume Next
        If Wks.Visible = False Then Wks.Visible = True
        If Wks.Visible = xlSheetVeryHidden Then Wks.Visible = True
        
        Wks.Activate
        If ActiveSheet.ProtectContents = True Then ActiveSheet.Unprotect ("inhead")
     Next Wks
    
    If ThisWorkbook.Worksheets("data").Visible = True Then
        ThisWorkbook.Worksheets("data").Unprotect ("inhead")
        Application.Run "UpdateDashboard"
    End If
    
End If
    
End Sub
