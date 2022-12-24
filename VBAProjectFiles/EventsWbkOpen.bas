Attribute VB_Name = "EventsWbkOpen"
Option Explicit

Private Sub Workbook_Open_Unprotect()

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

Private Sub Workbook_Open_ChangePaths()

'~~~~~~~~~~~~~~~~~~~~change paths depending on user~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Dim Cell As Range 'triggered upon Open event in the Headcount, RMR and the Available capacity reports
Application.Calculation = xlCalculationAutomatic

Worksheets("MACROS").Activate

If Application.UserName = "Angelina Teneva" Then

'~~~~~~~~~~~~~updates string values~~~~~~~~~~~~~~~~~~~~
    For Each Cell In ActiveSheet.Range("A4:A24")
        Cell.Value = Cell.Offset(0, 10)
    Next Cell
    
    Else
    
    For Each Cell In ActiveSheet.Range("A4:A24")
        Cell.Value = Cell.Offset(0, 12)
    Next Cell
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
End If

End Sub

Private Sub Workbook_Open_Refresh()

If Application.UserName = "Angelina" Then
    ThisWorkbook.RefreshAll
    ThisWorkbook.saveas "\\VBOXSVR\Virtual_Machine_\Dashboards\Plarium_Weekly\plarium_conversion_source_" & Format(Date, "yyyymmdd") & ".xlsb"
    ThisWorkbook.saveas "C:\Users\Angelina\Dashboards\weeklyreports\plarium_conversion_source_" & Format(Date, "yyyymmdd") & ".xlsb"
    ThisWorkbook.Close
End If
End Sub


