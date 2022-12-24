Attribute VB_Name = "WbkProtection"
Option Explicit
Sub ShowUser()

ActiveCell.Value = Environ("UserName")                                                 'returns the domain name
ActiveCell.Offset(1, 0).Value = Application.UserName                                   'the user registered name
ActiveCell.Offset(2, 0).Value = Application.UserLibraryPath                            'location where COM-Addins are installed
ActiveCell.Offset(3, 0).Value = "C:\Users\" & UCase(Environ("UserName")) & "\Desktop"  'returns a custom directory
ActiveCell.Offset(4, 0).Value = ActiveWorkbook.FullName                                'the path + the name of the file <-- assimes file has been saved
ActiveCell.Offset(5, 0).Value = ActiveWorkbook.name                                    'returns the name of the file alone
ActiveCell.Offset(6, 0).Value = ActiveWorkbook.path                                    'returns the path where it is saved
ActiveCell.Offset(7, 0).Value = Application.PathSeparator                              'returns the dash

End Sub

'*******************************************************************************************
                'protection levels

'Wbk.ProtectSharing ---------> password required to view the file
'Wbk.ProtectStructure--------> password required to add/delete/hide/unhide/refresh sheets
'Wks.ProtectContents---------> password required to make changes to a worksheet
'********************************************************************************************

Sub ProtectWorkbookOpening()
 
Dim Wbk As Workbook
Dim strPwd As String
Dim strSharePwd As String
Set Wbk = Application.ActiveWorkbook
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
strPwd = InputBox("Enter password for the file")            'protects file with a password for opening
strSharePwd = InputBox("Enter password for sharing")        'assigns password for file sharing
 
Wbk.ProtectSharing Password:=strPwd, SharingPassword:=strSharePwd
 
End Sub

Sub AuthUser()

Dim Cell As Range
Dim person As String
Dim authperson As String
person = Application.UserName
'----------------------------------------
'written by Angelina Teneva, 2013
'-----------------------------------------

For Each Cell In ActiveWorkbook.Worksheets("Names").Range("E1:E100")
    authperson = Cell.Value

If person = authperson Then
        
        Worksheets("data").Visible = xlSheetVisible
    Else
        Worksheets("data").Visible = xlSheetVeryHidden
        MsgBox ("Sorry, you are not authorized to view this data")
End If

Exit For
Next Cell

End Sub

Sub KeepData()

Dim Wks As Worksheet
'~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva
'toggling sheet and workbook protection on/off with a password

If ActiveWorkbook.ProtectStructure = True Then
ActiveWorkbook.Unprotect ("ANNIE")

    For Each Wks In ActiveWorkbook.Worksheets
       If Wks.Visible = False Then Wks.Visible = True
       If Wks.Visible = xlSheetVeryHidden Then Wks.Visible = True 'unhides all very hidden sheets
       
        Wks.Activate
        If ActiveSheet.ProtectContents = True Then ActiveSheet.Unprotect ("ANNIE")
        
    Next Wks
    
    Else
    
    '--------hide confidential sheets (comment out if necessary)---------------------
    For Each Wks In ActiveWorkbook.Worksheets
        If Wks.name = "Firm R numbers" _
            Or Wks.name = "Industry Dashboard" _
            Or Wks.name = "charting" _
            Or Wks.name Like "Products*" _
            Or Wks.name Like "MResearch*" Then Wks.Visible = xlSheetVeryHidden
    
    Next Wks
    '--------protect wbk and visible sheets-------------------------------------------
    ActiveWorkbook.Protect ("ANNIE"), Structure:=True
      
    For Each Wks In ActiveWorkbook.Worksheets
        If Wks.Visible = True Then Wks.Activate
        
        ActiveSheet.Protect ("ANNIE"), DrawingObjects:=True, Contents:=True, _
        Scenarios:=True, AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:=True
    
    Next Wks

End If
End Sub

Sub KeepOutbrainData()

Dim Wks As Worksheet
'~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva
'toggling sheet and workbook structure protection on/off with a password

If ActiveWorkbook.ProtectStructure = True Then 'if workbook strucuture is protected (i.e. visible)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ActiveWorkbook.Unprotect ("outdash") 'unprotect workbook
    
    'unhide hidden sheets and unprotect sheets
    For Each Wks In ActiveWorkbook.Worksheets
        If Wks.Visible = False Then Wks.Visible = True
        If Wks.Visible = xlVeryHidden Then Wks.Visible = True
        
            Wks.Activate
            If ActiveSheet.ProtectContents = True Then ActiveSheet.Unprotect ("inhead")
    Next Wks
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Else 'i.e. if workbook is not proected
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ActiveWorkbook.Protect ("outdash"), Structure:=True 'protect workbook
    
    'protect sheets but allow PivotTables to be filtered
    If Wks.Visible = True Then Wks.Activate
        ActiveSheet.Protect ("inhead"), DrawingObjects:=True, Contents:=True, _
        Scenarios:=True, AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:=True
    Next Wks
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
End If

End Sub

Sub ProtectAllSheetsPassUserDefinedInput()
'written by Angelina Teneva

Dim wSheet As Worksheet
Dim Pwd As String
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Pwd = InputBox("Enter your password to protect all worksheets", "Password Input")

For Each wSheet In Worksheets
    wSheet.Protect Password:=Pwd
Next wSheet

End Sub

Sub UnProtectAllPassUserDefinedInput()
'written by Angelina Teneva

Dim wSheet As Worksheet
Dim Pwd As String

Pwd = InputBox("Enter your password to unprotect all worksheets", "Password Input")
On Error Resume Next
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each wSheet In Worksheets
    wSheet.Unprotect Password:=Pwd
Next wSheet

If Err <> 0 Then
    MsgBox "You have entered an incorect password. All worksheets could not " & _
    "be unprotected.", vbCritical, "Incorect Password"
End If

On Error GoTo 0

End Sub

Sub ProtectSheetsWithDefinedPass()
'written by Angelina Teneva

Dim Wks As Worksheet

For Each Wks In ThisWorkbook.Worksheets
    Wks.Activate
    If ActiveSheet.ProtectContents = True Then ActiveSheet.Unprotect ("12q")
Next Wks

ThisWorkbook.RefreshAll

End Sub

Sub UnprotectSheetsNoPass()
'unprotectting multuple non-password protected sheets
'written by Angelina Teneva

Dim Wks As Worksheet

For Each Wks In ActiveWorkbook.Worksheets
    Wks.Activate
    If ActiveSheet.ProtectContents = True Then ActiveSheet.Unprotect
Next Wks

End Sub

Sub PasswordBreaker()

Dim i As Integer, j As Integer, k As Integer
Dim l As Integer, m As Integer, n As Integer
Dim i1 As Integer, i2 As Integer, i3 As Integer
Dim i4 As Integer, i5 As Integer, i6 As Integer
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error Resume Next

For i = 65 To 66: For j = 65 To 66: For k = 65 To 66
For l = 65 To 66: For m = 65 To 66: For i1 = 65 To 66
For i2 = 65 To 66: For i3 = 65 To 66: For i4 = 65 To 66
For i5 = 65 To 66: For i6 = 65 To 66: For n = 32 To 126

ActiveSheet.Unprotect Chr(i) & Chr(j) & Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _
Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)

If ActiveSheet.ProtectContents = False Then
    MsgBox "One usable password is " & Chr(i) & Chr(j) & Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & """" & Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
Exit Sub

End If
Next: Next: Next: Next: Next: Next
Next: Next: Next: Next: Next: Next
End Sub
