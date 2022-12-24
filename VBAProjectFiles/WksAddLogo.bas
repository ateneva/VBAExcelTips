Attribute VB_Name = "WksAddLogo"
Option Explicit

Sub logo()

Dim Wks As Worksheet
Dim Sh As Shape

Dim x As Integer
Dim y As Integer

Dim Cell As Range
'******************************************
'written by Angelina Teneva, 2013
'******************************************

For Each Wks In ActiveWorkbook.Worksheets
If Wks.Visible = True Then Wks.Activate

If ActiveSheet.Shapes.Count > 0 Then   'replaces previous logo
    
    'the code assumes that the only picture in the respective tab is the previous logo
    'and there are no other pictures that should remain there)
        
        For Each Sh In ActiveSheet.Shapes
            If Sh.Type = msoPicture Then Sh.Delete  'removes previous logo
        Next Sh
   
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Set Cell = ActiveSheet.Range("B2")
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
       
    Cell.Select 'makes sure the logo is always inserted in the same cell
    ActiveSheet.Pictures.Insert ("C:\Users\hp\Desktop\logo.png")
    
    For Each Sh In ActiveSheet.Shapes 'centers picture in cell
        If Sh.TopLeftCell.Address(0, 0) = "B2" Then
        
            Sh.Height = 33
            Sh.width = 79
            Sh.Top = 10
        
        End If
    Next Sh

Else
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'adds a new brand logo
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Set Cell = ActiveSheet.Range("B2")
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Cell.Select
    ActiveSheet.Pictures.Insert ("C:\Users\hp\Desktop\logo.png")
    
    For Each Sh In ActiveSheet.Shapes
        If Sh.TopLeftCell.Address(0, 0) = "B2" Then
        
            Sh.LockAspectRatio = msoTrue    'locks width-to-height ration
            Sh.Height = 33
            Sh.width = 79
            Sh.Top = 10

        End If
    Next Sh

End If
Next Wks
End Sub

Sub ShowShapeDimensions()

Dim l As Long
Dim t As Long
Dim h As Long
Dim w As Long

Dim Sh As Shape
'----------------------------------------------
    For Each Sh In ActiveSheet.Shapes
        If Sh.TopLeftCell.Address(0, 0) = "B2" Then
            l = Sh.Left
            t = Sh.Top
            h = Sh.Height
            w = Sh.width
        End If
    Next Sh

'Return Dimensions
    ActiveSheet.Range("A1") = "Left: " & l
    ActiveSheet.Range("A2") = "Top: " & t
    ActiveSheet.Range("A3") = "Height: " & h
    ActiveSheet.Range("A4") = "Width: " & w

End Sub

