Attribute VB_Name = "OutBrainLogo"
Option Explicit

Sub AddOutbrainlogo()

Dim Wks As Worksheet
Dim Sh As Shape

Dim Cell As Range
'******************************************

For Each Wks In ActiveWorkbook.Worksheets
If Wks.Visible = True Then Wks.Activate

    If ActiveSheet.Shapes.Count > 0 Then
    
    For Each Sh In ActiveSheet.Shapes
        If Sh.Type = msoPicture Or Sh.Type = msoLinkedPicture Then Sh.Delete   'removes previous logo (the code assumes that the only picture in the respective tab is the previous logo and there are no other pictures that should remain there)
    Next Sh
   
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Set Cell = ActiveSheet.Range("A1")
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
       
    Cell.Select 'makes sure the logo is always inserted in the same cell
    ActiveSheet.Pictures.Insert ("C:\Users\Angelina\Desktop\logo.png")
    
    For Each Sh In ActiveSheet.Shapes 'centers picture in cell
        If Sh.Type = msoPicture Or Sh.Type = msoLinkedPicture Then
        
            Sh.ScaleWidth 0.5012441057, msoFalse, msoScaleFromTopLeft
            Sh.ScaleHeight 0.5012437596, msoFalse, msoScaleFromTopLeft
        End If
    Next Sh

Else

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Set Cell = ActiveSheet.Range("A1")
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Cell.Select
    ActiveSheet.Pictures.Insert ("C:\Users\Angelina\Desktop\logo.png")
    
    For Each Sh In ActiveSheet.Shapes
    
    If Sh.Type = msoPicture Or Sh.Type = msoLinkedPicture Then
        Sh.ScaleWidth 0.5012441057, msoFalse, msoScaleFromTopLeft
        Sh.ScaleHeight 0.5012437596, msoFalse, msoScaleFromTopLeft
    End If
    Next Sh

End If

Next Wks

End Sub

