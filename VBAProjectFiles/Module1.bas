Attribute VB_Name = "Module1"
Option Explicit

Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    ActiveSheet.Shapes.Range(Array("Picture 1")).Select
    Selection.ShapeRange.ScaleWidth 0.2758803271, msoFalse, msoScaleFromTopLeft
    Selection.ShapeRange.ScaleHeight 0.2758803437, msoFalse, msoScaleFromTopLeft
    Selection.ShapeRange.ScaleWidth 1.1736462074, msoFalse, msoScaleFromTopLeft
    Selection.ShapeRange.ScaleHeight 1.1736464983, msoFalse, msoScaleFromTopLeft
    
    Selection.ShapeRange.IncrementLeft -2.0454330709
    Selection.ShapeRange.IncrementTop 8.8637007874
    Range("B220").Select
End Sub
