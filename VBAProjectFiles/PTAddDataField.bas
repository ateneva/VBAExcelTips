Attribute VB_Name = "PTAddDataField"
Option Explicit

Sub AddDataField()

Dim Wks As Worksheet
Dim PC As PivotCache
Dim PT As PivotTable
Dim PF As PivotField
Dim i As Integer
Dim FieldPosition As Integer
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each Wks In ActiveWorkbook.Worksheets
    Wks.Activate
    i = ActiveSheet.Index       'not having the Worksheet Index condition causes the field to be added multiple times to the same sheet
    If i >= 5 And i <= 12 Then
    
        With ActiveSheet
      
        For Each PT In ActiveSheet.PivotTables
            FieldPosition = PivotFields("Paid Coverage").Position - 1
            
            For Each PF In PT.PivotFields
                If PF.name = "PaidWidgetFillRate" Then
                    PF.Orientation = xlDataField
                    PF.Position = FieldPosition
                    PF.NumberFormat = "0.00%"
                End If
            Next PF
        Next PT
        End With
    End If
Next Wks

End Sub

Sub AddDataField2()

Dim Wks As Worksheet
Dim PC As PivotCache
Dim PT As PivotTable
Dim PF As PivotField
Dim i As Integer
Dim FieldPosition As Integer
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each Wks In ActiveWorkbook.Worksheets
    Wks.Activate
   
        With ActiveSheet
      
        For Each PT In ActiveSheet.PivotTables
                       
            For Each PF In PT.PivotFields
                If PF.name Like "*PV*" Or PF.name Like "*Listing*" Then
                    PF.Orientation = xlDataField
                End If
            Next PF
        Next PT
        End With
Next Wks

End Sub

Sub HideRowField()

Dim Wks As Worksheet
Dim PC As PivotCache
Dim PT As PivotTable
Dim PF As PivotField
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

With ActiveSheet
      
    For Each PT In ActiveSheet.PivotTables
       If PT.name = "keywords" Then
       
            For Each PF In PT.PivotFields
                If PF.name Like "kw*" Then PF.Orientation = xlHidden
            Next PF
            
        End If
    Next PT
    
End With
End Sub

Sub AddRowFieldInCertainPosition()

Dim Wks As Worksheet
Dim PC As PivotCache
Dim PT As PivotTable
Dim PF As PivotField
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
For Each Wks In ActiveWorkbook.Worksheets
    Wks.Activate
   
    With ActiveSheet
      
    For Each PT In ActiveSheet.PivotTables
       
            For Each PF In PT.PivotFields
                If PF.name Like "Date*" Then
                    PF.Orientation = xlRowField
                    PF.Position = 4
                End If
            Next PF
            
    Next PT
    
    End With
Next Wks
End Sub



