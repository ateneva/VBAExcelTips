Attribute VB_Name = "PTModifyAllFields"
Option Explicit

Sub ShowFieldinPT()

Dim Wks As Worksheet
Dim PT As PivotTable
Dim PF As PivotField
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva, Sept 2016
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each Wks In ActiveWorkbook.Worksheets
    For Each PT In Wks.PivotTables
    
        For Each PF In PT.PivotFields
            Set PF = PT.PivotFields("Country")
        
            If PF.Orientation <> xlHidden Then
            
                'comment out as needed
                    PF.Orientation = xlPageField     'as ReportFilter
                    PF.Orientation = xlRowField      'as RowField
                    PF.Orientation = xlColumnField   'as ColumnField
                    PF.Orientation = xlDataField     'as Value Field
        
            End If
        Next PF
    
    Next PT
Next Wks

End Sub
