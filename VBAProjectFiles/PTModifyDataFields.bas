Attribute VB_Name = "PTModifyDataFields"
Option Explicit

Sub ModifyDataFieldsHP()

Dim PT As PivotTable
Dim PF As PivotField
Dim PI As PivotItem
Dim i As Integer
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva, 2014
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For i = 1 To 2
Worksheets(i).Activate
Set PT = ActiveSheet.PivotTables(1)

Select Case i

Case 1
ActiveSheet.name = Format(ActiveSheet.Range("K2"), "dd-mmm")
For Each PF In PT.DataFields
'must use DataFields Collection if you are going to change the method of consolidation
    If PF.Position > 4 Then PF.Function = xlCountNums
    If PF.Position <= 4 Then PF.Function = xlSum
    If PF.Position <= 4 Then PF.NumberFormat = "0.0"
Next PF

Case 2
ActiveSheet.name = "weeks" & Format(ActiveSheet.Range("K2"), "dd-mmm")

For Each PF In PT.DataFields
'must use DataFields Collection if you are going to change the method of consolidation
    If PF.Position > 3 Then PF.Function = xlCountNums
    If PF.Position <= 3 Then PF.Function = xlSum
    If PF.Position <= 3 Then PF.NumberFormat = "0.0"
    Next PF

End Select

'PT.PivotSelect "TS IC HC Country[Romania,Croatia]", xlDataAndLabel, True
'Selection.Group
Next i
End Sub

Sub ModifyDataFields()

Dim PT As PivotTable
Dim PF As PivotField
Dim PI As PivotItem
Dim i As Integer

For i = 1 To 2
Worksheets(i).Activate
Set PT = ActiveSheet.PivotTables(1)

For Each PF In PT.DataFields

    If PF.Position > 4 Then PF.Function = xlCountNums
    If PF.Position <= 4 Then PF.Function = xlSum
    If PF.Position <= 4 Then PF.NumberFormat = "0.0"
    
Next PF

Next i
End Sub

Sub ModifyPTFieldsFunction()

Dim PT As PivotTable
Dim PF As PivotField
Dim Title As String

For Each PT In ActiveSheet.PivotTables

'.DataBodyRange --> Object in the PivotTable
'.DataRange --> Object in the PivotField and PivotItems

PT.DataBodyRange.NumberFormat = "#,#" 'formats all fields in values section
PT.DataBodyRange.NumberFormat = "#,###" 'formats all the fields currenty in the values area

'adjust the datafields
    For Each PF In PT.DataFields
        PF.Function = xlSum
        PF.NumberFormat = "#,##"
        Title = PF.name
        PF.name = Mid(Title, 8, Len(Title) - 7) & " " 'removes the sum of
    Next PF
    
PT.CalculatedFields.Add "CTR", "=PaidClicks/PaidListings*100", True
PT.CalculatedFields.Add "CPC", "=GrossRevenue/PaidClicks", True
   
    'adjust the calculated fields
    For Each PF In PT.CalculatedFields
        PF.Orientation = xlDataField
        PF.Function = xlSum
        PF.NumberFormat = "#,##0.000"
        Title = PF.name
        PF.name = Mid(Title, 8, Len(Title) - 7) & " " 'removes the sum of
    Next PF
Next PT
End Sub

Sub AddDefaultName()

Dim Wks As Worksheet
Dim PT As PivotTable
Dim PF As PivotField
Dim Title As String
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva, September 2016
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each Wks In ActiveWorkbook.Worksheets
    For Each PT In Wks.PivotTables
        
        On Error Resume Next
        For Each PF In PT.DataFields
                Title = PF.SourceName & " "
                PF.Caption = Title

        Next PF
    
    Next PT
        
Next Wks
End Sub

Sub ChangeDefaultPFName()

Dim Wks As Worksheet
Dim PT As PivotTable
Dim PF As PivotField
Dim Title As String
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva, September 2016
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each Wks In ActiveWorkbook.Worksheets
    For Each PT In Wks.PivotTables
        
        On Error Resume Next
        For Each PF In PT.DataFields
                Title = PF.name

                'comment out the line(s) that you do not need
                PF.name = Mid(Title, 8, Len(Title) - 7) & " "   'removes the "sum of", "max of", "min of"
                PF.name = Mid(Title, 10, Len(Title) - 9) & " "  'removes the "count of"
                PF.name = Mid(Title, 12, Len(Title) - 11) & " " 'removes the "average of", "product of"

        Next PF
    
    Next PT
        
Next Wks
End Sub

Sub ChangeSummaryFunctions()

Dim Wks As Worksheet
Dim PT As PivotTable
Dim PF As PivotField
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva, September 2016
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each Wks In ActiveWorkbook.Worksheets
    For Each PT In Wks.PivotTables
        For Each PF In PT.DataFields
        
            'comment as needed
            PF.Function = xlSum
            PF.Function = xlCount
            PF.Function = xlCountNums
            PF.Function = xlAverage
            PF.Function = xlProduct
            PF.Function = xlMax
            PF.Function = xlMin

        Next PF
    Next PT
        
Next Wks
End Sub

Sub ChangeNumberFormats()

Dim Wks As Worksheet
Dim PT As PivotTable
Dim PF As PivotField
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva, September 2016
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each Wks In ActiveWorkbook.Worksheets
    For Each PT In Wks.PivotTables
        For Each PF In PT.DataFields
        
            If PF.SourceName Like "*Spend" Then
            
                'comment as needed
                PF.NumberFormat = "0.00"                    'shows only two decimals
                PF.NumberFormat = "#,###"                   'shows numbers as thosands
                PF.NumberFormat = "0,##"                    'shows a leading zero in numbers smaller than a thousand

                PF.NumberFormat = "#,##0"                   'shows minus sign for negative values
                PF.NumberFormat = "#,##0_);(#,##0)"         'shows negative numbers in brackets

                PF.NumberFormat = "[$$-en-US]0.00"          'dollar currency symbol
                PF.NumberFormat = "[$€-x-euro2] #,##0.00"   'euro currency symbol
                PF.NumberFormat = "[$£-en-GB]#,##0.00"      'pound currency symbol
                PF.NumberFormat = "0%; -0%;"                'hides zero values but keeps the negative ones
                PF.NumberFormat = "0.0%"                    'formats as %
            End If
        Next PF
    Next PT
        
Next Wks
End Sub

Sub ChangePFPositionDependingOnItsName()

Dim Wks As Worksheet
Dim PT As PivotTable
Dim PF As PivotField
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva, Feb 2017
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each Wks In ActiveWorkbook.Worksheets

    For Each PT In Wks.PivotTables
    
        For Each PF In PT.DataFields

            If PF.Caption Like "*CPC*" Then PF.Position = 40

        Next PF
        
    Next PT
    
Next Wks

End Sub

Sub HideALLCalculatedFields()

Dim Wks As Worksheet
Dim PT As PivotTable
Dim PF As PivotField
Dim DF As PivotField
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each Wks In ActiveWorkbook.Worksheets

    For Each PT In Wks.PivotTables
        For Each PF In PT.CalculatedFields
        
            'trying to change the orientation of the calculated field without going through the data field will result in an error
            For Each DF In PT.DataFields
                If DF.SourceName = PF.name Then DF.Parent.PivotItems(DF.name).Visible = False
            Next DF
        Next PF
    Next PT
    
Next Wks
End Sub

Sub DeleteCalculateFields()

Dim Wks As Worksheet
Dim PT As PivotTable
Dim PF As PivotField
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva, September 2016
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each Wks In ActiveWorkbook.Worksheets

    For Each PT In Wks.PivotTables
        For Each PF In PT.CalculatedFields
           PF.Delete
        Next PF
    Next PT
        
Next Wks
End Sub

Sub RemovePTFieldsFromLayout()

Dim Wks As Worksheet
Dim PT As PivotTable
Dim PF As PivotField
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva, November 2016
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each Wks In ActiveWorkbook.Worksheets
    For Each PT In Wks.PivotTables
        For Each PF In PT.DataFields
            If PF.SourceName Like "*Paid*" Or PF.SourceName Like "*CPC*" Then PF.Orientation = xlHidden
        Next PF
    Next PT
Next Wks

End Sub

Sub Hide_Add_DataFieldsFromPT()

Dim Wks As Worksheet
Dim PT As PivotTable
Dim PF As PivotField
Dim DF As PivotField
Dim OriginalPFName As String
Dim Order As Variant
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva, Feb 2017
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each Wks In ActiveWorkbook.Worksheets
Wks.Activate
    For Each PT In ActiveSheet.PivotTables
    
        ''''''''''''''''''''''''''''''''''hiding widget impressions fields (data fields object collection)'''''''''''''''''''''''''''''''''''''''
        For Each PF In PT.DataFields
            OriginalPFName = PF.SourceName
            
            Select Case OriginalPFName
                Case "RequestedWidgetImpressions", "ServedOrganicWidgetImpressions", "ServedPaidWidgetImpressions", "BlockedPaidImpressions", "ViewedImpressions", "AvailableWidgetImpressions": PF.Orientation = xlHidden
            End Select
            
        Next PF
        
        ''''''''''''''''''''''''''''''''''adding PV fields (PivotFields object collection)'''''''''''''''''''''''''''''''''''''''''''''''''''''''
        On Error Resume Next
        For Each PF In PT.PivotFields
            OriginalPFName = PF.SourceName
            
            Select Case OriginalPFName
                Case "TotalPVs", "ServedPaidPVs", "BlockedPaidPVs", "ViewedPVs", "ViewedPaidPVs", "AvailablePVs": PF.Orientation = xlDataField
                Case "TotalPVs", "ServedPaidPVs", "BlockedPaidPVs", "ViewedPVs", "ViewedPaidPVs", "AvailablePVs": PF.NumberFormat = "#,###"
            End Select
            
        Next PF
        
        ''''''''''''''''''''''''''''''''''hding USD calculated fields (calculated fields object collection)''''''''''''''''''''''''''''''''''''''
        On Error Resume Next
        For Each PF In PT.CalculatedFields
        
            If PF.Caption Like "*USD*" Then
                    For Each DF In PT.DataFields
                        If DF.SourceName = PF.name Then DF.Parent.PivotItems(DF.name).Visible = False
                    Next DF
            End If
        Next PF
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Next PT
        
Next Wks

End Sub


