Attribute VB_Name = "PTSlicersAngelina"
Option Explicit

Sub ReFilterSlicers()

Dim Wbk As Workbook
Dim Wks As Worksheet
Dim SC As SlicerCache
Dim SIm As SlicerItem
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'refilter all PivotTables connected to a slicer
With ActiveWorkbook.SlicerCaches("Slicer_Quarter1")     'manipulating selected items for all slicers sharing the same cache
        .SlicerItems("Q1").Selected = True
        .SlicerItems("Q2").Selected = False
        .SlicerItems("Q3").Selected = False
        .SlicerItems("Q4").Selected = False
End With

'OR filter multiple slicers at once
For Each SC In ActiveWorkbook.SlicerCaches
    
    Select Case SC.name
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Case "Slicer_Platform"
        For Each SIm In SC.SlicerItems
            If SIm.name = "desktop" Then
                SIm.Selected = True
                Else
                SIm.Selected = False
            End If
        Next SIm
        
    Case "Slicer_Week"
        For Each SIm In SC.SlicerItems
            If SIm.name = "34" Then
                SIm.Selected = True
                Else
                SIm.Selected = False
            End If
        Next SIm
    
    Case "Slicer_RepBusinessLocation"
        For Each SIm In SC.SlicerItems
            If SIm.name = "Germany" Then
                SIm.Selected = True
                Else
                SIm.Selected = False
            End If
        Next SIm
    
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    End Select
Next SC
End Sub

Sub CreateNewSlicer()

Dim Wks As Worksheet
Dim PT As PivotTable
Dim SC2 As SlicerCache
Dim SL As Slicer
Dim i As Integer


Set SC2 = ActiveWorkbook.SlicerCaches.Add2(Worksheets(1).PivotTables(1), "SourceName")        'create the Slicer Cache
SC2.Slicers.Add Worksheets(1), , "SourceName", "Select Pages", 252, 611, 144, 199                           'create the Slicer
SC2.PivotTables.AddPivotTable (Worksheets(2).PivotTables(1))                                 'links the needed PivotTables
SC2.PivotTables.AddPivotTable (Worksheets(3).PivotTables(1))
SC2.PivotTables.AddPivotTable (Worksheets(4).PivotTables(1))
SC2.PivotTables.AddPivotTable (Worksheets(5).PivotTables(1))
SC2.PivotTables.AddPivotTable (Worksheets(6).PivotTables(1))

'copy the slicer on all relevant sheets
Worksheets(1).Activate
ActiveSheet.Shapes.Range(Array("SourceName")).Select                                               'copying the slicer directly will not do, needs to be selected first
Selection.Copy

On Error Resume Next                                                                        'offsets the effect of hidden sheets
For i = 6 To ActiveWorkbook.Worksheets.Count                                                'will return an error if there is a hidden sheet
    Worksheets(i).Paste
Next i


End Sub

Sub DeleteAllSC()
Dim SC As SlicerCache
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'deleting the slicer caches also deletes all slicers using that cache
For Each SC In ActiveWorkbook.SlicerCaches
    SC.Delete
Next SC

End Sub

Sub DeleteAllSlicerShapes()
    
Dim Wks As Worksheet
Dim Sh As Shape
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
For Each Wks In ActiveWorkbook.Worksheets
    If Wks.Visible = True Then Wks.Activate
    
    If ActiveSheet.Shapes.Count > 0 Then
        For Each Sh In ActiveSheet.Shapes
            If Sh.Type = msoSlicer Then Sh.Delete                                           'deletes all slicer shapes in this workbook but retains the caches
        Next Sh
    End If
    
Next Wks

End Sub

Sub ChangeAllSlicerCaptions()
    
Dim SC As SlicerCache
Dim SL As Slicer
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each SC In ActiveWorkbook.SlicerCaches
    For Each SL In SC.Slicers
        If SL.name = "URL" Then SL.Caption = "Annie"                                        'the Sl.name is unique to each slicer, you can rename all by omitting the condition
    Next SL
Next SC

End Sub

Sub ChangeAllSlicerCaptionsInACache()

Dim SC As SlicerCache
Dim SL As Slicer
Dim SCName As String
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each SC In ActiveWorkbook.SlicerCaches
    SCName = SC.name
    
    Select Case SCName
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Case "Slicer_Platform"
        For Each SL In SC.Slicers
           SL.Caption = "Platform (Does not affect Platform comparison elements)"
        Next SL
        
    Case "Slicer_Week"
        For Each SL In SC.Slicers
           SL.Caption = "Week (Does not affect Weekly Performance Column)"
        Next SL
    
    Case "Slicer_SalesRepLocation"
        For Each SL In SC.Slicers
           SL.Caption = "Country (of Sales rep)"
        Next SL
    
    Case "Slicer_SalesRepRegion"
        For Each SL In SC.Slicers
           SL.Caption = "Region (of Sales rep)"
        Next SL
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    End Select

Next SC
End Sub

Sub ChangeSlicersNumberOfColumns()

Dim SC As SlicerCache
Dim SL As Slicer

For Each SC In ActiveWorkbook.SlicerCaches
    For Each SL In SC.Slicers
        Select Case SC.name
                Case "Slicer_SalesRepRegion": SL.NumberOfColumns = 3
                Case "Slicer_SalesRepLocation": SL.NumberOfColumns = 4
                Case "Slicer_Platform": SL.NumberOfColumns = 4
                Case "Slicer_Week": SL.NumberOfColumns = 14
                Case "Slicer_Quarter": SL.NumberOfColumns = 4
                Case "Slicer_Month": SL.NumberOfColumns = 6
        End Select
    Next SL
Next SC
End Sub

Sub MultiplePivotSlicerCaches()
      
Dim Wks As Worksheet
Dim PT As PivotTable
Dim SC As SlicerCache
Dim SL As Slicer
    
    For Each SC In ActiveWorkbook.SlicerCaches
        
        For Each PT In SC.PivotTables
            PT.Parent.Activate
            
            MsgBox SC.name & ", " & PT.name & Chr(32) & PT.TableRange1.Address              'lists all Slicer Caches and their associated PivotTables
        Next PT
    
    Next SC

End Sub
