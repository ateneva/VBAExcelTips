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
Dim SC As SlicerCache
Dim Sl As Slicer
Dim i As Integer


Set SC = ActiveWorkbook.SlicerCaches.Add2(Worksheets(1).PivotTables(1), "URL")              'create the Slicer Cache
SC.Slicers.Add Worksheets(1), , "URL", "Link", 252, 611, 144, 199                           'create the Slicer
SC.PivotTables.AddPivotTable (Worksheets(2).PivotTables(1))                                 'links the needed PivotTables
SC.PivotTables.AddPivotTable (Worksheets(3).PivotTables(1))

'copy the slicer on all relevant sheets
Worksheets(1).Activate
ActiveSheet.Shapes.Range(Array("URL")).Select                                               'copying the slicer directly will not do, needs to be selected first
Selection.Copy

On Error Resume Next                                                                        'offsets the effect of hidden sheets
For i = 2 To ActiveWorkbook.Worksheets.Count                                                'will return an error if there is a hidden sheet
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
Dim Sl As Slicer
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each SC In ActiveWorkbook.SlicerCaches
    For Each Sl In SC.Slicers
        If Sl.name = "URL" Then Sl.Caption = "Annie"                                        'the Sl.name is unique to each slicer, you can rename all by omitting the condition
    Next Sl
Next SC

End Sub

Sub ChangeAllSlicerCaptionsInACache()

Dim SC As SlicerCache
Dim Sl As Slicer
Dim SCName As String
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each SC In ActiveWorkbook.SlicerCaches
    SCName = SC.name
    
    Select Case SCName
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Case "Slicer_Platform"
        For Each Sl In SC.Slicers
           Sl.Caption = "Platform (Does not affect Platform comparison elements)"
        Next Sl
        
    Case "Slicer_Week"
        For Each Sl In SC.Slicers
           Sl.Caption = "Week (Does not affect Weekly Performance Column)"
        Next Sl
    
    Case "Slicer_SalesRepLocation"
        For Each Sl In SC.Slicers
           Sl.Caption = "Country (of Sales rep)"
        Next Sl
    
    Case "Slicer_SalesRepRegion"
        For Each Sl In SC.Slicers
           Sl.Caption = "Region (of Sales rep)"
        Next Sl
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    End Select

Next SC
End Sub

Sub ChangeSlicersNumberOfColumns()

Dim SC As SlicerCache
Dim Sl As Slicer

For Each SC In ActiveWorkbook.SlicerCaches
    For Each Sl In SC.Slicers
        Select Case SC.name
                Case "Slicer_SalesRepRegion": Sl.NumberOfColumns = 3
                Case "Slicer_SalesRepLocation": Sl.NumberOfColumns = 4
                Case "Slicer_Platform": Sl.NumberOfColumns = 4
                Case "Slicer_Week": Sl.NumberOfColumns = 14
                Case "Slicer_Quarter": Sl.NumberOfColumns = 4
                Case "Slicer_Month": Sl.NumberOfColumns = 6
        End Select
    Next Sl
Next SC
End Sub

Sub MultiplePivotSlicerCaches()
      
Dim Wks As Worksheet
Dim PT As PivotTable
Dim SC As SlicerCache
Dim Sl As Slicer
    
    For Each SC In ActiveWorkbook.SlicerCaches
        
        For Each PT In SC.PivotTables
            PT.Parent.Activate
            
            MsgBox SC.name & ", " & PT.name & Chr(32) & PT.TableRange1.Address              'lists all Slicer Caches and their associated PivotTables
        Next PT
    
    Next SC

End Sub
