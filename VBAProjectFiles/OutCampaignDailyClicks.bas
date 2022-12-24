Attribute VB_Name = "OutCampaignDailyClicks"
Option Explicit

Sub SummarizeCampaignData()

Dim Wks As Worksheet
Dim PTCache As PivotCache
Dim PT As PivotTable
Dim PF As PivotField
Dim i As Integer
Dim Title As String
Dim Field As String
'~~~~~~~~~~~~~~~~~~~~~~~~

'Create the cache from a normal cell reference
Set PTCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, sourcedata:=Worksheets("data").UsedRange)

'**************************************************************************************************************
'create 3 worksheets and Beautify them
Worksheets.Add.name = "summary"
'Worksheets.Add.name = "WidgetLevel"

For Each Wks In ActiveWorkbook.Worksheets

    If Wks.name <> "data" And Wks.name <> "interface" Then
    
        Wks.Activate
        'beautify worksheet
            ActiveWindow.Zoom = 80
            ActiveWindow.DisplayGridlines = False
        
        'create_pivottable with datafields
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    With ActiveSheet
        'Create the pivot table from the created cache and apply pivottable style
        Set PT = ActiveSheet.PivotTables.Add(PivotCache:=PTCache, TableDestination:=Range("A5"))
        
        'set PT default settings
        PT.HasAutoFormat = False 'prevent autosort columns
        PT.EnableDrilldown = False 'prevent users from reaching the raw data
        PT.ColumnGrand = False 'turn off the column totals
        PT.RowGrand = False 'turn off the column totals
        PT.DisplayErrorString = True 'shows nothing on DIV errors
        PT.TableStyle2 = "PivotStyleLight19"
        PT.RowAxisLayout xlTabularRow  'changes to tabular orientation
        PT.DisplayFieldCaptions = False 'removes filtering buttons
        PT.ShowDrillIndicators = False 'turns off drill indicators
        
        'add PivotFields common for all PivotTables
        For Each PF In PT.PivotFields
            
            'add Publisher and Platform as a ReportFilter
                                              
            If PF.name = "Campaign" Then PF.Orientation = xlPageField
                               
'            If PF.name = "Campaign" Then
'                PF.Orientation = xlRowField
'                PF.Subtotals(1) = False
'            End If
            
            If PF.name = "SourceURL" Then PF.Orientation = xlRowField
            If PF.name = "Date" Then PF.Orientation = xlPageField
           
        Next PF
        
        On Error Resume Next 'to offset the loop from trying to create the same calculated fields twice
        'create the calculated fields for PageLevel
'        PT.CalculatedFields.Add "CTR", "=Clicks/Listings", True
'        PT.CalculatedFields.Add "CPC", "=Spend/Clicks", True
'        PT.CalculatedFields.Add "CPM", "=Spend/(Listings/1000)", True
        PT.CalculatedFields.Add "ConversionRate", "=Conversions/Clicks", True
        PT.CalculatedFields.Add "CPA", "=Spend/Conversions", True
        
        'adds all available data fields
        For i = 5 To PT.PivotFields.Count
            PT.PivotFields(i).Orientation = xlDataField
        Next i
        
        'change the orientation of the values
        PT.DataPivotField.Orientation = xlColumnField
        
        'adjust the retrieved datafields to a presentble format
        For Each PF In PT.DataFields
            PF.Function = xlSum
                 
            If PF.name Like "*CPM*" Or PF.name Like "*CPC*" Or PF.name Like "*CPA*" Then PF.NumberFormat = "[$$-en-US]0.00"
            If PF.name Like "*CTR*" Or PF.name Like "*Rate*" Then PF.NumberFormat = "0.0%"
            If PF.Position = 1 Or PF.Position <= 2 Then PF.NumberFormat = "#,##"
                               
            Title = PF.name
            PF.name = Mid(Title, 8, Len(Title) - 7) & " " 'removes the sum of
        Next PF
        
        'adjust sorting and layouts for row fields --> it needs to be in a separate loop; if defined when the field is being added, it's not working
        For Each PF In PT.RowFields
            If PF.name = "SourceURL" Then PF.AutoSort xlDescending, "Clicks "
        Next PF
        
      PT.ShowPages PageField:="Campaign" 'whether it should add each item in the PageField as a separate PT on a separate sheet
 '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'adjust column width
    Cells.EntireColumn.AutoFit
       
    End With
    End If

    'protect and hide source data
    If Wks.name = "data" Then

        Wks.Protect ("inhead"), DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:=True
        Wks.Visible = xlSheetVeryHidden
    End If
    
        Application.DisplayAlerts = False 'prevents a pop-out message
    If Wks.name = "summary" Then Wks.Delete 'deletes the original Wks summary
    Application.DisplayAlerts = True

Next Wks

'beautify the tab
For Each Wks In ActiveWorkbook.Worksheets

    If Wks.Visible = True Then
        Wks.Activate
            With ActiveSheet
                ActiveWindow.Zoom = 80
                ActiveWindow.DisplayGridlines = False
                Cells.EntireColumn.AutoFit       'adjust column width
                Rows("2:2").EntireRow.Hidden = True
            End With
    End If
Next Wks

'On Error Resume Next
'Call AddOutbrainlogo ''add the Outbrain Logo on top of each visible sheet

End Sub

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
ActiveSheet.Pictures.Insert ("C:\Users\Angelina\Desktop\Outbrain.png")

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
ActiveSheet.Pictures.Insert ("C:\Users\Angelina\Desktop\Outbrain.png")

For Each Sh In ActiveSheet.Shapes

If Sh.Type = msoPicture Or Sh.Type = msoLinkedPicture Then
Sh.ScaleWidth 0.5012441057, msoFalse, msoScaleFromTopLeft
Sh.ScaleHeight 0.5012437596, msoFalse, msoScaleFromTopLeft
End If
Next Sh

End If
Next Wks
End Sub
