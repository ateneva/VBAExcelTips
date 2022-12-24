Attribute VB_Name = "PTFiltering"
Option Explicit

Sub PTFiltering()

Dim PT As PivotTable
Dim PF As PivotField

For Each PT In ActiveWorkbook.Worksheets

'**********************************************************filtering*********************************************************************************
    PT.AllowMultipleFilters = True                                             'allows multiple filters to be set on pivot fields
    
    PT.PivotFields("Description").ClearAllFilters                              'if the field has any previously set filters
    PT.PivotFields("Description").ClearManualFilter                            'clears only manual fitlers but retains items and value filters if any
    PT.PivotFields("Description").ClearLabelFilters                            'clears only label fitlers
    PT.PivotFields("Description").ClearValueFilters                            'clears only value fitlers

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                                                'DateFilters
                                                
            '.Add2 will only work in Excel 2013 or later
            '.Add will only work in Excel 2010 or earlier --> This object, member, or enumeration is deprecated and is not intended to be used in your code.
            
            'if you have users with different version, best use an If-Then-Else Statement
            
    If Application.VERSION = "14.0" Then
                                                
    'past period
    PF.PivotFilters.Add xlDateYesterday
    PF.PivotFilters.Add xlDateLastMonth                                   'assumes pivot field has been declared
    PF.PivotFilters.Add xlDateLastQuarter
    PF.PivotFilters.Add xlDateLastWeek
    PF.PivotFilters.Add xlDateLastYear
    
    PF.PivotFilters.Add xlBefore, Value1:="01-07-2016"
    PF.PivotFilters.Add xlBeforeOrEqualTo, Value1:="01-07-2016"
        
    Else
    
    'present period
    PF.PivotFilters.Add2 xlDateToday
    PF.PivotFilters.Add2 xlDateThisMonth
    PF.PivotFilters.Add2 xlDateThisQuarter
    PF.PivotFilters.Add2 xlDateThisWeek
    PF.PivotFilters.Add2 xlDateThisYear
    
    'future period
    PF.PivotFilters.Add2 xlDateTomorrow
    PF.PivotFilters.Add2 xlDateNextMonth
    PF.PivotFilters.Add2 xlDateNextQuarter
    PF.PivotFilters.Add2 xlDateNextWeek
    PF.PivotFilters.Add2 xlDateNextYear
    
    PF.PivotFilters.Add2 xlAfter, Value1:="01-07-2016"
    PF.PivotFilters.Add2 xlAfterOrEqualTo, Value1:="01-07-2016"
    
    'all in calendar month
    PF.PivotFilters.Add2 xlAllDatesInPeriodJanuary
    PF.PivotFilters.Add2 xlAllDatesInPeriodFebruary
    PF.PivotFilters.Add2 xlAllDatesInPeriodMarch
    PF.PivotFilters.Add2 xlAllDatesInPeriodApril
    PF.PivotFilters.Add2 xlAllDatesInPeriodMay
    PF.PivotFilters.Add2 xlAllDatesInPeriodJune
    PF.PivotFilters.Add2 xlAllDatesInPeriodJuly
    PF.PivotFilters.Add2 xlAllDatesInPeriodAugust
    PF.PivotFilters.Add2 xlAllDatesInPeriodSeptember
    PF.PivotFilters.Add2 xlAllDatesInPeriodDecember
    
    'all dates in calendar quarter
    PF.PivotFilters.Add2 xlAllDatesInPeriodQuarter1
    PF.PivotFilters.Add2 xlAllDatesInPeriodQuarter2
    PF.PivotFilters.Add2 xlAllDatesInPeriodQuarter3
    PF.PivotFilters.Add2 xlAllDatesInPeriodQuarter4
    PF.PivotFilters.Add2 xlYearToDate
    
    PF.PivotFilters.Add2 xlDateBetween, Value1:="01-07-2016", Value2:="31-08-2016"
    PF.PivotFilters.Add2 xlDateNotBetween, Value1:="01-07-2016", Value2:="31-08-2016"
    
    End If
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                                                'CaptionFilters
    
    '.Add2 will only work in Excel 2013 or later
    '.Add will only work in Excel 2010 or earlier --> This object, member, or enumeration is deprecated and is not intended to be used in your code.
            
    'if you have users with different version, best use an If-Then-Else Statement
    
    If Application.VERSION = "14.0" Then
    
    '.Add Type is only available in Excel 2010
    PT.PivotFields("Description").PivotFilters.Add xlCaptionEquals, Value1:="Sales"
    PT.PivotFields("City").PivotFilters.Add xlCaptionContains, Value1:="LO"
    PT.PivotFields("City").PivotFilters.Add xlCaptionIsBetween, Value1:="M", Value2:="U" 'shows all the cities that statt wit a letter b/n M and O
    
    Else
    
    PF.PivotFilters.Add2 xlCaptionEquals, Value1:="134"
    PF.PivotFilters.Add2 xlCaptionDoesNotEqual, Value1:="134"
    
    PF.PivotFilters.Add2 xlCaptionBeginsWith, Value1:="134"
    PF.PivotFilters.Add2 xlCaptionDoesNotBeginWith, Value1:="134"
    
    PF.PivotFilters.Add2 xlCaptionEndsWith, Value1:="134"
    PF.PivotFilters.Add2 xlCaptionDoesNotEndWith, Value1:="134"
    
    PF.PivotFilters.Add2 xlCaptionContains, Value1:="134"
    PF.PivotFilters.Add2 xlCaptionDoesNotContain, Value1:="134"
   
    PF.PivotFilters.Add2 xlCaptionIsGreaterThan, Value1:="134"
    PF.PivotFilters.Add2 xlCaptionIsGreaterThanOrEqualTo, Value1:="134"
    
    PF.PivotFilters.Add2 xlCaptionIsLessThan, Value1:="134"
    PF.PivotFilters.Add2 xlCaptionIsLessThanOrEqualTo, Value1:="134"
    
    PF.PivotFilters.Add2 xlCaptionIsBetween, Value1:="M", Value2:="U"
    PF.PivotFilters.Add2 xlCaptionIsNotBetween, Value1:="M", Value2:="U"
    
    End If
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                                                'ValueFilters
                                                
    '.Add2 will only work in Excel 2013 or later
    '.Add will only work in Excel 2010 or earlier --> This object, member, or enumeration is deprecated and is not intended to be used in your code.
    
    If Application.VERSION = "14.0" Then

    PT.PivotFields("Quarter").PivotFilters.Add xlValueIsGreaterThan, PT.PivotFields("Sum of Visits (000)"), Value1:=1000
    PT.PivotFields("Purpose").PivotFilters.Add xlValueIsGreaterThan, PT.PivotFields("Sum of Visits (000)"), Value1:=1000
    
    Else
    
    PF.PivotFilters.Add2 xlValueEquals, PT.PivotFields("PaidClicks "), Value1:=7281
    PF.PivotFilters.Add2 xlValueDoesNotEqual, PT.PivotFields("PaidClicks "), Value1:=7281
    
    PF.PivotFilters.Add2 xlValueIsGreaterThan, PT.PivotFields("PaidClicks "), Value1:=7281
    PF.PivotFilters.Add2 xlValueIsGreaterThanOrEqualTo, PT.PivotFields("PaidClicks "), Value1:=7281
    
    PF.PivotFilters.Add2 xlValueIsLessThan, PT.PivotFields("PaidClicks "), Value1:=7281
    PF.PivotFilters.Add2 xlValueIsLessThanOrEqualTo, PT.PivotFields("PaidClicks "), Value1:=7281
        
    PF.PivotFilters.Add2 xlValueIsBetween, PT.PivotFields("PaidClicks "), Value1:=5000, Value2:=10000
    PF.PivotFilters.Add2 xlValueIsNotBetween, PT.PivotFields("PaidClicks "), Value1:=5000, Value2:=10000
    
    End If
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                                                'TopCountFilters

   PT.PivotFields("City").PivotFilters.Add2 xlTopCount, PT.PivotFields("Sales"), Value1:=3      'top 3 cities with the greater value of sales
   PT.PivotFields("City").PivotFilters.Add2 xlTopPercent, PT.PivotFields("Sales"), Value1:=3    'top 3 cities whose sales account for the greatest proportion
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'*****************************************************************************************************************************************************

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~excluding/including items;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'field orientation doesn't matter although I would usually prefer caption filters for Row and Column Fields~
Set PF = PT.PivotFields("Import")

For Each PI In PF.PivotItems
    If PI.Value = "7000O02" Then PI.Visible = False
    If PI.Value = "3000164" Then PI.Visible = False
Next PI
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
For Each PI In PF.PivotItems
      If PI.Value = "S" Then PI.Visible = False
      If PI.Value = "I" Then PI.Visible = True
      If PI.Value = "C" Then PI.Visible = True
Next PI
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
For Each PI In PF.PivotItems 'reset business area (multiple items in page field)
    If PI.Value = "G400" Or PI.Value = "6000" Then
        PI.Visible = True
        Else
        PI.Visible = False
    End If
Next PI

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~no declaration and loops~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
PT.PivotFields("Description").PivotItems("Interface").Visible = True
PT.PivotFields("Source").PivotItems("S").Visible = True

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~with declaration; no loops~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Set PF = PT.PivotFields("Month")
PF.Orientation = xlPageField
PF.ClearAllFilters
PF.EnableMultiplePageItems = True       'allows more than 1 item to be selected in a page filter
PF.EnableItemSelection = False          'do not confuse with the previous one; When set to False, disables the ability to use the field dropdown in the user interface.
PF.IncludeNewItemsInFilter              'often used when a field has been manually filtered to exclude 1 or two predefined items to make sure that all newly added items will appaear in pivot items

With PF
  .PivotItems("2").Visible = False
  .PivotItems("3").Visible = False

  .PivotItems("3").Visible = True
  .PivotItems("4").Visible = True
  .PivotItems("5").Visible = True
  .PivotItems("6").Visible = True
End With

''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~resetting the page field for a single filter~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'filter a pivot field for latest month
Set PT = Worksheets("Utilization").PivotTables("MTD PL cluster")                'PT fed through Access connection, and all pTs share the same cache; no point to refresh them all
Set PF = PT.PivotFields("Month")

PF.ClearAllFilters
PF.CurrentPage = "DCC"                                                              'constant
PF.CurrentPage = Today                                                              'today declared as today's date or string to be put via InputBox
PF.CurrentPage = Format(Application.WorksheetFunction.EoMonth(Date, -1), "m")       'calendar month, hence -1; obtained from computer date
PF.CurrentPage = Format(Application.WorksheetFunction.EoMonth(Date, 1), "m")        'fiscal period, hence + 1; obtained from computer date
PF.CurrentPage = "Period " & Format(Application.WorksheetFunction.EoMonth(Date, 1), "mm") & Chr(32) & UTC.Range("H1") ' obtained from computer date + constant


Set PF = PT1.PivotFields("Date")                                                    'filter a pivot table for a variable pivotitem from the pivot table dataset
monthly = Format(Application.WorksheetFunction.EoMonth(Date, -2) + 1, "dd/mm/yyyy") 'returns the 1st day of a month in "dd/mmm/yyy" format

Set PI = PF.PivotItems(monthly)
PF.CurrentPage = PI                                                                 'filter a pivot table for a variable pivotitem from the pivot table dataset


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~working with slicers~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
With ActiveWorkbook.SlicerCaches("Slicer_Quarter1") 'manipulating selected items for all slicers sharing the same cache
        .SlicerItems("Q1").Selected = True
        .SlicerItems("Q2").Selected = False
        .SlicerItems("Q3").Selected = False
        .SlicerItems("Q4").Selected = False
End With
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Next PT
End With
Next Wks

End Sub
