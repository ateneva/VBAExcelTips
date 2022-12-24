Attribute VB_Name = "PTAddCalculatedField"
Option Explicit

Sub AddCalcs()

Dim Wks As Worksheet
Dim PC As PivotCache
Dim PT As PivotTable
Dim PF As PivotField
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

ActiveWorkbook.Worksheets("vertica").Activate
        
    For Each PT In ActiveSheet.PivotTables
    
        'add the CTR, and RPM
        PT.CalculatedFields.Add "PaidCTR", "=PaidClicks/PaidPVs", True
        PT.CalculatedFields.Add "PaidListingCTR", "=PaidClicks/PaidListings", True
        PT.CalculatedFields.Add "PageRPM", "=GrossRevenue/PaidPVs*1000", True
        PT.CalculatedFields.Add "ListingRPM", "=GrossRevenue/PaidListings*1000", True
        PT.CalculatedFields.Add "PaidListingPerPage", "=PaidListings/PaidPVs", True
        PT.CalculatedFields.Add "CPC", "=GrossRevenue/PaidClicks", True
        PT.CalculatedFields.Add "Gross Revenue % Change", "=GrossRevenue", True
        PT.CalculatedFields.Add "Paid PVs % Change", "=PaidPVs", True
        PT.CalculatedFields.Add "CPC % Change", "=CPC", True
        PT.CalculatedFields.Add "Paid CTR % Change", "=PaidCTR", True
        PT.CalculatedFields.Add "Page RPM % Change", "=PageRPM", True
    
    Next PT
      
        
End Sub
