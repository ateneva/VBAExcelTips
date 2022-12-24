Attribute VB_Name = "PTSourceData"
Option Explicit

Sub PivotCacheReport()

    Dim PC As PivotCache
    Dim S As String, sn As String
    Dim ws As Worksheet
    Dim PT As PivotTable
    
    With ActiveWorkbook
        For Each PC In .PivotCaches
            S = "Pivotcache " & CStr(PC.Index) & " uses " & CStr(PC.MemoryUsed) & " and has " _
                    & CStr(PC.RecordCount) & " records"
            S = S & Chr(10) & "The following pivot tables use it"
            
            For Each ws In .Worksheets
                sn = ws.name
                For Each PT In ws.PivotTables
                    If PT.CacheIndex = PC.Index Then

                        If Len(sn) > 0 Then
                            S = S & Chr(10) & sn & Chr(10)
                            sn = ""
                        End If
                        S = S & Replace(PT.name, "PivotTable", "PT") & ","
                    End If
                Next PT
            Next ws
            MsgBox (S)
        Next PC
        
        
        sn = Chr(10) & "Couldnt find the pivotcache for these pivot tables"
        S = ""
        For Each ws In .Worksheets
            For Each PT In ws.PivotTables
                If PT.CacheIndex < 1 Or PT.CacheIndex > .PivotCaches.Count Then
                    S = S & Chr(10) & ws.name & ":" & Replace(PT.name, "PivotTable", "PT")
                End If
            Next PT
        Next ws
        If (Len(S) > 0) Then
            MsgBox (sn & S)
        End If
    End With
    

End Sub
