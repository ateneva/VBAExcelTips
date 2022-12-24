Attribute VB_Name = "CellInR_AddRemoveHyperlinks"
Option Explicit

Sub AddRemoveHyperLinks()

Dim prv As String
Dim Cell As Range

For Each Cell In ActiveSheet.Range("W2:W" & ActiveSheet.UsedRange.Rows.Count)
    
    prv = Cell.Value
    Cell.Hyperlinks.Add Anchor:=Cell, Address:=prv                  'adds hyperlink to a website
    Cell.Hyperlinks.Add Anchor:=Cell, Address:="mailto:" & prv      'adds hyperlink to an e-mail

    Cell.Hyperlinks.Delete                                          'delete a hyperlink from a cell

Next Cell

End Sub
