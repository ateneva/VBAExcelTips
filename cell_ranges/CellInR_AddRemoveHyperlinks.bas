
Sub AddRemoveHyperLinks()

Dim prv As String
Dim Cell As Range
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'written by Angelina Teneva
'https://datageeking.wordpress.com/2017/06/05/how-to-addremove-links-with-vba/
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For Each Cell In ActiveSheet.Range("W2:W" & ActiveSheet.UsedRange.Rows.Count)

    If Cell.Hyperlinks.Count = 0 Then

        prv = Cell.Value

        If InStr(1, Cell, "@", vbTextCompare) <> 0 Then
            Cell.Hyperlinks.Add Anchor:=Cell, Address:="mailto:" & prv      'adds hyperlink to an e-mail
        Else
            Cell.Hyperlinks.Add Anchor:=Cell, Address:=prv                  'adds hyperlink to a website
        End If
    Else

        Cell.Hyperlinks.Delete                                              'delete a hyperlink from a cell

    End If

Next Cell
End Sub
