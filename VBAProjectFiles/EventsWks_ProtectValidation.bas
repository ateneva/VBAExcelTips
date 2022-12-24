Attribute VB_Name = "EventsWks_ProtectValidation"
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
Dim VT As Long

If Target.Column = 5 Then

On Error Resume Next
'the range must be specified in fixed length
'as stating a whole range or column creates a constnat loop
VT = Range("Zali").Validation.Type
    
    If Err.Number <> 0 Then
        Application.Undo
        MsgBox "Your last operation was canceled." & _
        "It would have deleted data validation rules.", vbCritical
    End If
End If
End Sub
