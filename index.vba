Sub loopForNext()

For Each celula In Range("Q2:Q50")
Dim x As Long
    For x = 0 To 9
        If ((celula.Offset(0, -1).Value + x) Mod 10) = 0 Then
            celula.Value = (celula.Offset(0, -1).Value + x) - celula.Offset(0, -1).Value
        End If
            
    Next
Next
End Sub


End Sub
