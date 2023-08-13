Sub estrutura_repeticao()

For Each celula In Range("Q2:Q50")
    If (celula.Offset(0, -1).Value) / 10 = 0 Then
        celula.Value = 0
    ElseIf (celula.Offset(0, -1).Value + 1) Mod 10 = 0 Then
        celula.Value = (celula.Offset(0, -1).Value + 1) - celula.Offset(0, -1).Value
    ElseIf ((celula.Offset(0, -1).Value + 2) Mod 10) = 0 Then
        celula.Value = (celula.Offset(0, -1).Value + 2) - celula.Offset(0, -1).Value
    ElseIf ((celula.Offset(0, -1).Value + 3) Mod 10) = 0 Then
        celula.Value = (celula.Offset(0, -1).Value + 3) - celula.Offset(0, -1).Value
    ElseIf (celula.Offset(0, -1).Value + 4) Mod 10 = 0 Then
        celula.Value = (celula.Offset(0, -1).Value + 4) - celula.Offset(0, -1).Value
    ElseIf (celula.Offset(0, -1).Value + 5) Mod 10 = 0 Then
        celula.Value = (celula.Offset(0, -1).Value + 5) - celula.Offset(0, -1).Value
    ElseIf (celula.Offset(0, -1).Value + 6) Mod 10 = 0 Then
        celula.Value = (celula.Offset(0, -1).Value + 6) - celula.Offset(0, -1).Value
    ElseIf (celula.Offset(0, -1).Value + 7) Mod 10 = 0 Then
        celula.Value = (celula.Offset(0, -1).Value + 7) - celula.Offset(0, -1).Value
    ElseIf (celula.Offset(0, -1).Value + 8) Mod 10 = 0 Then
        celula.Value = (celula.Offset(0, -1).Value + 8) - celula.Offset(0, -1).Value
    ElseIf (celula.Offset(0, -1).Value + 9) Mod 10 = 0 Then
        celula.Value = (celula.Offset(0, -1).Value + 9) - celula.Offset(0, -1).Value
    Else
        celula.Value = "erro"
    End If
    
Next

End Sub
