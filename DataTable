Sub Dados_T()

Dim max As Variant
Dim i As Long

i = 2

Sheets("Produto_embarcado").Activate

max = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row

Range("A2" & ":A" & max).Select
Selection.Copy

Sheets("Dados_tratados").Activate
Range("A2").PasteSpecial xlPasteValues

Call preencher_form

Do While i <= max
    If Range("A" & i).Value = "TOTAL:Lotes:" Or _
        Range("A" & i).Value = "" Or _
        Range("A" & i).Value = "LoteCorrida" Or _
        Cells(i, 8).Value = "" Or Cells(i, 9).Value = "" Or Cells(i, 10).Value = "" Or _
        Range("A" & i).Value Like "AL: Lotes: 000*" Then

        Rows(i).Delete
        max = max - 1

    Else

        i = i + 1
    End If
    
Loop
    

End Sub
