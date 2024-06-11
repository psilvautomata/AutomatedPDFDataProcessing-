Sub preencher_form()

Dim w As Variant
Dim max As Long

Sheets("Dados_Tratados").Activate

max = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    
Range("B2").FormulaR1C1 = _
    "=IFERROR(IF(XLOOKUP(RC[-1],Produto_Embarcado!C1,Produto_Embarcado!C2)="""","""",(XLOOKUP(RC[-1],Produto_Embarcado!C1,Produto_Embarcado!C2))),"""")"
Range("C2").FormulaR1C1 = _
    "=IFERROR(IF(XLOOKUP(RC[-2],Produto_Embarcado!C1,Produto_Embarcado!C3)="""","""",XLOOKUP(RC[-2],Produto_Embarcado!C1,Produto_Embarcado!C3)),"""")"
Range("D2").FormulaR1C1 = _
    "=IFERROR(IF(XLOOKUP(RC[-3],Produto_Embarcado!C1,Produto_Embarcado!C4)="""","""",XLOOKUP(RC[-3],Produto_Embarcado!C1,Produto_Embarcado!C4)),"""")"
Range("E2").FormulaR1C1 = _
    "=IFERROR(IF(XLOOKUP(RC[-4],Prop_Mec!C1,Prop_Mec!C4)="""","""",XLOOKUP(RC[-4],Prop_Mec!C1,Prop_Mec!C4)),"""")"
Range("F2").FormulaR1C1 = _
    "=IFERROR(IF(XLOOKUP(RC[-5],Prop_Mec!C1,Prop_Mec!C5)="""","""",XLOOKUP(RC[-5],Prop_Mec!C1,Prop_Mec!C5)),"""")"
Range("G2").FormulaR1C1 = _
    "=IFERROR(IF(XLOOKUP(RC[-6],Prop_Mec!C1,Prop_Mec!C7)="""","""",XLOOKUP(RC[-6],Prop_Mec!C1,Prop_Mec!C7)),"""")"
Range("H2").FormulaR1C1 = _
    "=IFERROR(IF(XLOOKUP(RC[-4],Comp_Quim!C1,Comp_Quim!C[-6])="""","""",(XLOOKUP(RC[-4],Comp_Quim!C1,Comp_Quim!C[-6]))),"""")"
Range("I2").FormulaR1C1 = _
    "=IFERROR(IF(XLOOKUP(RC[-5],Comp_Quim!C1,Comp_Quim!C[-3])="""","""",(XLOOKUP(RC[-5],Comp_Quim!C1,Comp_Quim!C[-3]))),"""")"
Range("J2").FormulaR1C1 = _
    "=IFERROR(IF(XLOOKUP(RC[-6],Comp_Quim!C1,Comp_Quim!C[-7])="""","""",(XLOOKUP(RC[-6],Comp_Quim!C1,Comp_Quim!C[-7]))),"""")"
Range("K2").FormulaR1C1 = _
    "=IFERROR(IF(XLOOKUP(RC[-7],Comp_Quim!C1,Comp_Quim!C[-7])="""","""",(XLOOKUP(RC[-7],Comp_Quim!C1,Comp_Quim!C[-7]))),"""")"
Range("L2").FormulaR1C1 = _
    "=IFERROR(IF(XLOOKUP(RC[-8],Comp_Quim!C1,Comp_Quim!C[-7])="""","""",(XLOOKUP(RC[-8],Comp_Quim!C1,Comp_Quim!C[-7]))),"""")"
Range("M2").FormulaR1C1 = _
    "=IFERROR(IF(XLOOKUP(RC[-9],Comp_Quim!C1,Comp_Quim!C[-1])="""","""",(XLOOKUP(RC[-9],Comp_Quim!C1,Comp_Quim!C[-1]))),"""")"
Range("N2").FormulaR1C1 = _
    "=IFERROR(IF(XLOOKUP(RC[-10],Comp_Quim!C1,Comp_Quim!C[-7])="""","""",(XLOOKUP(RC[-10],Comp_Quim!C1,Comp_Quim!C[-7]))),"""")"
Range("O2").FormulaR1C1 = _
    "=IFERROR(IF(XLOOKUP(RC[-11],Comp_Quim!C1,Comp_Quim!C[-1])="""","""",(XLOOKUP(RC[-11],Comp_Quim!C1,Comp_Quim!C[-1]))),"""")"
Range("P2").FormulaR1C1 = _
"=IFERROR(IF(XLOOKUP(RC[-12],Comp_Quim!C1,Comp_Quim!C[-1])="""","""",(XLOOKUP(RC[-12],Comp_Quim!C1,Comp_Quim!C[-1]))),"""")"
Range("Q2").FormulaR1C1 = _
"=IFERROR(IF(XLOOKUP(RC[-13],Comp_Quim!C1,Comp_Quim!C[-1])="""","""",(XLOOKUP(RC[-13],Comp_Quim!C1,Comp_Quim!C[-1]))),"""")"
Range("R2").FormulaR1C1 = _
"=IFERROR(IF(XLOOKUP(RC[-14],Comp_Quim!C1,Comp_Quim!C[-9])="""","""",(XLOOKUP(RC[-14],Comp_Quim!C1,Comp_Quim!C[-9]))),"""")"
Range("S2").FormulaR1C1 = _
"=IFERROR(IF(XLOOKUP(RC[-15],Comp_Quim!C1,Comp_Quim!C[-11])="""","""",(XLOOKUP(RC[-15],Comp_Quim!C1,Comp_Quim!C[-11]))),"""")"
Range("T2").FormulaR1C1 = _
"=IFERROR(IF(XLOOKUP(RC[-16],Comp_Quim!C1,Comp_Quim!C[-10])="""","""",(XLOOKUP(RC[-16],Comp_Quim!C1,Comp_Quim!C[-10]))),"""")"
Range("U2").FormulaR1C1 = _
"=IFERROR(IF(XLOOKUP(RC[-17],Comp_Quim!C1,Comp_Quim!C[-8])="""","""",(XLOOKUP(RC[-17],Comp_Quim!C1,Comp_Quim!C[-8]))),"""")"
Range("V2").FormulaR1C1 = _
    "=INDEX(Automate!C[-20], MATCH(""*"" & RC[-20] & ""*"", Automate!C[-18], 0))"

With Range("A1:V1").FormulaR1C1 = _
    Array("Lote ID", "Lote", "Corrida", "Placa", "LE", "LR", "AL%", "C", "Si", "Mn", "P", "S", "Al", "Cu", "Nb", "V", "Ti", "Cr", "Ni", "Mo", "N", "Aço")
End With

Range("A1:V1").Select

With Selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .ThemeColor = xlThemeColorDark1
    .TintAndShade = -0.349986266670736
    .PatternTintAndShade = 0
End With
With Selection.Font
    .ThemeColor = xlThemeColorDark1
    .TintAndShade = 0
End With
    
Range("B2:V2").AutoFill Destination:=Range("B2:V" & max), Type:=xlFillDefault

Range("A2:V" & max).Select

Range("A2:V" & max).Borders(xlEdgeTop).LineStyle = xlContinuous
Range("A2:V" & max).Borders(xlEdgeTop).ColorIndex = 0
Range("A2:V" & max).Borders(xlEdgeTop).TintAndShade = 0
Range("A2:V" & max).Borders(xlEdgeTop).Weight = xlThin

Range("A2:V" & max).Borders(xlEdgeBottom).LineStyle = xlContinuous
Range("A2:V" & max).Borders(xlEdgeBottom).ColorIndex = 0
Range("A2:V" & max).Borders(xlEdgeBottom).TintAndShade = 0
Range("A2:V" & max).Borders(xlEdgeBottom).Weight = xlThin

Range("A2:V" & max).Borders(xlEdgeLeft).LineStyle = xlContinuous
Range("A2:V" & max).Borders(xlEdgeLeft).ColorIndex = 0
Range("A2:V" & max).Borders(xlEdgeLeft).TintAndShade = 0
Range("A2:V" & max).Borders(xlEdgeLeft).Weight = xlThin

Range("A2:V" & max).Borders(xlEdgeRight).LineStyle = xlContinuous
Range("A2:V" & max).Borders(xlEdgeRight).ColorIndex = 0
Range("A2:V" & max).Borders(xlEdgeRight).TintAndShade = 0
Range("A2:V" & max).Borders(xlEdgeRight).Weight = xlThin

Range("A2:V" & max).Borders(xlInsideVertical).LineStyle = xlContinuous
Range("A2:V" & max).Borders(xlInsideVertical).ColorIndex = 0
Range("A2:V" & max).Borders(xlInsideVertical).TintAndShade = 0
Range("A2:V" & max).Borders(xlInsideVertical).Weight = xlThin

Range("A2:V" & max).Borders(xlInsideHorizontal).LineStyle = xlContinuous
Range("A2:V" & max).Borders(xlInsideHorizontal).ColorIndex = 0
Range("A2:V" & max).Borders(xlInsideHorizontal).TintAndShade = 0
Range("A2:V" & max).Borders(xlInsideHorizontal).Weight = xlThin

For w = 2 To max
    Sheets("Dados_Tratados").Activate
    If Range("F" & w).Value = "OK" Then 'If it is "OK", it doesn't refer to mechanical properties, but to hardness.
        Range("E" & w & ":G" & w).ClearContents
    End If
Next
        
End Sub
