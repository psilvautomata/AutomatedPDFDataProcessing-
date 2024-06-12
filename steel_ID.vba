Sub f_aco()

'Formula to steel infos

Sheets("Aço").Activate
Range("A1").Select
ActiveCell.FormulaR1C1 = "=IF(Automate!RC[1]="""","""",TEXTAFTER(Automate!RC[1],""Especificação ""))"
Range("A1").Select
Selection.AutoFill Destination:=Range("A1:A1000")
Range("B1").Select
ActiveCell.FormulaR1C1 = "=LEFT(RC[-1],12)"
Selection.AutoFill Destination:=Range("B1:B1000")
Range("A1").Select

End Sub
