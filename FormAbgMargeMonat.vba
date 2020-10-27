Sub FormAbgMargeMonat()
'
' FormAbgMargeMonat Makro
'

'
    Range("K2").Select
    ActiveCell.FormulaR1C1 = _
        "=[@[Preis pro Monat in €]]-([@[Preis pro Monat in €]]*[@[Wert Marge]]/100)"
    
    Range("L2").Select
    ActiveCell.FormulaR1C1 = _
        "=[@[Preis pro Jahr in €]]-([@[Preis pro Jahr in €]]*[@[Wert Marge]]/100)"
End Sub
