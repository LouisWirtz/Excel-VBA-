Sub paste_rename_col()
'Dieses Makro soll Spalten in einer Tabelle aus einem Register einfügen und diese benennen

'Füge Spalten hinzu
    Selection.ListObject.ListColumns.Add Position:=7
    Selection.ListObject.ListColumns.Add Position:=8
    Selection.ListObject.ListColumns.Add Position:=11
    Selection.ListObject.ListColumns.Add Position:=12
    
'Benenne die Spalten
'(Da die Spalten immer an der selben stelle sind haben sie absolute Verweise)
    Range("G1").Value = "Wartungsart"
    Range("H1").Value = "Wert Marge"
    
    Range("K1").Value = "Abg Marge pro Monat"
    Range("L1").Value = "Abg Marge pro Jahr"

End Sub
