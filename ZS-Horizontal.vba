Sub ZS-Horizontal()
'Dieses Macro soll gleiche AB Pos. finden und die Zahlschritte Fakturieren
    
    'Die Nummer aller Zeilen (nützlich für die For schleifen)
    FinalRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Debug.Print "FinalRow: "; FinalRow
    
    Dim AnzZS As Integer, temp() As String, a As Integer
    
    
    For i = 4 To FinalRow
        Debug.Print Cells(i, 1).Value
        
'        If Cells(i, 1).Value <> temp() Then
'            'temp(i) = Cells(i, 1).Value
'
'        End If
    Next i
    
    
End Sub
