Sub FormatierungUebertragen()
'
' FormatierungUebertragen Makro
'

'
    Sheets("HilfsTab").Range("E1").Copy
'    Selection.Copy
'    Sheets("ACS Armoured Car Systems GmbH").Select
    Range(ActiveCell.ListObject.Name & "[Wartungsart]").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("G3").Select
End Sub
