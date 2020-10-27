Sub ErstelleDropdown()
'
' ErstelleDropdown Makro
'

    
    Range(ActiveCell.ListObject.Name & "[Wartungsart]").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=HilfsTab!$A$2:$A$15"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
End Sub
