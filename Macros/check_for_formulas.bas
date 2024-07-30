Sub CheckForFormulas()
    'Purpose: Check for formulas in the active sheet
    Dim ws As Worksheet
    Dim cell As Range
    Dim hasFormulas As Boolean

    Set ws = ActiveSheet
    hasFormulas = False

    For Each cell In ws.Cells
        If cell.HasFormula Then
            hasFormulas = True
            Exit For
        End If
    Next cell

    If hasFormulas Then
        MsgBox "Formulas found on this sheet!"
    Else
        MsgBox "No Formulas Here!"
    End If
End Sub
