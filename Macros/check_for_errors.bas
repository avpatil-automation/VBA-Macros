Sub CheckForNAs()
'Purpose: Check for #N/As in the current sheet
On Error GoTo err
    Dim errorRange As Range
    Set errorRange = Cells.SpecialCells(xlCellTypeFormulas, xlErrors)
    
    If Not errorRange Is Nothing Then
        errorRange.Select
    Else
        MsgBox "No Errors Here!"
    End If
    Exit Sub

err:
    MsgBox "An error occurred: " & Err.Description
End Sub
