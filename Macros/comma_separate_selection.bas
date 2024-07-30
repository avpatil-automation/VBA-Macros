Sub CommaSeparateSelection()
    'Purpose: Comma separates all cells in selection and outputs them to an unused adjacent cell
    'Current sheet only
    Dim outputcell As Range
    Dim apos As Boolean
    Dim cell As Range
    
    Set outputcell = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Offset(0, 1)
    
    'Wrap comma separated values in quotes yes/no
    apos = MsgBox("Add apostrophes?", vbYesNo, "Add apostrophes and wrap selections in quotes?")
    
    If apos = vbYes Then
        apos = True
    Else
        apos = False
    End If
    
    If Not IsEmpty(Selection) Then
        For Each cell In Selection
            If cell.Value <> "" Then
                If apos = False Then
                    outputcell.Value = outputcell.Value & cell.Value & ", "
                Else
                    outputcell.Value = outputcell.Value & "'" & cell.Value & "', "
                End If
            End If
        Next cell
        
        'Removes trailing comma
        outputcell.Value = Left(outputcell.Value, Len(outputcell.Value) - 2)
    Else
        outputcell.Value = "No selection to process"
    End If
End Sub
