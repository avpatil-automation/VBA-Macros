Sub SelectUnique()
    Dim response As Long
    Dim uniq_counter As Integer
    Dim cell As Range
    Dim uniques As Range
    Dim vals() As Variant

    'Selection does not need to be a single range, but it does need to be on the same sheet.
    If Selection.Count > 5000 Then
        response = MsgBox("This could take a while", vbOKCancel + vbInformation)
        If response = vbCancel Then Exit Sub
    End If

    ReDim vals(1 To Selection.Count)

    'Cycle through all values in selection
    For Each cell In Selection
        'Skip blank cells and errored cells
        If Not IsError(cell.Value2) And Not IsEmpty(cell) Then
            'Set first value
            If uniques Is Nothing Then
                Set uniques = cell
                vals(1) = cell.Value2
                uniq_counter = 2
            End If
            'Check each cell against previously set unique values
            Dim checker As Integer
            For checker = 1 To uniq_counter - 1
                If vals(checker) = cell.Value2 Then Exit For
                If checker = uniq_counter - 1 Then
                    Set uniques = Union(uniques, cell)
                    vals(uniq_counter) = cell.Value2
                    uniq_counter = uniq_counter + 1
                End If
            Next checker
        End If
    Next cell

    'Highlight unique range if it exists instead of selecting
    If Not uniques Is Nothing Then
        uniques.Interior.Color = RGB(255, 255, 0) 'Yellow color example for highlighting
    End If
End Sub
