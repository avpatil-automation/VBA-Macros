Sub FilterBySelection()
    Dim rng As Range
    Dim filterColumn As Long
    Dim cellValue As Variant
    
    ' Check if a single cell is selected
    If Selection.Cells.Count > 1 Then
        MsgBox "Please select a single cell to apply the filter.", vbExclamation
        Exit Sub
    End If
    
    ' Determine the range for autofilter
    On Error Resume Next
    Set rng = ActiveSheet.AutoFilter.Range
    On Error GoTo 0
    
    If rng Is Nothing Then
        MsgBox "No AutoFilter range found. Please select the data range and apply AutoFilter first.", vbExclamation
        Exit Sub
    End If
    
    ' Calculate the filter column dynamically
    filterColumn = Selection.Column - rng.Cells(1).Column + 1
    
    ' Handle error cells
    If IsError(Selection.Value) Then
        MsgBox "Cannot filter due to error value in selected cell.", vbExclamation
        Exit Sub
    Else
        cellValue = Selection.Value
    End If
    
    ' Apply filter
    rng.AutoFilter Field:=filterColumn, Criteria1:="=" & cellValue
End Sub
