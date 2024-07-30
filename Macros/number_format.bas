Sub NumberFormat()
    ' Check if there is a selection
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range of cells first.", vbExclamation
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Set ws = ActiveSheet ' You may need to specify a specific worksheet here

    ' Format selected cells
    With ws.Range(Selection.Address)
        .NumberFormat = "0" ' Formats as integer (removes decimals)
        .HorizontalAlignment = xlCenter
    End With
End Sub
