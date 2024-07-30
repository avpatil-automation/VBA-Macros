Sub FormatTopRow()
    ' Purpose: Freezes and formats the top row of the active sheet's table
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim toprow As Range
    Dim tbl As Range
    
    ' Check if there are any used cells in row 1
    If Application.WorksheetFunction.CountA(ws.Rows(1)) > 0 Then
        ' Determine the range from A1 to the last used cell in row 1
        Set toprow = ws.Range(ws.Cells(1, 1), ws.Cells(1, ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column))
    Else
        ' If there are no used cells in row 1, use the current selection's region
        On Error Resume Next
        Set tbl = Selection.CurrentRegion
        On Error GoTo 0
        
        If tbl Is Nothing Then
            MsgBox "Couldn't find a table to format! Click a cell in the table and run again", vbExclamation, "Couldn't find table!"
            Exit Sub
        End If
        
        ' Set the top row to the first row of the table
        Set toprow = tbl.Rows(1)
    End If
    
    ' Unfreeze panes before applying formatting
    ws.Cells(toprow.Row + 1, 1).Select
    ws.Cells(1, 1).Select ' Ensure we are not inside the top row before freezing
    ws.Activate
    ws.Application.Goto Reference:=ws.Cells(1, 1), Scroll:=True

    ' Freeze panes
    ws.Application.ActiveWindow.FreezePanes = False
    ws.Application.ActiveWindow.FreezePanes = True

    ' Format the top row
    With toprow.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    toprow.Font.Bold = True
    toprow.Font.Color = vbWhite
End Sub

