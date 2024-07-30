Sub BetterAutoFilter()
    ' Purpose: One button that turns on autofilter (when off), clear the filter (when filtered), or shuts autofilter (when on and not filtered)
    
    ' Check if AutoFilter is applied
    On Error Resume Next

    If ActiveSheet.AutoFilterMode = False Then
        ' Autofilter is off, turning it on
        ActiveSheet.Range("A1").AutoFilter
    ElseIf ActiveSheet.FilterMode = True Then
        ' Autofilter is on and data is filtered, clearing the filter
        ActiveSheet.ShowAllData
    Else
        ' Autofilter is on but no filter is applied, turning it off
        ActiveSheet.AutoFilterMode = False
    End If
End Sub


