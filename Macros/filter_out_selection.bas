Sub FilterOutSelection()
    Dim ws As Worksheet
    Dim cell As Range
    Dim filtercolumn As Integer
    Dim cellvalue As Variant
    
    Set ws = ActiveSheet
    
    If TypeName(Selection) = "Range" Then
        If Selection.Count > 1 Then
            Selection.Cells(1, 1).Select
        End If
        
        If Not ws.AutoFilterMode Then
            ws.AutoFilterMode = True
        End If
        
        If Not Intersect(Selection, ws.AutoFilter.Range) Is Nothing Then
            filtercolumn = Selection.Cells(1, 1).Column - ws.AutoFilter.Range.Columns(1).Column + 1
            
            For Each cell In Selection
                If IsError(cell.Value) Then
                    cellvalue = cell.Text
                Else
                    cellvalue = cell.Value
                End If
                
                ws.AutoFilter.Filters(filtercolumn).Criteria1 = "<>""" & cellvalue & """"
            Next cell
        Else
            MsgBox "Selection is not within the AutoFilter range.", vbExclamation
        End If
    Else
        MsgBox "Please select a range of cells.", vbExclamation
    End If
End Sub
