Sub TopLeftActiveCell()
    'Purpose: Sets active cell to top left ($A$1) for all sheets
    Dim currsheet As Worksheet
    Dim sheet As Worksheet
    Set currsheet = ActiveSheet
    'Change A1 to suit your preference
    Const TopLeft As String = "A1"
    
    'Loop through all the sheets in the workbook
    For Each sheet In ThisWorkbook.Worksheets
        'Only does this for visible worksheets by using Excel object qualifier
        If sheet.Visible = Excel.xlSheetVisibility.xlSheetVisible Then
            With Application
                .Goto sheet.Range(TopLeft), Scroll:=True
            End With
        End If
    Next sheet

    currsheet.Activate
End Sub
