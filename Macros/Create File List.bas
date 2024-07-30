Attribute VB_Name = "Module1"
Option Explicit

Sub ListFiles()
    Dim Directory As String
    Dim r As Long
    Dim f As String
    Dim FileSize As Double
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Sheets(1)

    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = Application.DefaultFilePath & "\"
        .Title = "Select a location containing the files you want to list."
        .Show
        If .SelectedItems.Count = 0 Then
            Exit Sub
        Else
            Directory = .SelectedItems(1) & "\"
        End If
    End With
    r = 1

    ws.Cells.ClearContents
    ws.Cells(r, 1) = "Files in " & Directory
    ws.Cells(r, 2) = "Size"
    ws.Cells(r, 3) = "Date/Time"
    ws.Range("A1:C1").Font.Bold = True

    f = Dir(Directory & "*.*")
    Do While f <> ""
        If Not (f = "." Or f = "..") Then ' Exclude current and parent directory entries
            r = r + 1
            ws.Cells(r, 1) = f
            FileSize = FileLen(Directory & f)
            ws.Cells(r, 2) = FileSize
            ws.Cells(r, 3) = FileDateTime(Directory & f)
        End If
        f = Dir
    Loop
End Sub


