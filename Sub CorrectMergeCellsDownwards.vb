Sub CorrectMergeCellsDownwards()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long, col As Integer

    ' Set active worksheet
    Set ws = ActiveSheet

    ' Find last used row in column B (assuming column B always has data)
    lastRow = ws.Cells(Rows.Count, 2).End(xlUp).Row

    ' Loop through every 4-row group
    For i = 2 To lastRow Step 4 ' Step 4 ensures we merge every 4 rows correctly
        For col = 2 To 7 ' Columns B to G (2 = B, 7 = G)
            ' Ensure there are enough rows left to merge
            If i + 3 <= lastRow Then
                With ws.Range(ws.Cells(i, col), ws.Cells(i + 3, col))
                    .Merge
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                End With
            End If
        Next col
    Next i

    ' Notify user
    MsgBox "Cells B to G merged down every 4 rows successfully!", vbInformation
End Sub

