Sub MergeAndTransposeDropdowns()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim startRow As Long
    Dim i As Long
    Dim col As Integer
    Dim currentPID As Variant
    Dim dropdownValues As String
    
    ' Set active sheet
    Set ws = ActiveSheet

    ' Find last row with data
    lastRow = ws.Cells(Rows.Count, 2).End(xlUp).Row ' Column B (pid)

    ' Start at first row
    startRow = 1
    currentPID = ws.Cells(startRow, 2).Value ' Column B (pid)

    ' Loop through rows
    For i = 2 To lastRow + 1
        ' Check if column B (pid) changes or reached last row
        If ws.Cells(i, 2).Value <> currentPID Or i > lastRow Then
            ' Merge C, D, E for the grouped rows
            For col = 3 To 5  ' Columns C, D, E
                ws.Range(ws.Cells(startRow, col), ws.Cells(i - 1, col)).Merge
                ws.Cells(startRow, col).HorizontalAlignment = xlCenter
                ws.Cells(startRow, col).VerticalAlignment = xlCenter
            Next col
            
            ' Combine dropdown values in Column G
            dropdownValues = ""
            For j = startRow To i - 1
                If ws.Cells(j, 7).Value <> "" Then ' Column G (dropdown values)
                    dropdownValues = dropdownValues & ws.Cells(j, 7).Value & ", "
                End If
            Next j
            
            ' Remove trailing comma
            If Len(dropdownValues) > 2 Then
                dropdownValues = Left(dropdownValues, Len(dropdownValues) - 2)
            End If
            
            ' Insert merged dropdown values in first row of group
            ws.Cells(startRow, 7).Value = dropdownValues

            ' Move to next group
            startRow = i
            currentPID = ws.Cells(i, 2).Value
        End If
    Next i
    
    ' Notify user
    MsgBox "Merging and transposition completed successfully!", vbInformation
End Sub

