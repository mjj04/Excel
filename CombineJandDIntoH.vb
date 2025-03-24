VBA Macro: Merge Text from J and D into H
vba
Copy
Edit
Sub CombineJandDIntoH()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim textJ As String
    Dim textD As String
    
    ' Set the active worksheet
    Set ws = ActiveSheet

    ' Find the last row with data in Column J or D
    lastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row ' Column J

    ' Loop through each row and combine J and D into H
    For i = 1 To lastRow
        ' Get text from J and D (handling empty cells)
        textJ = ws.Cells(i, 10).Value ' Column J
        textD = ws.Cells(i, 4).Value  ' Column D

        ' Combine text with a space separator
        ws.Cells(i, 8).Value = Trim(textJ & " " & textD) ' Paste into Column H
    Next i

    ' Notify user
    MsgBox "Text from J and D has been combined into H successfully!", vbInformation
End Sub