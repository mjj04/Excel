Sub InsertBlankRowsBelowEachRow()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Set active worksheet
    Set ws = ActiveSheet

    ' Find last used row in Column A
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' Loop from bottom to top to prevent row shifting issues
    For i = lastRow To 1 Step -1
        ws.Rows(i + 1 & ":" & i + 3).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Next i

    ' Notify user
    MsgBox "3 blank rows inserted below each row successfully!", vbInformation
End Sub

