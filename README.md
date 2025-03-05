Skipping Unwanted Field Types

Sub InsertRowsAndTransposeData_Filtered()
    Dim ws As Worksheet
    Dim lastRow As Long, rowIndex As Long, numRows As Integer
    Dim colStart As Integer, colEnd As Integer
    Dim sourceRange As Range, targetCell As Range
    Dim arr As Variant
    Dim i As Integer, j As Integer, nonEmptyCount As Integer
    Dim firstRow As Range, cell As Range
    Dim regex As Object
    Dim fieldTypeCol As Integer
    Dim excludedTypes As Variant
    Dim shouldSkip As Boolean
    
    ' Set the active worksheet
    Set ws = ActiveSheet

    ' Find last used row dynamically
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' Find last used column dynamically
    colEnd = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' Define excluded field types
    excludedTypes = Array("text", "descriptive", "calc", "yesno", "notes")

    ' Identify the column containing "Field Type" (assuming it's in row 1)
    fieldTypeCol = 5 ' Adjust based on your column structure

    ' Create regex for detecting enumerated options (e.g., "1, Option")
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "^\d+,\s?.+" ' Looks for "1, Text" pattern
    regex.IgnoreCase = True
    regex.Global = False

    ' Find the first column that contains enumerated options and isn't an excluded type
    Set firstRow = ws.Range(ws.Cells(1, 1), ws.Cells(1, colEnd))
    For Each cell In firstRow
        shouldSkip = False
        
        ' Check if this column's field type should be skipped
        For Each excludedType In excludedTypes
            If LCase(Trim(ws.Cells(cell.Row + 1, fieldTypeCol).Value)) = excludedType Then
                shouldSkip = True
                Exit For
            End If
        Next excludedType

        ' If column should not be skipped and contains enumerated values, set colStart
        If Not shouldSkip And regex.Test(cell.Value) Then
            colStart = cell.Column
            Exit For
        End If
    Next cell

    ' If no valid column found, exit to prevent errors
    If colStart = 0 Then
        MsgBox "No valid enumerated options detected!", vbExclamation, "Error"
        Exit Sub
    End If

    ' Loop from bottom to top to prevent shifting issues
    For rowIndex = lastRow To 2 Step -1 ' Start at row 2, assuming headers in row 1
        ' Check if the row's field type is one we should process
        shouldSkip = False
        For Each excludedType In excludedTypes
            If LCase(Trim(ws.Cells(rowIndex, fieldTypeCol).Value)) = excludedType Then
                shouldSkip = True
                Exit For
            End If
        Next excludedType

        If shouldSkip Then GoTo NextRow

        ' Count non-empty cells in the row from colStart to colEnd
        nonEmptyCount = Application.WorksheetFunction.CountA(ws.Range(ws.Cells(rowIndex, colStart), ws.Cells(rowIndex, colEnd)))

        ' Calculate the number of blank rows to insert (nonEmptyCount - 1)
        numRows = nonEmptyCount - 1

        ' Insert blank rows only if multiple non-empty values exist
        If numRows > 0 Then
            ' Insert blank rows
            ws.Rows(rowIndex + 1 & ":" & rowIndex + numRows).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

            ' Select data from dynamically determined column range
            Set sourceRange = ws.Range(ws.Cells(rowIndex, colStart), ws.Cells(rowIndex, colEnd))

            ' Store data as an array
            arr = sourceRange.Value

            ' Transpose values into the inserted rows
            For i = 0 To numRows ' Start at 0 to ensure first value is included
                Set targetCell = ws.Cells(rowIndex + i, colStart) ' Start from first inserted row
                targetCell.Resize(1, UBound(arr, 2)).Value = Application.Transpose(Application.Index(arr, 1, i + 1))
            Next i
        End If

NextRow:
    Next rowIndex

    ' Cleanup
    Set regex = Nothing

    MsgBox "Processing complete!", vbInformation, "Success"
End Sub


---

skip columns labeled as "text", "descriptive", "calc", "yesno", and "notes", ensuring that it only processes "radio", "dropdown", and "checkbox" fields.


✅ Auto-detects the first column with enumerated values dynamically.
✅ Skips columns that have "text," "descriptive," "calc," "yesno," or "notes" in the "Field Type" column.
✅ Works for radio buttons, dropdowns, and checkboxes only.
✅ Prevents errors if no valid enumerated values are found.

# Excel
Excel formulas
