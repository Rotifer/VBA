' Name all the cells of a sorted column using the unique values
'  of the starting cell column.
' Used to name blocks of ranges.
Sub AssignNamesToRange(rngForNames As Range, startCell As Range)
    Dim rngRowCount As Long: rngRowCount = rngForNames.Rows.Count
    Dim nextRowNum As Long
    Dim rngStart As Long
    Dim i As Long
    Dim j As Long
    Do While i <= rngRowCount
        j = 0
        Do While startCell.Offset(j, 0).Value = startCell.Value
            j = j + 1
        Loop
        If startCell.Offset(j, 0).Value = "" Then GoTo Skip
        Range(startCell, startCell.Offset(j - 1, 0)).Name = startCell.Value
        If Len(startCell.Offset(j, 0).Offset(1, 0)) = 0 Then
            startCell.Offset(j, 0).Name = startCell.Offset(j, 0).Value
        End If
        Set startCell = startCell.Offset(j - 1 + 1, 0)
Skip:
        i = i + j
    Loop
End Sub

Sub RunAssignNamesToRange()
    Dim rngForNames As Range: Set rngForNames = ActiveWorkbook.Worksheets("Sheet1").UsedRange
    Dim startCell As Range: Set startCell = ActiveWorkbook.Worksheets("Sheet1").Range("B1")
    Call AssignNamesToRange(rngForNames, startCell)
End Sub


Sub ClearNames()
    Dim name_ As Name
    For Each name_ In ActiveWorkbook.Names
        name_.Delete
    Next name_
End Sub
